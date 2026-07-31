"""Business-safe missing-value classification, cleaning, and audit utilities."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from enum import Enum
from typing import Any, Iterable, Mapping, Sequence
from uuid import uuid4

import pandas as pd

from src.integrity import (
    IntegrityReport,
    RelationshipRule,
    ValidatedRelationship,
    detect_validated_relationships,
    recover_deterministic_value,
    relationship_columns,
    validate_integrity,
)

SOURCE_FILE_COLUMN = "source_file"
SourceSchemas = dict[str, frozenset[str]]

STRATEGY_LABELS = {
    "fill_mean": "Filled with the average",
    "fill_median": "Filled with the median",
    "fill_mode": "Filled with the most common value",
    "fill_zero": "Filled with zero",
    "drop_rows": "Removed incomplete rows",
    "leave_blank": "Left blank",
    "recover_relationship": "Recovered from a validated relationship",
}


class MissingDecisionState(str, Enum):
    """Explicit lifecycle state for one missing-value decision."""

    PENDING_REVIEW = "pending_review"
    CHANGED = "changed"
    APPROVED_UNCHANGED = "approved_unchanged"
    UNAVAILABLE_FROM_SOURCE = "unavailable_from_source"
    FAILED_OR_UNRESOLVED = "failed_or_unresolved"


@dataclass(frozen=True)
class MissingDecisionSummary:
    """Computed counts for the mutually exclusive decision states."""

    changed: int
    approved_unchanged: int
    unavailable_from_source: int
    pending_review: int
    failed_or_unresolved: int
    incomplete_rows_removed: int = 0

    @property
    def reviewed(self) -> int:
        return (
            self.changed
            + self.approved_unchanged
            + self.incomplete_rows_removed
        )

    @property
    def complete(self) -> bool:
        return self.pending_review == 0 and self.failed_or_unresolved == 0


@dataclass(frozen=True)
class ReviewSummary:
    """Authoritative counts used by every cleaning-completion component."""

    recovered_or_filled_count: int
    approved_unchanged_count: int
    unavailable_from_source_count: int
    pending_decision_count: int
    failed_or_unresolved_count: int
    duplicate_rows_removed: int
    integrity_status: str

    @property
    def complete(self) -> bool:
        return (
            self.pending_decision_count == 0
            and self.failed_or_unresolved_count == 0
            and self.integrity_status == "Passed"
        )

    def activity_lines(self) -> tuple[str, ...]:
        """Return stable completion copy without inspecting physical nulls."""
        return (
            f"Removed {self.duplicate_rows_removed:,} repeated "
            f"{'row' if self.duplicate_rows_removed == 1 else 'rows'}",
            f"Recovered {self.recovered_or_filled_count:,} missing "
            f"{'value' if self.recovered_or_filled_count == 1 else 'values'}",
            f"Approved {self.approved_unchanged_count:,} "
            f"{'value' if self.approved_unchanged_count == 1 else 'values'} "
            "to remain blank",
            f"{self.pending_decision_count:,} "
            f"{'decision' if self.pending_decision_count == 1 else 'decisions'} "
            "pending",
            f"Integrity validation {self.integrity_status.lower()}",
        )


@dataclass(frozen=True)
class MissingColumnStatus:
    """Decision-aware blank counts for one field."""

    approved_blank: int
    unavailable_from_source: int
    decisions_pending: int
    failed_or_unresolved: int = 0


@dataclass(frozen=True)
class CleaningAuditEntry:
    """One traceable cleaning decision or data-quality condition."""

    action_type: str
    action: str
    column: str
    source_file: str
    row_identifier: str
    row_index: str
    original_state: str
    resulting_value: Any
    strategy: str
    strategy_scope: str
    missing_type: str
    rows_removed: int
    reason: str
    recorded_at: str
    audit_event_id: str = field(
        default_factory=lambda: f"AUD-{uuid4().hex.upper()}"
    )
    original_source_row: str = ""
    business_record_identifier: str = ""
    original_value: Any = None
    formula_or_strategy: str = ""
    calculation_scope: str = ""
    timestamp: str = ""
    decision_state: MissingDecisionState | str = ""
    affected_row_count: int = 1
    affected_record_identifiers: str = ""
    resulting_state: str = ""

    def as_record(self) -> dict[str, Any]:
        """Return a serialization-friendly audit record."""
        record = asdict(self)
        record["original_source_row"] = (
            self.original_source_row or self.row_index
        )
        record["business_record_identifier"] = (
            self.business_record_identifier or self.row_identifier
        )
        record["original_value"] = (
            self.original_value
            if self.original_value is not None
            else self.original_state
        )
        record["formula_or_strategy"] = (
            self.formula_or_strategy or self.strategy
        )
        record["calculation_scope"] = (
            self.calculation_scope or self.strategy_scope
        )
        record["timestamp"] = self.timestamp or self.recorded_at
        record["decision_state"] = (
            self.decision_state.value
            if isinstance(self.decision_state, MissingDecisionState)
            else self.decision_state
        )
        return record


@dataclass(frozen=True)
class MissingValueResult:
    """Result of applying user-approved missing-value strategies."""

    cleaned: pd.DataFrame
    audit: tuple[CleaningAuditEntry, ...]
    messages: tuple[str, ...]
    integrity_report: IntegrityReport
    decision_summary: MissingDecisionSummary


def build_source_schemas(
    datasets: Iterable[tuple[str, pd.DataFrame]],
) -> SourceSchemas:
    """Capture the original columns supplied by each source file."""
    return {
        str(name): frozenset(str(column) for column in frame.columns)
        for name, frame in datasets
    }


def _schema_for_source(
    source: object,
    source_schemas: Mapping[str, Iterable[str]] | None,
) -> set[str] | None:
    if not source_schemas:
        return None
    schema = source_schemas.get(str(source))
    return set(schema) if schema is not None else None


def structural_missing_mask(
    df: pd.DataFrame,
    column: str,
    source_schemas: Mapping[str, Iterable[str]] | None,
) -> pd.Series:
    """Return cells blank because their source never supplied ``column``."""
    mask = pd.Series(False, index=df.index, dtype=bool)
    if source_schemas is None:
        source_schemas = df.attrs.get("source_schemas")
    if (
        column not in df.columns
        or SOURCE_FILE_COLUMN not in df.columns
        or not source_schemas
    ):
        return mask

    missing = df[column].isna()
    for source, indices in df.groupby(SOURCE_FILE_COLUMN, dropna=False).groups.items():
        schema = _schema_for_source(source, source_schemas)
        if schema is not None and column not in schema:
            mask.loc[list(indices)] = missing.loc[list(indices)]
    return mask


def row_level_missing_mask(
    df: pd.DataFrame,
    column: str,
    source_schemas: Mapping[str, Iterable[str]] | None,
) -> pd.Series:
    """Return blanks where the source supplied the column but the row did not."""
    if column not in df.columns:
        return pd.Series(False, index=df.index, dtype=bool)
    return df[column].isna() & ~structural_missing_mask(
        df,
        column,
        source_schemas,
    )


def classify_missing_values(
    df: pd.DataFrame,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
) -> pd.DataFrame:
    """Summarize structural and genuine row-level blanks by column."""
    columns = [
        "column",
        "missing_count",
        "missing_pct",
        "structural_count",
        "row_level_count",
        "dtype",
        "structural_by_source",
        "row_level_by_source",
    ]
    records: list[dict[str, Any]] = []
    row_count = len(df)

    for column_object in df.columns:
        column = str(column_object)
        missing = df[column_object].isna()
        structural = structural_missing_mask(df, column, source_schemas)
        row_level = missing & ~structural
        structural_by_source: dict[str, int] = {}
        row_level_by_source: dict[str, int] = {}

        if SOURCE_FILE_COLUMN in df.columns:
            for source in df[SOURCE_FILE_COLUMN].drop_duplicates().tolist():
                source_mask = df[SOURCE_FILE_COLUMN].eq(source)
                structural_count = int((structural & source_mask).sum())
                row_level_count = int((row_level & source_mask).sum())
                if structural_count:
                    structural_by_source[str(source)] = structural_count
                if row_level_count:
                    row_level_by_source[str(source)] = row_level_count

        missing_count = int(missing.sum())
        records.append({
            "column": column_object,
            "missing_count": missing_count,
            "missing_pct": round(
                (missing_count / row_count * 100) if row_count else 0.0,
                1,
            ),
            "structural_count": int(structural.sum()),
            "row_level_count": int(row_level.sum()),
            "dtype": str(df[column_object].dtype),
            "structural_by_source": structural_by_source,
            "row_level_by_source": row_level_by_source,
        })

    result = pd.DataFrame(records, columns=columns)
    if result.empty:
        return result
    return result.sort_values(
        ["row_level_count", "structural_count", "missing_count"],
        ascending=False,
    ).reset_index(drop=True)


def recommend_row_level_strategy(
    df: pd.DataFrame,
    column: str,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    *,
    configured_relationships: Sequence[RelationshipRule] | None = None,
    maximum_numeric_missing_rate: float = 0.20,
) -> str:
    """Return a business-safe recommendation for genuine row-level blanks."""
    mask = row_level_missing_mask(df, column, source_schemas)
    if not mask.any():
        return "leave_blank"

    available = df.loc[~structural_missing_mask(df, column, source_schemas), column]
    valid = available.dropna()
    if valid.empty:
        return "leave_blank"

    relationships = detect_validated_relationships(
        df,
        configured_relationships,
    )
    if any(
        recover_deterministic_value(df, index, column, relationships)
        is not None
        for index in df.index[mask]
    ):
        return "recover_relationship"

    normalized = str(column).strip().lower().replace("-", "_").replace(" ", "_")
    safe_blank_tokens = {
        "identifier",
        "code",
        "discount",
        "promotion",
        "promo",
        "city",
        "region",
        "account",
        "status",
        "sku",
        "product_code",
    }
    if (
        normalized == "id"
        or normalized.endswith("_id")
        or any(token in normalized for token in safe_blank_tokens)
    ):
        return "leave_blank"

    if pd.api.types.is_numeric_dtype(df[column]):
        if column in relationship_columns(relationships):
            return "leave_blank"
        if float(mask.mean()) > maximum_numeric_missing_rate:
            return "leave_blank"
        return "fill_median"
    return "leave_blank"


def recommendation_explanation(strategy: str) -> tuple[str, str]:
    """Return safe client-facing recommendation copy for a strategy."""
    if strategy == "recover_relationship":
        return (
            "Recommended: Recover from Total ÷ Unit Price",
            "The relationship was validated across the complete records, so "
            "the missing value can be reconstructed without estimation.",
        )
    if strategy == "leave_blank":
        return (
            "Recommended: Leave blank",
            "No reliable value can be derived from the available data. Filling "
            "this field would create unverified information.",
        )
    if strategy == "fill_median":
        return (
            "Recommended: Use the middle value",
            "This statistical replacement is available for a numeric field, "
            "but it remains an estimate.",
        )
    return (
        f"Recommended: {STRATEGY_LABELS.get(strategy, strategy)}",
        "Review this action before applying it.",
    )


def _identifier_column(df: pd.DataFrame) -> str | None:
    candidates = [
        column
        for column in df.columns
        if str(column).lower() == "id"
        or str(column).lower().endswith("_id")
    ]
    return str(candidates[0]) if candidates else None


def row_identifier(df: pd.DataFrame, index: object) -> str:
    """Return a business identifier where possible, otherwise a stable row label."""
    identifier_column = _identifier_column(df)
    if identifier_column and index in df.index:
        value = df.at[index, identifier_column]
        if not pd.isna(value):
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value)
    return f"Row {index}"


def source_row_identifier(df: pd.DataFrame, index: object) -> str:
    """Return the original one-based source data row where available."""
    source_rows = df.attrs.get("source_row_numbers", {})
    value = source_rows.get(index)
    return str(value) if value is not None else str(index)


def _display_value(value: Any) -> str:
    if value is None:
        return "—"
    try:
        if bool(pd.isna(value)):
            return "—"
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _strategy_label(strategy: str) -> str:
    if strategy.startswith("fill_value:"):
        return "Filled with a user-provided value"
    return STRATEGY_LABELS.get(strategy, "Applied selected cleaning action")


def _replacement_phrase(strategy: str) -> str:
    if strategy.startswith("fill_value:"):
        return "a user-provided value"
    return {
        "fill_mean": "the average",
        "fill_median": "the median",
        "fill_mode": "the most common value",
        "fill_zero": "zero",
    }.get(strategy, "the selected strategy")


def _calculated_value(series: pd.Series, strategy: str) -> Any:
    valid = series.dropna()
    if valid.empty:
        return None
    if strategy == "fill_mean":
        return valid.mean()
    if strategy == "fill_median":
        return valid.median()
    if strategy == "fill_mode":
        mode = valid.mode(dropna=True)
        return mode.iloc[0] if not mode.empty else None
    return None


def _replacement_for_row(
    df: pd.DataFrame,
    column: str,
    index: object,
    strategy: str,
    business_group_columns: Sequence[str] | None,
) -> tuple[Any, str]:
    valid_mask = df[column].notna()

    if SOURCE_FILE_COLUMN in df.columns:
        source = df.at[index, SOURCE_FILE_COLUMN]
        same_source = df[SOURCE_FILE_COLUMN].eq(source) & valid_mask
        value = _calculated_value(df.loc[same_source, column], strategy)
        if value is not None and not pd.isna(value):
            return value, f"Within source file: {source}"

    configured_groups = [
        group
        for group in (business_group_columns or [])
        if group in df.columns and not pd.isna(df.at[index, group])
    ]
    if configured_groups:
        group_mask = valid_mask.copy()
        for group in configured_groups:
            group_mask &= df[group].eq(df.at[index, group])
        value = _calculated_value(df.loc[group_mask, column], strategy)
        if value is not None and not pd.isna(value):
            return value, (
                "Within configured group: "
                + ", ".join(str(group) for group in configured_groups)
            )

    value = _calculated_value(df.loc[valid_mask, column], strategy)
    if value is not None and not pd.isna(value):
        return value, "Global fallback"
    return None, "No replacement available"


def _group_indices_by_source(
    df: pd.DataFrame,
    indices: Sequence[object],
) -> list[tuple[str, list[object]]]:
    """Group selected row indices by their original source file."""
    if not indices:
        return []
    if SOURCE_FILE_COLUMN not in df.columns:
        return [("—", list(indices))]
    groups: dict[str, list[object]] = {}
    for index in indices:
        source = _display_value(df.at[index, SOURCE_FILE_COLUMN])
        groups.setdefault(source, []).append(index)
    return list(groups.items())


def _aggregate_blank_decision(
    df: pd.DataFrame,
    column: str,
    indices: Sequence[object],
    *,
    decision_state: MissingDecisionState,
    action: str,
    resulting_state: str,
    reason: str,
    strategy: str,
    recorded_at: str,
) -> list[CleaningAuditEntry]:
    """Create concise per-source audit entries for unchanged/failed blanks."""
    entries: list[CleaningAuditEntry] = []
    for source, source_indices in _group_indices_by_source(df, indices):
        identifiers = [row_identifier(df, index) for index in source_indices]
        source_rows = [
            source_row_identifier(df, index)
            for index in source_indices
        ]
        count = len(source_indices)
        entries.append(CleaningAuditEntry(
            action_type=decision_state.value,
            action=action,
            column=str(column),
            source_file=source,
            row_identifier=f"{count:,} reviewed records",
            row_index="Aggregated",
            original_state=(
                "Source contained the field, but the record had no value"
            ),
            resulting_value=None,
            strategy=strategy,
            strategy_scope=f"Source file and field: {source} / {column}",
            missing_type="row-level",
            rows_removed=0,
            reason=reason,
            recorded_at=recorded_at,
            original_source_row=", ".join(source_rows),
            business_record_identifier=", ".join(identifiers),
            original_value="Blank",
            formula_or_strategy=strategy,
            calculation_scope=f"Source file and field: {source} / {column}",
            timestamp=recorded_at,
            decision_state=decision_state,
            affected_row_count=count,
            affected_record_identifiers=", ".join(identifiers),
            resulting_state=resulting_state,
        ))
    return entries


def apply_missing_value_strategies(
    df: pd.DataFrame,
    strategies: Mapping[str, str],
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    *,
    business_group_columns: Sequence[str] | None = None,
    allow_structural_override: bool = False,
    configured_relationships: Sequence[RelationshipRule] | None = None,
    recorded_at: str | None = None,
) -> MissingValueResult:
    """Apply strategies only to eligible blanks and return a detailed audit.

    Structural blanks are excluded unless ``allow_structural_override`` is
    explicitly enabled. Calculated fills use the same source file first, then
    configured business groups, then a documented global fallback.
    """
    cleaned = df.copy(deep=True)
    audit: list[CleaningAuditEntry] = []
    messages: list[str] = []
    recorded_at = recorded_at or datetime.now(timezone.utc).isoformat(
        timespec="seconds"
    )
    trusted_relationships = detect_validated_relationships(
        df,
        configured_relationships,
    )
    trusted_columns = relationship_columns(trusted_relationships)

    fill_strategies = {
        column: strategy
        for column, strategy in strategies.items()
        if strategy not in {"drop_rows", "leave_blank"}
    }
    drop_columns = [
        column
        for column, strategy in strategies.items()
        if strategy == "drop_rows"
    ]
    leave_blank_columns = [
        column
        for column, strategy in strategies.items()
        if strategy == "leave_blank"
    ]
    attempted_fill_indices: dict[str, list[object]] = {}

    for column, strategy in fill_strategies.items():
        if column not in cleaned.columns:
            continue
        structural = structural_missing_mask(cleaned, column, source_schemas)
        eligible = cleaned[column].isna()
        if not allow_structural_override:
            eligible &= ~structural
        eligible_indices = list(cleaned.index[eligible])
        attempted_fill_indices[str(column)] = eligible_indices
        filled_count = 0
        scopes: set[str] = set()

        for index in eligible_indices:
            source = (
                _display_value(cleaned.at[index, SOURCE_FILE_COLUMN])
                if SOURCE_FILE_COLUMN in cleaned.columns
                else "—"
            )
            deterministic = recover_deterministic_value(
                df,
                index,
                str(column),
                trusted_relationships,
            )
            used_deterministic_recovery = deterministic is not None
            formula = ""
            relationship_label = ""
            inputs: Mapping[str, Any] = {}
            if deterministic is not None:
                replacement, formula, relationship_label, inputs = deterministic
                scope = f"Validated relationship: {relationship_label}"
            elif strategy == "recover_relationship":
                continue
            elif (
                strategy in {"fill_mean", "fill_median", "fill_mode"}
                and column in trusted_columns
            ):
                # Do not introduce an estimate that would break a trusted rule.
                continue
            elif strategy == "fill_zero":
                replacement, scope = 0, "User-selected constant"
            elif strategy.startswith("fill_value:"):
                replacement = strategy.split(":", 1)[1]
                scope = "User-provided value"
            else:
                replacement, scope = _replacement_for_row(
                    df,
                    column,
                    index,
                    strategy,
                    business_group_columns,
                )

            if replacement is None or pd.isna(replacement):
                continue

            missing_type = "structural" if bool(structural.loc[index]) else "row-level"
            cleaned.at[index, column] = replacement
            filled_count += 1
            scopes.add(scope)
            action_type = (
                "deterministic_recovery"
                if used_deterministic_recovery
                else "fill_blank"
            )
            action = (
                "Recovered from "
                + formula.split("=", 1)[1].strip()
                .replace("total", "Total")
                .replace("unit_price", "Unit Price")
                .replace("quantity", "Quantity")
                if used_deterministic_recovery
                else _strategy_label(strategy)
            )
            formula_detail = (
                f"{formula}; "
                + ", ".join(
                    f"{key}={_display_value(value)}"
                    for key, value in inputs.items()
                )
                if used_deterministic_recovery
                else _strategy_label(strategy)
            )
            audit.append(CleaningAuditEntry(
                action_type=action_type,
                action=action,
                column=str(column),
                source_file=source,
                row_identifier=row_identifier(df, index),
                row_index=str(index),
                original_state=(
                    "Not provided by source file"
                    if missing_type == "structural"
                    else "Column existed, but this record had no value"
                ),
                resulting_value=replacement,
                strategy=(
                    formula
                    if used_deterministic_recovery
                    else _strategy_label(strategy)
                ),
                strategy_scope=scope,
                missing_type=missing_type,
                rows_removed=0,
                reason=(
                    "Advanced structural override selected"
                    if missing_type == "structural"
                    else "Recovered from validated business relationship"
                    if used_deterministic_recovery
                    else "User-approved missing-value action"
                ),
                recorded_at=recorded_at,
                original_source_row=source_row_identifier(df, index),
                business_record_identifier=row_identifier(df, index),
                original_value=None,
                formula_or_strategy=formula_detail,
                calculation_scope=scope,
                timestamp=recorded_at,
                decision_state=MissingDecisionState.CHANGED,
                resulting_state=f"Value set to {_display_value(replacement)}",
            ))

        if filled_count:
            deterministic_count = sum(
                entry.action_type == "deterministic_recovery"
                and entry.column == str(column)
                for entry in audit
            )
            estimated_count = filled_count - deterministic_count
            if deterministic_count:
                messages.append(
                    f"Recovered {deterministic_count} blank "
                    f"{'value' if deterministic_count == 1 else 'values'} in "
                    f"'{column}' from a validated arithmetic relationship."
                )
            if not estimated_count:
                continue
            messages.append(
                f"Filled {estimated_count} blank "
                f"{'value' if estimated_count == 1 else 'values'} in "
                f"'{column}' using {_replacement_phrase(strategy)} "
                f"({'; '.join(sorted(scopes))})."
            )

    if drop_columns:
        drop_mask = pd.Series(False, index=cleaned.index, dtype=bool)
        trigger_columns: dict[object, list[str]] = {}
        for column in drop_columns:
            if column not in cleaned.columns:
                continue
            structural = structural_missing_mask(cleaned, column, source_schemas)
            eligible = cleaned[column].isna()
            if not allow_structural_override:
                eligible &= ~structural
            drop_mask |= eligible
            for index in cleaned.index[eligible]:
                trigger_columns.setdefault(index, []).append(str(column))

        for index in cleaned.index[drop_mask]:
            source = (
                _display_value(cleaned.at[index, SOURCE_FILE_COLUMN])
                if SOURCE_FILE_COLUMN in cleaned.columns
                else "—"
            )
            columns = ", ".join(trigger_columns.get(index, []))
            audit.append(CleaningAuditEntry(
                action_type="remove_incomplete_row",
                action="Removed incomplete row",
                column=columns,
                source_file=source,
                row_identifier=row_identifier(df, index),
                row_index=str(index),
                original_state="One or more source-provided values were blank",
                resulting_value=None,
                strategy="Removed incomplete rows",
                strategy_scope="Selected row-level blank columns",
                missing_type="row-level",
                rows_removed=1,
                reason="User approved removal of affected rows",
                recorded_at=recorded_at,
                original_source_row=source_row_identifier(df, index),
                business_record_identifier=row_identifier(df, index),
                original_value="One or more source-provided values were blank",
                formula_or_strategy="Removed incomplete row",
                calculation_scope="Selected missing-value columns",
                timestamp=recorded_at,
                decision_state=MissingDecisionState.CHANGED,
                resulting_state="Record removed from the cleaned dataset",
            ))

        removed = int(drop_mask.sum())
        if removed:
            cleaned = cleaned.loc[~drop_mask].copy()
            messages.append(
                f"Removed {removed} incomplete "
                f"{'row' if removed == 1 else 'rows'}."
            )

    for column in leave_blank_columns:
        if column not in cleaned.columns:
            continue
        structural = structural_missing_mask(cleaned, column, source_schemas)
        approved = cleaned[column].isna()
        if not allow_structural_override:
            approved &= ~structural
        approved_indices = list(cleaned.index[approved])
        if not approved_indices:
            continue
        audit.extend(_aggregate_blank_decision(
            cleaned,
            str(column),
            approved_indices,
            decision_state=MissingDecisionState.APPROVED_UNCHANGED,
            action="Approved to remain blank",
            resulting_state="Blank retained by user decision",
            reason="User approved Leave blank",
            strategy="Leave blank",
            recorded_at=recorded_at,
        ))
        messages.append(
            f"Approved {len(approved_indices):,} blank "
            f"{'value' if len(approved_indices) == 1 else 'values'} in "
            f"'{column}' to remain blank."
        )

    for column, attempted_indices in attempted_fill_indices.items():
        if column not in cleaned.columns:
            continue
        failed_indices = [
            index
            for index in attempted_indices
            if index in cleaned.index and pd.isna(cleaned.at[index, column])
        ]
        if not failed_indices:
            continue
        audit.extend(_aggregate_blank_decision(
            cleaned,
            column,
            failed_indices,
            decision_state=MissingDecisionState.FAILED_OR_UNRESOLVED,
            action="Approved action could not be completed",
            resulting_state="Blank remains unresolved",
            reason="The selected action did not produce a valid value",
            strategy=str(strategies.get(column, "")),
            recorded_at=recorded_at,
        ))

    integrity_report = validate_integrity(
        cleaned,
        trusted_relationships,
    )
    current_missing = classify_missing_values(cleaned, source_schemas)
    decision_summary = missing_decision_summary(current_missing, audit)
    return MissingValueResult(
        cleaned=cleaned,
        audit=tuple(audit),
        messages=tuple(messages),
        integrity_report=integrity_report,
        decision_summary=decision_summary,
    )


def structural_blank_audit(
    df: pd.DataFrame,
    source_schemas: Mapping[str, Iterable[str]] | None,
    *,
    recorded_at: str | None = None,
    detailed: bool = False,
) -> list[CleaningAuditEntry]:
    """Return aggregated source-column traceability for unavailable cells.

    Set ``detailed=True`` only when row-level structural traceability is
    explicitly required.
    """
    if source_schemas is None:
        source_schemas = df.attrs.get("source_schemas")
    if not source_schemas:
        return []
    recorded_at = recorded_at or datetime.now(timezone.utc).isoformat(
        timespec="seconds"
    )
    entries: list[CleaningAuditEntry] = []
    for column in df.columns:
        structural = structural_missing_mask(df, str(column), source_schemas)
        if not structural.any():
            continue
        if detailed:
            groups = [
                (index, [index])
                for index in df.index[structural]
            ]
        else:
            groups = [
                (source, list(indices))
                for source, indices in df.loc[structural].groupby(
                    SOURCE_FILE_COLUMN,
                    dropna=False,
                ).groups.items()
            ]
        for group_value, indices in groups:
            index = indices[0]
            source = _display_value(df.at[index, SOURCE_FILE_COLUMN])
            count = len(indices)
            aggregate = not detailed
            entries.append(CleaningAuditEntry(
                action_type="structural_blank",
                action="Kept blank because the source file did not include the field",
                column=str(column),
                source_file=source,
                row_identifier=(
                    f"{count:,} source records"
                    if aggregate
                    else row_identifier(df, index)
                ),
                row_index="Aggregated" if aggregate else str(index),
                original_state=(
                    f"{count:,} unavailable cells"
                    if aggregate
                    else "Not provided by source file"
                ),
                resulting_value=None,
                strategy="Leave blank",
                strategy_scope=f"Source schema: {source}",
                missing_type="structural",
                rows_removed=0,
                reason="Column not provided by source",
                recorded_at=recorded_at,
                original_source_row=(
                    "Aggregated source-column entry"
                    if aggregate
                    else source_row_identifier(df, index)
                ),
                business_record_identifier=(
                    f"{count:,} source records"
                    if aggregate
                    else row_identifier(df, index)
                ),
                original_value=(
                    f"{count:,} unavailable cells"
                    if aggregate
                    else None
                ),
                formula_or_strategy="Kept blank",
                calculation_scope=f"Source file and column: {source} / {column}",
                timestamp=recorded_at,
                decision_state=MissingDecisionState.UNAVAILABLE_FROM_SOURCE,
                affected_row_count=count,
                resulting_state="Intentionally unavailable",
            ))
    return entries


def audit_records(
    entries: Iterable[CleaningAuditEntry | Mapping[str, Any]],
) -> list[dict[str, Any]]:
    """Normalize typed and dictionary audit entries."""
    records: list[dict[str, Any]] = []
    for entry in entries:
        if isinstance(entry, CleaningAuditEntry):
            records.append(entry.as_record())
        else:
            records.append(dict(entry))
    return records


def audit_summary(
    entries: Iterable[CleaningAuditEntry | Mapping[str, Any]],
) -> dict[str, int]:
    """Return reconciled counts from a cleaning audit."""
    records = audit_records(entries)

    def structural_count(record: Mapping[str, Any]) -> int:
        if str(record.get("row_index")) != "Aggregated":
            return 1
        text = str(record.get("original_state", ""))
        try:
            return int(text.split(" ", 1)[0].replace(",", ""))
        except ValueError:
            return 1

    return {
        "values_filled": sum(
            1
            for record in records
            if record.get("action_type") in {
                "fill_blank",
                "deterministic_recovery",
            }
        ),
        "deterministic_recoveries": sum(
            1
            for record in records
            if record.get("action_type") == "deterministic_recovery"
        ),
        "estimated_values": sum(
            1
            for record in records
            if record.get("action_type") == "fill_blank"
        ),
        "incomplete_rows_removed": sum(
            int(record.get("rows_removed", 0))
            for record in records
            if record.get("action_type") == "remove_incomplete_row"
        ),
        "duplicate_rows_removed": sum(
            int(record.get("rows_removed", 0))
            for record in records
            if record.get("action_type") == "remove_duplicate"
        ),
        "structural_blanks_documented": sum(
            structural_count(record)
            for record in records
            if record.get("action_type") == "structural_blank"
        ),
        "approved_unchanged": sum(
            int(record.get("affected_row_count", 1) or 1)
            for record in records
            if record.get("decision_state") == (
                MissingDecisionState.APPROVED_UNCHANGED.value
            )
            or record.get("action_type") == (
                MissingDecisionState.APPROVED_UNCHANGED.value
            )
        ),
        "failed_or_unresolved": sum(
            int(record.get("affected_row_count", 1) or 1)
            for record in records
            if record.get("decision_state") == (
                MissingDecisionState.FAILED_OR_UNRESOLVED.value
            )
            or record.get("action_type") == (
                MissingDecisionState.FAILED_OR_UNRESOLVED.value
            )
        ),
    }


def missing_decision_summary(
    missing_summary: pd.DataFrame,
    entries: Iterable[CleaningAuditEntry | Mapping[str, Any]],
    *,
    pending_review_override: int | None = None,
    integrity_failures: int = 0,
) -> MissingDecisionSummary:
    """Return mutually exclusive missing-decision counts from data and audit."""
    counts = audit_summary(entries)
    current_row_level = (
        int(missing_summary["row_level_count"].sum())
        if "row_level_count" in missing_summary
        else 0
    )
    approved = int(counts["approved_unchanged"])
    failed = max(
        int(counts["failed_or_unresolved"]),
        int(integrity_failures),
    )
    pending = (
        max(current_row_level - approved - int(counts["failed_or_unresolved"]), 0)
        if pending_review_override is None
        else max(int(pending_review_override), 0)
    )
    unavailable = (
        int(missing_summary["structural_count"].sum())
        if "structural_count" in missing_summary
        else 0
    )
    return MissingDecisionSummary(
        changed=int(counts["values_filled"]),
        approved_unchanged=approved,
        unavailable_from_source=unavailable,
        pending_review=pending,
        failed_or_unresolved=failed,
        incomplete_rows_removed=int(counts["incomplete_rows_removed"]),
    )


def build_review_summary(
    missing_summary: pd.DataFrame,
    entries: Iterable[CleaningAuditEntry | Mapping[str, Any]],
    *,
    duplicate_rows_removed: int = 0,
    pending_review_override: int | None = None,
    integrity_failures: int = 0,
) -> ReviewSummary:
    """Build the only summary contract used by completed cleaning views."""
    decision = missing_decision_summary(
        missing_summary,
        entries,
        pending_review_override=pending_review_override,
        integrity_failures=integrity_failures,
    )
    return ReviewSummary(
        recovered_or_filled_count=decision.changed,
        approved_unchanged_count=decision.approved_unchanged,
        unavailable_from_source_count=decision.unavailable_from_source,
        pending_decision_count=decision.pending_review,
        failed_or_unresolved_count=decision.failed_or_unresolved,
        duplicate_rows_removed=max(int(duplicate_rows_removed), 0),
        integrity_status=(
            "Passed" if decision.failed_or_unresolved == 0 else "Needs review"
        ),
    )


def missing_status_by_column(
    df: pd.DataFrame,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    entries: Iterable[CleaningAuditEntry | Mapping[str, Any]] = (),
) -> dict[object, MissingColumnStatus]:
    """Return mutually exclusive blank statuses for every dataset field."""
    classified = classify_missing_values(df, source_schemas)
    records = audit_records(entries)
    approved_by_column: dict[str, int] = {}
    failed_by_column: dict[str, int] = {}
    for record in records:
        column = str(record.get("column", ""))
        count = int(record.get("affected_row_count", 1) or 1)
        decision_state = str(record.get("decision_state", ""))
        action_type = str(record.get("action_type", ""))
        if (
            decision_state == MissingDecisionState.APPROVED_UNCHANGED.value
            or action_type == MissingDecisionState.APPROVED_UNCHANGED.value
        ):
            approved_by_column[column] = (
                approved_by_column.get(column, 0) + count
            )
        elif (
            decision_state == MissingDecisionState.FAILED_OR_UNRESOLVED.value
            or action_type == MissingDecisionState.FAILED_OR_UNRESOLVED.value
        ):
            failed_by_column[column] = failed_by_column.get(column, 0) + count

    statuses: dict[object, MissingColumnStatus] = {}
    for _, row in classified.iterrows():
        column = row["column"]
        approved = approved_by_column.get(str(column), 0)
        failed = failed_by_column.get(str(column), 0)
        row_level = int(row["row_level_count"])
        statuses[column] = MissingColumnStatus(
            approved_blank=approved,
            unavailable_from_source=int(row["structural_count"]),
            decisions_pending=max(row_level - approved - failed, 0),
            failed_or_unresolved=failed,
        )
    return statuses


def missing_completion_counts(
    missing_summary: pd.DataFrame,
    entries: Iterable[CleaningAuditEntry | Mapping[str, Any]],
) -> dict[str, int]:
    """Return the computed counts shown by the completion view."""
    summary = missing_decision_summary(missing_summary, entries)
    return {
        "values_handled": summary.changed,
        "approved_unchanged": summary.approved_unchanged,
        "unavailable_from_source": summary.unavailable_from_source,
        "decisions_pending": summary.pending_review,
        "failed_or_unresolved": summary.failed_or_unresolved,
    }


def business_action_text(action: object) -> str:
    """Convert legacy internal strategy codes into readable report language."""
    text = str(action)
    replacements = {
        "fill_median": "filled with the median",
        "fill_mean": "filled with the average",
        "fill_mode": "filled with the most common value",
        "fill_zero": "filled with zero",
        "drop_rows": "removed incomplete rows",
        "leave_blank": "approved to remain blank",
    }
    for internal, readable in replacements.items():
        text = text.replace(internal, readable)
    return text


def grouped_audit_change_summaries(
    entries: Iterable[CleaningAuditEntry | Mapping[str, Any]],
) -> list[str]:
    """Summarize approved missing-value changes by column and method."""
    records = audit_records(entries)
    grouped: dict[tuple[str, str], int] = {}
    removed_rows = 0
    for record in records:
        action_type = str(record.get("action_type", ""))
        if action_type == "remove_incomplete_row":
            removed_rows += int(record.get("rows_removed", 0))
            continue
        if action_type in {
            MissingDecisionState.APPROVED_UNCHANGED.value,
            "approved_unchanged",
        }:
            column = str(record.get("column", "Field"))
            count = int(record.get("affected_row_count", 1) or 1)
            grouped[(column, "approved to remain blank")] = (
                grouped.get((column, "approved to remain blank"), 0) + count
            )
            continue
        if action_type not in {"deterministic_recovery", "fill_blank"}:
            continue
        column = str(record.get("column", "Field"))
        if action_type == "deterministic_recovery":
            method = str(
                record.get("action")
                or "Recovered from a validated relationship"
            )
        else:
            method = str(
                record.get("action")
                or record.get("formula_or_strategy")
                or "Handled using an approved action"
            )
        grouped[(column, method)] = grouped.get((column, method), 0) + 1

    summaries = [
        f"{str(column).replace('_', ' ').title()}: {count:,} "
        f"{'value' if count == 1 else 'values'} {method[0].lower() + method[1:]}."
        for (column, method), count in grouped.items()
    ]
    if removed_rows:
        summaries.append(
            f"{removed_rows:,} incomplete "
            f"{'row was' if removed_rows == 1 else 'rows were'} removed "
            "using an approved action."
        )
    return summaries


def grouped_audit_change_table(
    entries: Iterable[CleaningAuditEntry | Mapping[str, Any]],
) -> pd.DataFrame:
    """Return compact grouped completion details without record-ID lists."""
    records = audit_records(entries)
    grouped: dict[tuple[str, str], int] = {}
    for record in records:
        action_type = str(record.get("action_type", ""))
        if action_type == "remove_incomplete_row":
            column = str(record.get("column") or "Affected rows")
            method = "Removed incomplete rows"
            count = int(record.get("rows_removed", 0))
        elif action_type == MissingDecisionState.APPROVED_UNCHANGED.value:
            column = str(record.get("column") or "Field")
            method = "Approved to remain blank"
            count = int(record.get("affected_row_count", 1) or 1)
        elif action_type in {"deterministic_recovery", "fill_blank"}:
            column = str(record.get("column") or "Field")
            method = str(
                record.get("action")
                or record.get("formula_or_strategy")
                or "Approved cleaning action"
            )
            count = 1
        else:
            continue
        grouped[(column, method)] = grouped.get((column, method), 0) + count

    return pd.DataFrame(
        [
            {
                "Field": str(column).replace("_", " ").title(),
                "Decision or method": method,
                "Records affected": count,
            }
            for (column, method), count in grouped.items()
        ],
        columns=["Field", "Decision or method", "Records affected"],
    )


def cleaning_change_preview(
    entries: Iterable[CleaningAuditEntry | Mapping[str, Any]],
) -> pd.DataFrame:
    """Return a client-facing before/after preview from the cleaning audit."""
    preview_rows: list[dict[str, Any]] = []
    for record in audit_records(entries):
        action_type = str(record.get("action_type", ""))
        if action_type not in {
            "deterministic_recovery",
            "fill_blank",
            "remove_incomplete_row",
            MissingDecisionState.APPROVED_UNCHANGED.value,
        }:
            continue
        original = record.get("original_value")
        if original is None or (
            not isinstance(original, (list, dict))
            and pd.isna(original)
        ):
            original = "Blank"
        current = record.get("resulting_value")
        if action_type == "remove_incomplete_row":
            current = "Row removed"
        elif action_type == MissingDecisionState.APPROVED_UNCHANGED.value:
            current = "Blank"
        elif current is None or (
            not isinstance(current, (list, dict))
            and pd.isna(current)
        ):
            current = "—"
        method = str(
            record.get("action")
            or record.get("formula_or_strategy")
            or "Approved cleaning action"
        )
        record_identifiers = [
            identifier.strip()
            for identifier in str(
                record.get("affected_record_identifiers") or ""
            ).split(",")
            if identifier.strip()
        ]
        if not record_identifiers:
            record_identifiers = [
                record.get("business_record_identifier")
                or record.get("row_identifier")
                or "—"
            ]
        for record_id in record_identifiers:
            preview_rows.append({
                "Record ID": record_id,
                "Field": record.get("column") or "—",
                "Original": original,
                "Current": current,
                "Decision or method": method,
                "Source file": record.get("source_file") or "—",
            })
    return pd.DataFrame(
        preview_rows,
        columns=[
            "Record ID",
            "Field",
            "Original",
            "Current",
            "Decision or method",
            "Source file",
        ],
    )
