"""Merge, clean, and transform DataFrames."""

from datetime import datetime, timezone
from typing import Any, Iterable, Mapping, Sequence

import pandas as pd

from src.data_quality import (
    CleaningAuditEntry,
    SourceSchemas,
    apply_missing_value_strategies,
    audit_records,
    audit_summary,
    build_source_schemas,
    classify_missing_values,
    missing_decision_summary,
    recommend_row_level_strategy,
    row_identifier,
    source_row_identifier,
)
from src.integrity import (
    IntegrityReport,
    RelationshipRule,
    integrity_issues_frame,
    validate_integrity,
)
from src.logger_setup import setup_logger

logger = setup_logger(__name__)


class MergeError(Exception):
    """Raised when merging fails due to incompatible data."""


class SchemaWarning:
    """Represents a detected schema difference between files."""

    def __init__(
        self,
        message: str,
        columns_only_in_some: dict[str, list[str]],
    ) -> None:
        self.message = message
        self.columns_only_in_some = columns_only_in_some

    def __repr__(self) -> str:
        return f"SchemaWarning({self.message})"


def add_source_column(
    dfs: list[tuple[str, pd.DataFrame]]
) -> list[pd.DataFrame]:
    """Add a 'source_file' column to each DataFrame for traceability.

    Args:
        dfs: List of (filename, DataFrame) tuples.

    Returns:
        List of DataFrames with a new 'source_file' column.
    """
    result = []
    for name, df in dfs:
        df_with_source = df.copy()
        df_with_source["source_file"] = name
        result.append(df_with_source)
    return result


def detect_schema_differences(
    dfs: list[pd.DataFrame],
    file_names: list[str],
) -> list[SchemaWarning]:
    """Compare column sets across DataFrames and report differences.

    Args:
        dfs: List of DataFrames to compare.
        file_names: Corresponding filenames for labeling.

    Returns:
        List of SchemaWarning objects. Empty if all schemas match.
    """
    warnings: list[SchemaWarning] = []
    all_columns = set()
    for df in dfs:
        all_columns.update(df.columns)

    for df, fname in zip(dfs, file_names):
        missing = all_columns - set(df.columns)
        if missing:
            col_list = sorted(missing)
            warnings.append(SchemaWarning(
                message=(
                    f"'{fname}' is missing columns: {col_list}. "
                    "These will be filled with NaN."
                ),
                columns_only_in_some={fname: col_list},
            ))

    return warnings


def compute_schema_compatibility(
    dfs: list[pd.DataFrame],
    file_names: list[str],
) -> dict[tuple[str, str], float]:
    """Compute Jaccard similarity between column sets of every file pair.

    Jaccard similarity = |intersection| / |union|.
    A score of 1.0 means identical columns; 0.0 means no overlap.

    Args:
        dfs: List of DataFrames.
        file_names: Corresponding filenames.

    Returns:
        Dict mapping (file_a, file_b) tuples to similarity scores.
    """
    column_sets = [set(df.columns) for df in dfs]
    scores: dict[tuple[str, str], float] = {}

    for i in range(len(dfs)):
        for j in range(i + 1, len(dfs)):
            intersection = column_sets[i] & column_sets[j]
            union = column_sets[i] | column_sets[j]
            score = len(intersection) / len(union) if union else 0.0
            scores[(file_names[i], file_names[j])] = round(score, 3)

    return scores


def merge_datasets(
    dfs: list[tuple[str, pd.DataFrame]],
    keep_all_columns: bool = True,
) -> tuple[pd.DataFrame, list[SchemaWarning]]:
    """Merge multiple DataFrames row-wise into a single dataset.

    Each row is tagged with its source file. Schema differences are
    reported as warnings.

    Args:
        dfs: List of (filename, DataFrame) tuples.
        keep_all_columns: If True, keeps all columns (union).
                          If False, keeps only common columns (intersection).

    Returns:
        Tuple of (merged DataFrame, list of SchemaWarning).

    Raises:
        MergeError: If no DataFrames are provided.
    """
    if not dfs:
        raise MergeError("No DataFrames to merge.")

    file_names = [name for name, _ in dfs]
    source_schemas = build_source_schemas(dfs)

    # Tag each row with its origin
    tagged_dfs = add_source_column(dfs)

    # Detect schema differences before merging
    schema_warnings = detect_schema_differences(tagged_dfs, file_names)

    if keep_all_columns:
        merged = pd.concat(tagged_dfs, ignore_index=True)
        logger.info(
            "Merged %d files -> %d rows, %d columns (union)",
            len(dfs), len(merged), len(merged.columns),
        )
    else:
        common_cols = set(tagged_dfs[0].columns)
        for df in tagged_dfs[1:]:
            common_cols &= set(df.columns)
        common_cols = list(common_cols)
        merged = pd.concat(
            [df[common_cols] for df in tagged_dfs],
            ignore_index=True,
        )
        logger.info(
            "Merged %d files -> %d rows, %d columns (intersection)",
            len(dfs), len(merged), len(merged.columns),
        )

    merged.attrs["source_schemas"] = source_schemas
    source_row_numbers: dict[object, int] = {}
    merged_index = 0
    for tagged in tagged_dfs:
        for source_position, _ in enumerate(tagged.index, start=2):
            source_row_numbers[merged_index] = source_position
            merged_index += 1
    merged.attrs["source_row_numbers"] = source_row_numbers
    return merged, schema_warnings


def get_merge_summary(merged: pd.DataFrame, file_names: list[str]) -> dict:
    """Return per-source row count and total summary.

    Args:
        merged: The merged DataFrame (must have 'source_file' column).
        file_names: Original file names for reporting.

    Returns:
        Dict with 'per_file' (list of dicts) and 'total_rows'.
    """
    per_file = []
    for name in file_names:
        count = int((merged["source_file"] == name).sum())
        per_file.append({"file": name, "rows": count})

    return {
        "per_file": per_file,
        "total_rows": len(merged),
        "total_columns": len(merged.columns),
    }


# ── Data cleaning ──────────────────────────────────────────


def recommend_strategy(series: pd.Series) -> str:
    """Recommend a missing-value strategy based on column data type and missing rate.

    Args:
        series: A pandas Series to analyze.

    Returns:
        Strategy name: 'fill_median', 'fill_mode', or 'drop_rows'.
    """
    missing_ratio = series.isna().mean()

    if pd.api.types.is_numeric_dtype(series):
        if missing_ratio < 0.05:
            return "drop_rows"
        return "fill_median"

    # Categorical / text / date columns
    if missing_ratio < 0.05:
        return "drop_rows"
    return "fill_mode"


def remove_duplicates(
    df: pd.DataFrame,
    subset: list[str] | None = None,
    keep: str = "first",
) -> tuple[pd.DataFrame, dict]:
    """Remove duplicate rows and report what was removed.

    Args:
        df: Source DataFrame.
        subset: Columns to check for duplicates. None = all columns.
        keep: Which duplicate to keep ('first', 'last', False).

    Returns:
        Tuple of (cleaned DataFrame, report dict).
    """
    before = len(df)
    duplicate_mask = df.duplicated(subset=subset, keep=keep)
    cleaned = df.loc[~duplicate_mask].copy()
    removed = before - len(cleaned)
    recorded_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
    audit = []
    for index in df.index[duplicate_mask]:
        source = (
            str(df.at[index, "source_file"])
            if "source_file" in df.columns
            else "—"
        )
        audit.append(CleaningAuditEntry(
            action_type="remove_duplicate",
            action="Removed repeated record",
            column=", ".join(str(column) for column in subset)
            if subset
            else "Complete row",
            source_file=source,
            row_identifier=row_identifier(df, index),
            row_index=str(index),
            original_state="Repeated record",
            resulting_value=None,
            strategy="Kept the first matching record",
            strategy_scope=(
                "Selected identity columns" if subset else "Complete row"
            ),
            missing_type="Not applicable",
            rows_removed=1,
            reason="User approved duplicate removal",
            recorded_at=recorded_at,
            original_source_row=source_row_identifier(df, index),
            business_record_identifier=row_identifier(df, index),
            original_value="Repeated record",
            formula_or_strategy="Kept the first matching record",
            calculation_scope=(
                "Selected identity columns" if subset else "Complete row"
            ),
            timestamp=recorded_at,
        ).as_record())

    report = {
        "before": before,
        "after": len(cleaned),
        "removed": removed,
        "subset": subset,
        "audit": audit,
        "integrity_report": validate_integrity(cleaned),
    }

    logger.info("Removed %d duplicate rows (%d -> %d)", removed, before, len(cleaned))
    return cleaned, report


def detect_missing_values(
    df: pd.DataFrame,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
) -> pd.DataFrame:
    """Return a DataFrame summarizing missing values per column.

    Args:
        df: Source DataFrame.

    Returns:
        DataFrame with columns: column, missing_count, missing_pct, dtype.
    """
    return classify_missing_values(df, source_schemas)


def recommend_missing_strategy(
    df: pd.DataFrame,
    column: str,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    *,
    configured_relationships: Sequence[RelationshipRule] | None = None,
) -> str:
    """Recommend a conservative strategy for genuine row-level blanks only."""
    return recommend_row_level_strategy(
        df,
        column,
        source_schemas,
        configured_relationships=configured_relationships,
    )


def handle_missing_values(
    df: pd.DataFrame,
    strategies: dict[str, str],
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    *,
    business_group_columns: Sequence[str] | None = None,
) -> tuple[pd.DataFrame, list[str]]:
    """Apply per-column missing-value strategies to a DataFrame.

    **Execution order**: fill strategies are applied first (non-destructive),
    then row drops are applied last (destructive). This ensures fills use
    the full dataset before any rows are removed.

    Supported strategies:
        - 'drop_rows': drop rows where this column is NaN (applied last)
        - 'fill_mean': fill with column mean (numeric only)
        - 'fill_median': fill with column median (numeric only)
        - 'fill_mode': fill with column mode (categorical)
        - 'fill_zero': fill with 0
        - 'fill_value:XYZ': fill with custom value XYZ

    Args:
        df: Source DataFrame.
        strategies: Dict mapping column name to strategy name.

    Returns:
        Tuple of (cleaned DataFrame, list of action descriptions).
    """
    result = apply_missing_value_strategies(
        df,
        strategies,
        source_schemas,
        business_group_columns=business_group_columns,
    )
    logger.info(
        "Applied %d missing-value audit actions across %d selected columns",
        len(result.audit),
        len(strategies),
    )
    legacy_compatible_messages = [
        message.replace("Removed ", "Dropped ", 1)
        if message.startswith("Removed ")
        else message
        for message in result.messages
    ]
    return result.cleaned, legacy_compatible_messages


def generate_cleaning_report(
    original: pd.DataFrame,
    cleaned: pd.DataFrame,
    duplicate_report: dict | None = None,
    missing_actions: list[str] | None = None,
    *,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    cleaning_audit: Iterable[CleaningAuditEntry | Mapping[str, Any]] | None = None,
    integrity_report: IntegrityReport | None = None,
    pending_review_override: int | None = None,
) -> dict:
    """Generate a full cleaning report.

    Args:
        original: DataFrame before cleaning.
        cleaned: DataFrame after cleaning.
        duplicate_report: Output from remove_duplicates().
        missing_actions: Action descriptions from handle_missing_values().

    Returns:
        Dict with keys: rows_before, rows_after, columns, duplicates_removed,
        missing_actions, total_missing_before, total_missing_after.
    """
    before_missing = classify_missing_values(original, source_schemas)
    after_missing = classify_missing_values(cleaned, source_schemas)
    audit = audit_records(cleaning_audit or [])
    summary = audit_summary(audit)
    integrity_report = integrity_report or validate_integrity(cleaned)
    structural_before = int(before_missing["structural_count"].sum())
    structural_after = int(after_missing["structural_count"].sum())
    row_level_before = int(before_missing["row_level_count"].sum())
    row_level_after = int(after_missing["row_level_count"].sum())
    decisions = missing_decision_summary(
        after_missing,
        audit,
        pending_review_override=pending_review_override,
        integrity_failures=integrity_report.severe_count,
    )
    warnings = []
    if structural_after:
        warnings.append(
            "Some blank cells are caused by source files that did not contain "
            "the corresponding column. These cells were preserved as unavailable "
            "rather than replaced with estimated business values."
        )
    if summary["estimated_values"]:
        warnings.append(
            "Estimated values are marked in the Cleaning Audit and should be "
            "validated before operational use."
        )
    rows_removed = (
        summary["incomplete_rows_removed"]
        + summary["duplicate_rows_removed"]
    )
    if rows_removed:
        warnings.append("Removed rows are listed in the Cleaning Audit.")
    if integrity_report.severe_count:
        warnings.append(
            f"{integrity_report.severe_count:,} severe relationship "
            "integrity issue(s) require acknowledgment before download."
        )

    return {
        "rows_before": len(original),
        "rows_after": len(cleaned),
        "columns": list(cleaned.columns),
        "duplicates_removed": duplicate_report["removed"] if duplicate_report else 0,
        "missing_actions": missing_actions or [],
        "total_missing_before": int(original.isna().sum().sum()),
        "total_missing_after": int(cleaned.isna().sum().sum()),
        "structural_missing_before": structural_before,
        "structural_missing_after": structural_after,
        "row_level_missing_before": row_level_before,
        "row_level_missing_after": row_level_after,
        "missing_values_reviewed": decisions.reviewed,
        "values_changed": decisions.changed,
        "approved_unchanged": decisions.approved_unchanged,
        "decisions_pending": decisions.pending_review,
        "unavailable_from_source": decisions.unavailable_from_source,
        "failed_or_unresolved": decisions.failed_or_unresolved,
        "values_filled": summary["values_filled"],
        "deterministic_recoveries": summary["deterministic_recoveries"],
        "estimated_values": summary["estimated_values"],
        "incomplete_rows_removed": summary["incomplete_rows_removed"],
        "integrity_passed": integrity_report.passed,
        "integrity_issue_count": len(integrity_report.issues),
        "integrity_failures": integrity_report.severe_count,
        "severe_integrity_issue_count": integrity_report.severe_count,
        "integrity_issues": integrity_report.issue_records(),
        "integrity_review": integrity_issues_frame(integrity_report),
        "warnings": warnings,
        "audit": audit,
    }
