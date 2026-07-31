"""Reusable arithmetic relationship detection and integrity validation."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from math import isfinite
from typing import Any, Iterable, Mapping, Sequence

import pandas as pd


@dataclass(frozen=True)
class RelationshipRule:
    """A candidate or explicitly configured arithmetic relationship."""

    rule_id: str
    result_column: str
    factor_columns: tuple[str, str]
    label: str
    tolerance: float = 0.01
    minimum_complete_rows: int = 8
    required_pass_rate: float = 0.98
    whole_number_columns: tuple[str, ...] = ()
    non_negative_columns: tuple[str, ...] = ()

    @property
    def involved_columns(self) -> tuple[str, ...]:
        return (*self.factor_columns, self.result_column)


@dataclass(frozen=True)
class ValidatedRelationship:
    """Evidence that a relationship is safe to use for validation or recovery."""

    rule: RelationshipRule
    complete_rows: int
    matching_rows: int
    pass_rate: float
    explicitly_configured: bool

    @property
    def expected_relationship(self) -> str:
        return self.rule.label


@dataclass(frozen=True)
class IntegrityIssue:
    """One record that does not satisfy a validated relationship."""

    rule: str
    affected_record_identifier: str
    original_source_row: str
    source_file: str
    involved_columns: str
    actual_values: str
    expected_relationship: str
    expected_value: float
    difference: float
    severity: str

    def as_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class IntegrityReport:
    """Structured result of all configured or validated integrity checks."""

    relationships: tuple[ValidatedRelationship, ...]
    issues: tuple[IntegrityIssue, ...]

    @property
    def severe_count(self) -> int:
        return sum(issue.severity == "Severe" for issue in self.issues)

    @property
    def passed(self) -> bool:
        return self.severe_count == 0

    def issue_records(self) -> list[dict[str, Any]]:
        return [issue.as_record() for issue in self.issues]


DEFAULT_CANDIDATE_RELATIONSHIPS = (
    RelationshipRule(
        rule_id="total_equals_quantity_times_unit_price",
        result_column="total",
        factor_columns=("quantity", "unit_price"),
        label="Total = Quantity × Unit Price",
        tolerance=0.01,
        minimum_complete_rows=8,
        required_pass_rate=0.98,
        whole_number_columns=("quantity",),
        non_negative_columns=("quantity", "unit_price", "total"),
    ),
)


def _column_lookup(df: pd.DataFrame) -> dict[str, str]:
    return {
        str(column).strip().lower(): str(column)
        for column in df.columns
    }


def _resolve_rule(
    df: pd.DataFrame,
    rule: RelationshipRule,
) -> RelationshipRule | None:
    lookup = _column_lookup(df)
    names = [
        rule.result_column,
        *rule.factor_columns,
    ]
    if any(name.strip().lower() not in lookup for name in names):
        return None
    return RelationshipRule(
        rule_id=rule.rule_id,
        result_column=lookup[rule.result_column.strip().lower()],
        factor_columns=(
            lookup[rule.factor_columns[0].strip().lower()],
            lookup[rule.factor_columns[1].strip().lower()],
        ),
        label=rule.label,
        tolerance=rule.tolerance,
        minimum_complete_rows=rule.minimum_complete_rows,
        required_pass_rate=rule.required_pass_rate,
        whole_number_columns=tuple(
            lookup.get(column.strip().lower(), column)
            for column in rule.whole_number_columns
        ),
        non_negative_columns=tuple(
            lookup.get(column.strip().lower(), column)
            for column in rule.non_negative_columns
        ),
    )


def _numeric_complete_rows(
    df: pd.DataFrame,
    rule: RelationshipRule,
) -> pd.DataFrame:
    numeric = pd.DataFrame(index=df.index)
    for column in rule.involved_columns:
        numeric[column] = pd.to_numeric(df[column], errors="coerce")
    return numeric.dropna(subset=list(rule.involved_columns))


def detect_validated_relationships(
    df: pd.DataFrame,
    configured_rules: Sequence[RelationshipRule] | None = None,
) -> tuple[ValidatedRelationship, ...]:
    """Return explicitly configured or strongly evidenced relationships.

    Name matching only identifies candidates. An unconfigured candidate is
    accepted only when almost all complete rows satisfy the arithmetic rule.
    """
    explicitly_configured = configured_rules is not None
    candidates = (
        tuple(configured_rules)
        if configured_rules is not None
        else DEFAULT_CANDIDATE_RELATIONSHIPS
    )
    validated: list[ValidatedRelationship] = []

    for candidate in candidates:
        rule = _resolve_rule(df, candidate)
        if rule is None:
            continue
        complete = _numeric_complete_rows(df, rule)
        if complete.empty and not explicitly_configured:
            continue
        expected = complete[rule.factor_columns[0]] * complete[
            rule.factor_columns[1]
        ]
        difference = (complete[rule.result_column] - expected).abs()
        matching_rows = int((difference <= rule.tolerance).sum())
        complete_rows = len(complete)
        pass_rate = (
            matching_rows / complete_rows
            if complete_rows
            else 0.0
        )
        strong_evidence = (
            complete_rows >= rule.minimum_complete_rows
            and pass_rate >= rule.required_pass_rate
        )
        if explicitly_configured or strong_evidence:
            validated.append(ValidatedRelationship(
                rule=rule,
                complete_rows=complete_rows,
                matching_rows=matching_rows,
                pass_rate=pass_rate,
                explicitly_configured=explicitly_configured,
            ))
    return tuple(validated)


def _display_value(value: Any) -> str:
    if value is None:
        return "—"
    try:
        if bool(pd.isna(value)):
            return "—"
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _record_identifier(df: pd.DataFrame, index: object) -> str:
    candidates = [
        column
        for column in df.columns
        if str(column).strip().lower() == "id"
        or str(column).strip().lower().endswith("_id")
    ]
    if candidates:
        value = df.at[index, candidates[0]]
        if not pd.isna(value):
            return _display_value(value)
    return f"Row {index}"


def _source_row(df: pd.DataFrame, index: object) -> str:
    mapping = df.attrs.get("source_row_numbers", {})
    value = mapping.get(index)
    if value is not None:
        return str(value)
    return str(index)


def _source_file(df: pd.DataFrame, index: object) -> str:
    if "source_file" not in df.columns:
        return "—"
    return _display_value(df.at[index, "source_file"])


def recover_deterministic_value(
    df: pd.DataFrame,
    index: object,
    target_column: str,
    relationships: Sequence[ValidatedRelationship] | None = None,
) -> tuple[Any, str, str, Mapping[str, Any]] | None:
    """Recover one missing value from a validated product relationship."""
    relationships = tuple(
        relationships or detect_validated_relationships(df)
    )
    for validation in relationships:
        rule = validation.rule
        if target_column not in rule.involved_columns:
            continue
        if target_column == rule.result_column:
            first, second = rule.factor_columns
            first_value = pd.to_numeric(
                pd.Series([df.at[index, first]]),
                errors="coerce",
            ).iloc[0]
            second_value = pd.to_numeric(
                pd.Series([df.at[index, second]]),
                errors="coerce",
            ).iloc[0]
            if pd.isna(first_value) or pd.isna(second_value):
                continue
            recovered = float(first_value) * float(second_value)
            formula = (
                f"{rule.result_column} = {first} × {second}"
            )
            inputs = {first: first_value, second: second_value}
        else:
            other_factor = next(
                column
                for column in rule.factor_columns
                if column != target_column
            )
            result_value = pd.to_numeric(
                pd.Series([df.at[index, rule.result_column]]),
                errors="coerce",
            ).iloc[0]
            divisor = pd.to_numeric(
                pd.Series([df.at[index, other_factor]]),
                errors="coerce",
            ).iloc[0]
            if pd.isna(result_value) or pd.isna(divisor) or float(divisor) == 0:
                continue
            recovered = float(result_value) / float(divisor)
            formula = (
                f"{target_column} = {rule.result_column} ÷ {other_factor}"
            )
            inputs = {
                rule.result_column: result_value,
                other_factor: divisor,
            }

        if not isfinite(float(recovered)):
            continue
        if (
            target_column in rule.non_negative_columns
            and recovered < -rule.tolerance
        ):
            continue
        if target_column in rule.whole_number_columns:
            nearest = round(recovered)
            reconstructed = (
                nearest * float(inputs[other_factor])
                if target_column != rule.result_column
                else recovered
            )
            target_total = (
                float(inputs[rule.result_column])
                if target_column != rule.result_column
                else recovered
            )
            if abs(reconstructed - target_total) > rule.tolerance:
                continue
            recovered = int(nearest)
        else:
            recovered = round(float(recovered), 10)
        return recovered, formula, rule.label, inputs
    return None


def validate_integrity(
    df: pd.DataFrame,
    relationships: Sequence[ValidatedRelationship] | None = None,
    *,
    configured_rules: Sequence[RelationshipRule] | None = None,
) -> IntegrityReport:
    """Validate all complete records against trusted relationships."""
    trusted = tuple(
        relationships
        if relationships is not None
        else detect_validated_relationships(df, configured_rules)
    )
    issues: list[IntegrityIssue] = []
    for validation in trusted:
        rule = validation.rule
        complete = _numeric_complete_rows(df, rule)
        for index, row in complete.iterrows():
            expected = float(
                row[rule.factor_columns[0]]
                * row[rule.factor_columns[1]]
            )
            actual = float(row[rule.result_column])
            difference = round(actual - expected, 10)
            if abs(difference) <= rule.tolerance:
                continue
            actual_values = ", ".join(
                f"{column}={_display_value(df.at[index, column])}"
                for column in rule.involved_columns
            )
            issues.append(IntegrityIssue(
                rule=rule.rule_id,
                affected_record_identifier=_record_identifier(df, index),
                original_source_row=_source_row(df, index),
                source_file=_source_file(df, index),
                involved_columns=", ".join(rule.involved_columns),
                actual_values=actual_values,
                expected_relationship=rule.label,
                expected_value=round(expected, 2),
                difference=round(difference, 2),
                severity="Severe",
            ))
    return IntegrityReport(
        relationships=trusted,
        issues=tuple(issues),
    )


def integrity_issues_frame(report: IntegrityReport) -> pd.DataFrame:
    """Convert integrity issues into the shared Values to Review schema."""
    rows = []
    for issue in report.issues:
        rows.append({
            "Review Type": "Integrity check",
            "Record ID": issue.affected_record_identifier,
            "Original Source Row": issue.original_source_row,
            "Source File": issue.source_file,
            "Description": issue.expected_relationship,
            "Column": issue.involved_columns,
            "Value": issue.actual_values,
            "Lower Review Boundary": None,
            "Upper Review Boundary": None,
            "Reason": (
                f"Relationship differs by {issue.difference:+,.2f}."
            ),
            "Rule": issue.rule,
            "Severity": issue.severity,
            "Difference": issue.difference,
        })
    return pd.DataFrame(rows)


def combine_review_findings(
    statistical_findings: pd.DataFrame | None,
    integrity_report: IntegrityReport,
) -> pd.DataFrame:
    """Combine statistical review signals with relational integrity issues."""
    statistical = (
        pd.DataFrame()
        if statistical_findings is None
        else statistical_findings.copy()
    )
    if not statistical.empty:
        statistical["Review Type"] = "Statistical review"
        statistical["Severity"] = "Review"
    integrity = integrity_issues_frame(integrity_report)
    if statistical.empty:
        return integrity
    if integrity.empty:
        return statistical
    return pd.concat([integrity, statistical], ignore_index=True, sort=False)


def relationship_columns(
    relationships: Iterable[ValidatedRelationship],
) -> set[str]:
    """Return all columns involved in trusted relationships."""
    return {
        column
        for validation in relationships
        for column in validation.rule.involved_columns
    }
