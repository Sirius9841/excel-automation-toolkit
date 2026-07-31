"""Client-facing semantic profiles, statistics, and chart compatibility."""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, Iterable, Mapping

import pandas as pd

from src.analyzer import (
    detect_column_type,
    detect_outliers_iqr,
    friendly_column_name,
)
from src.data_quality import (
    CleaningAuditEntry,
    missing_status_by_column,
)

SEMANTIC_IDENTIFIER = "identifier"
SEMANTIC_NUMERIC = "numeric"
SEMANTIC_CATEGORICAL = "categorical"
SEMANTIC_DATE = "date"
SEMANTIC_UNSUPPORTED = "unsupported"

SEMANTIC_GROUP_LABELS = {
    SEMANTIC_NUMERIC: "Measures",
    SEMANTIC_CATEGORICAL: "Categories",
    SEMANTIC_DATE: "Dates",
    SEMANTIC_IDENTIFIER: "Identifiers",
    SEMANTIC_UNSUPPORTED: "Unavailable",
}

DEFAULT_COLUMN_OVERRIDES = {
    "source_file": SEMANTIC_IDENTIFIER,
}

IDENTIFIER_NAME_TOKENS = {
    "id",
    "identifier",
    "reference",
    "ref",
    "record_number",
    "row_number",
}

CURRENCY_NAME_TOKENS = {
    "amount",
    "cost",
    "price",
    "revenue",
    "sales",
    "total",
}

TIME_GROUP_OPTIONS = ("Day", "Week", "Month")
TIME_CALCULATION_OPTIONS = ("Total", "Average", "Median", "Record count")


@dataclass(frozen=True)
class InsightColumn:
    """One authoritative semantic classification for an Insights column."""

    name: object
    display_name: str
    semantic_type: str
    uniqueness_ratio: float

    @property
    def group_label(self) -> str:
        return SEMANTIC_GROUP_LABELS[self.semantic_type]

    @property
    def selector_label(self) -> str:
        return f"{self.group_label}: {self.display_name}"


def _normalise_name(column_name: object) -> str:
    return re.sub(
        r"[^a-z0-9]+",
        "_",
        str(column_name).lower(),
    ).strip("_")


def classify_column(
    column_name: object,
    series: pd.Series,
    *,
    configured_overrides: Mapping[str, str] | None = None,
    uniqueness_threshold: float = 0.95,
) -> InsightColumn:
    """Classify a column once for all Insights controls and presentations."""
    normalized_name = _normalise_name(column_name)
    overrides = {
        **DEFAULT_COLUMN_OVERRIDES,
        **{
            _normalise_name(name): semantic_type
            for name, semantic_type in (configured_overrides or {}).items()
        },
    }
    non_null = series.dropna()
    uniqueness_ratio = (
        float(non_null.nunique(dropna=True)) / len(non_null)
        if len(non_null)
        else 0.0
    )

    if normalized_name in overrides:
        semantic_type = overrides[normalized_name]
    elif non_null.empty:
        semantic_type = SEMANTIC_UNSUPPORTED
    elif detect_column_type(series) == "date":
        semantic_type = SEMANTIC_DATE
    else:
        identifier_name = (
            normalized_name in IDENTIFIER_NAME_TOKENS
            or normalized_name.endswith("_id")
            or normalized_name.endswith("_identifier")
            or normalized_name.endswith("_reference")
        )
        unique_identifier_hint = (
            uniqueness_ratio >= uniqueness_threshold
            and any(
                token in normalized_name
                for token in ("reference", "record", "number", "key")
            )
        )
        if identifier_name or unique_identifier_hint:
            semantic_type = SEMANTIC_IDENTIFIER
        elif pd.api.types.is_numeric_dtype(series):
            semantic_type = SEMANTIC_NUMERIC
        else:
            semantic_type = SEMANTIC_CATEGORICAL

    return InsightColumn(
        name=column_name,
        display_name=friendly_column_name(column_name),
        semantic_type=semantic_type,
        uniqueness_ratio=uniqueness_ratio,
    )


def classify_columns(
    df: pd.DataFrame,
    *,
    configured_overrides: Mapping[str, str] | None = None,
) -> dict[object, InsightColumn]:
    """Return the authoritative semantic profile for every column."""
    return {
        column: classify_column(
            column,
            df[column],
            configured_overrides=configured_overrides,
        )
        for column in df.columns
    }


def grouped_column_options(
    profiles: Mapping[object, InsightColumn],
) -> list[object]:
    """Order selector options by understandable client-facing groups."""
    order = (
        SEMANTIC_NUMERIC,
        SEMANTIC_CATEGORICAL,
        SEMANTIC_DATE,
        SEMANTIC_IDENTIFIER,
        SEMANTIC_UNSUPPORTED,
    )
    return [
        profile.name
        for semantic_type in order
        for profile in profiles.values()
        if profile.semantic_type == semantic_type
    ]


def compatible_chart_columns(
    profiles: Mapping[object, InsightColumn],
) -> dict[str, list[object]]:
    """Return only semantically valid columns for each chart type."""
    by_type = {
        semantic_type: [
            profile.name
            for profile in profiles.values()
            if profile.semantic_type == semantic_type
        ]
        for semantic_type in SEMANTIC_GROUP_LABELS
    }
    return {
        "Distribution": by_type[SEMANTIC_NUMERIC],
        "Range and review flags": by_type[SEMANTIC_NUMERIC],
        "Category comparison": by_type[SEMANTIC_CATEGORICAL],
        "Trend date": by_type[SEMANTIC_DATE],
        "Trend value": by_type[SEMANTIC_NUMERIC],
    }


def default_chart_config(profile: InsightColumn) -> dict[str, Any] | None:
    """Return the immediate, type-appropriate chart for a selected column."""
    if profile.semantic_type == SEMANTIC_NUMERIC:
        return {"type": "Distribution", "column": profile.name}
    if profile.semantic_type == SEMANTIC_CATEGORICAL:
        return {
            "type": "Category comparison",
            "column": profile.name,
            "top_n": 10,
        }
    if profile.semantic_type == SEMANTIC_DATE:
        return {"type": "Records over time", "column": profile.name}
    return None


def chart_config_is_compatible(
    config: Mapping[str, Any],
    profiles: Mapping[object, InsightColumn],
) -> bool:
    """Return whether every configured chart field has a compatible type."""
    compatible = compatible_chart_columns(profiles)
    chart_type = config.get("type")
    if chart_type in {"Distribution", "Range and review flags"}:
        return config.get("column") in compatible[chart_type]
    if chart_type == "Category comparison":
        return config.get("column") in compatible[chart_type]
    if chart_type == "Records over time":
        column = config.get("column")
        return column in compatible["Trend date"]
    if chart_type == "Trend over time":
        return (
            config.get("date_column") in compatible["Trend date"]
            and config.get("value_column") in compatible["Trend value"]
            and config.get("group_by", "Month") in TIME_GROUP_OPTIONS
            and config.get("aggregation", "sum")
            in {"sum", "mean", "median", "count"}
        )
    return False


def missing_counts(
    df: pd.DataFrame,
    column: object,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    cleaning_audit: Iterable[CleaningAuditEntry | Mapping[str, Any]] = (),
) -> tuple[int, int]:
    """Return unresolved and unavailable-from-source blank counts."""
    status = missing_status_by_column(
        df,
        source_schemas,
        cleaning_audit,
    ).get(column)
    if status is None:
        return 0, 0
    return status.decisions_pending, status.unavailable_from_source


def is_currency_column(column_name: object) -> bool:
    """Return whether a measure name is conventionally currency-like."""
    normalized = _normalise_name(column_name)
    return any(token in normalized for token in CURRENCY_NAME_TOKENS)


def format_business_value(value: Any, *, currency: bool = False) -> str:
    """Format a client-facing value without exposing pandas null markers."""
    if value is None:
        return "—"
    try:
        if bool(pd.isna(value)):
            return "—"
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        return value.strftime("%b %d, %Y")
    if isinstance(value, (int, float)):
        if currency:
            return f"${float(value):,.2f}"
        if isinstance(value, float) and value.is_integer():
            return f"{int(value):,}"
        return f"{float(value):,.2f}"
    return str(value)


def insight_interpretation(
    df: pd.DataFrame,
    profile: InsightColumn,
    *,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    cleaning_audit: Iterable[CleaningAuditEntry | Mapping[str, Any]] = (),
) -> str:
    """Return a concise factual description for one selected column."""
    series = df[profile.name]
    non_null = series.dropna()
    status = missing_status_by_column(
        df,
        source_schemas,
        cleaning_audit,
    ).get(profile.name)
    pending = status.decisions_pending if status else 0
    approved = status.approved_blank if status else 0
    unavailable = status.unavailable_from_source if status else 0
    status_parts = []
    if approved:
        status_parts.append(f"{approved:,} approved blank")
    if unavailable:
        status_parts.append(f"{unavailable:,} unavailable from source")
    if pending:
        status_parts.append(f"{pending:,} pending review")
    blank_suffix = (
        f" Blank status: {', '.join(status_parts)}."
        if status_parts
        else ""
    )

    if profile.semantic_type == SEMANTIC_IDENTIFIER:
        unique_count = int(non_null.nunique(dropna=True))
        repeated_count = max(0, len(non_null) - unique_count)
        repeated_text = (
            "no repeated values"
            if repeated_count == 0
            else f"{repeated_count:,} repeated "
            f"{'value' if repeated_count == 1 else 'values'}"
        )
        return (
            f"{profile.display_name} contains {unique_count:,} unique values "
            f"and {repeated_text}.{blank_suffix}"
        )

    if profile.semantic_type == SEMANTIC_NUMERIC:
        numeric = pd.to_numeric(non_null, errors="coerce").dropna()
        if numeric.empty:
            return f"{profile.display_name} has no numeric values to summarize."
        outlier_count = int(detect_outliers_iqr(series).sum())
        currency = is_currency_column(profile.name)
        review_text = (
            "No values were flagged for review."
            if outlier_count == 0
            else f"{outlier_count:,} "
            f"{'value was' if outlier_count == 1 else 'values were'} "
            "flagged for review."
        )
        return (
            f"{profile.display_name} has a median of "
            f"{format_business_value(numeric.median(), currency=currency)}. "
            f"{review_text}{blank_suffix}"
        )

    if profile.semantic_type == SEMANTIC_DATE:
        parsed = pd.to_datetime(series, errors="coerce").dropna()
        if parsed.empty:
            return f"{profile.display_name} has no valid dates to summarize."
        return (
            f"{profile.display_name} runs from "
            f"{format_business_value(parsed.min())} to "
            f"{format_business_value(parsed.max())} across "
            f"{parsed.dt.normalize().nunique():,} distinct dates.{blank_suffix}"
        )

    if profile.semantic_type == SEMANTIC_CATEGORICAL:
        if non_null.empty:
            return f"{profile.display_name} has no available values to summarize."
        most_common = non_null.mode().iloc[0]
        frequency = int((non_null == most_common).sum())
        return (
            f"{profile.display_name} contains {non_null.nunique():,} categories. "
            f"{most_common} is the most frequent with {frequency:,} "
            f"{'record' if frequency == 1 else 'records'}.{blank_suffix}"
        )

    return f"{profile.display_name} has no supported values to summarize."


def _missing_status_counts(
    df: pd.DataFrame,
    column: object,
    source_schemas: Mapping[str, Iterable[str]] | None,
    cleaning_audit: Iterable[CleaningAuditEntry | Mapping[str, Any]],
) -> tuple[int, int, int]:
    """Return approved, unavailable, and pending blank counts for a column."""
    status = missing_status_by_column(
        df,
        source_schemas,
        cleaning_audit,
    ).get(column)
    if status is None:
        return 0, 0, 0
    return (
        status.approved_blank,
        status.unavailable_from_source,
        status.decisions_pending,
    )


def blank_status_summary(
    df: pd.DataFrame,
    column: object,
    *,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    cleaning_audit: Iterable[CleaningAuditEntry | Mapping[str, Any]] = (),
) -> str:
    """Return a concise client-facing blank-status line."""
    approved, unavailable, pending = _missing_status_counts(
        df,
        column,
        source_schemas,
        cleaning_audit,
    )
    return (
        f"{approved:,} approved to remain blank · "
        f"{unavailable:,} unavailable from source · "
        f"{pending:,} pending decisions"
    )


def selected_column_metrics(
    df: pd.DataFrame,
    profile: InsightColumn,
    *,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    cleaning_audit: Iterable[CleaningAuditEntry | Mapping[str, Any]] = (),
) -> list[tuple[str, str]]:
    """Return only meaningful summary metrics for one selected column."""
    series = df[profile.name]
    non_null = series.dropna()
    _approved, _unavailable, pending = _missing_status_counts(
        df,
        profile.name,
        source_schemas,
        cleaning_audit,
    )

    if profile.semantic_type == SEMANTIC_NUMERIC:
        numeric = pd.to_numeric(non_null, errors="coerce").dropna()
        currency = is_currency_column(profile.name)
        review_flags = int(detect_outliers_iqr(series).sum())
        return [
            (
                "Average",
                format_business_value(
                    numeric.mean() if not numeric.empty else None,
                    currency=currency,
                ),
            ),
            (
                "Median",
                format_business_value(
                    numeric.median() if not numeric.empty else None,
                    currency=currency,
                ),
            ),
            (
                "Minimum",
                format_business_value(
                    numeric.min() if not numeric.empty else None,
                    currency=currency,
                ),
            ),
            (
                "Maximum",
                format_business_value(
                    numeric.max() if not numeric.empty else None,
                    currency=currency,
                ),
            ),
            ("Review flags", f"{review_flags:,}"),
        ]

    if profile.semantic_type == SEMANTIC_IDENTIFIER:
        unique_count = int(non_null.nunique(dropna=True))
        return [
            ("Unique values", f"{unique_count:,}"),
            ("Repeated values", f"{max(0, len(non_null) - unique_count):,}"),
            ("Total records", f"{len(series):,}"),
            ("Pending decisions", f"{pending:,}"),
        ]

    if profile.semantic_type == SEMANTIC_CATEGORICAL:
        mode = non_null.mode()
        most_common = mode.iloc[0] if not mode.empty else None
        frequency = (
            int((non_null == most_common).sum())
            if most_common is not None
            else 0
        )
        return [
            ("Categories", f"{int(non_null.nunique()):,}"),
            ("Most common value", format_business_value(most_common)),
            ("Most common count", f"{frequency:,}"),
            ("Review flags", "0"),
        ]

    if profile.semantic_type == SEMANTIC_DATE:
        parsed = pd.to_datetime(series, errors="coerce")
        valid = parsed.dropna()
        invalid_count = int((series.notna() & parsed.isna()).sum())
        return [
            (
                "Earliest date",
                format_business_value(valid.min() if not valid.empty else None),
            ),
            (
                "Latest date",
                format_business_value(valid.max() if not valid.empty else None),
            ),
            (
                "Distinct dates",
                f"{int(valid.dt.normalize().nunique()):,}",
            ),
            ("Review flags", f"{invalid_count:,}"),
        ]

    return [("Values available", f"{len(non_null):,}")]


def selected_column_technical_details(
    df: pd.DataFrame,
    profile: InsightColumn,
    *,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    cleaning_audit: Iterable[CleaningAuditEntry | Mapping[str, Any]] = (),
) -> list[tuple[str, str]]:
    """Build compact selected-column technical details with friendly labels."""
    series = df[profile.name]
    non_null = series.dropna()
    approved, unavailable, pending = _missing_status_counts(
        df,
        profile.name,
        source_schemas,
        cleaning_audit,
    )
    quality_rows = [
        ("Approved to remain blank", f"{approved:,}"),
        ("Unavailable from source", f"{unavailable:,}"),
        ("Pending decisions", f"{pending:,}"),
    ]

    if profile.semantic_type == SEMANTIC_NUMERIC:
        numeric = pd.to_numeric(non_null, errors="coerce").dropna()
        currency = is_currency_column(profile.name)
        return [
            ("Values available", f"{len(numeric):,}"),
            (
                "Average",
                format_business_value(
                    numeric.mean() if not numeric.empty else None,
                    currency=currency,
                ),
            ),
            (
                "Median",
                format_business_value(
                    numeric.median() if not numeric.empty else None,
                    currency=currency,
                ),
            ),
            (
                "Minimum",
                format_business_value(
                    numeric.min() if not numeric.empty else None,
                    currency=currency,
                ),
            ),
            (
                "Maximum",
                format_business_value(
                    numeric.max() if not numeric.empty else None,
                    currency=currency,
                ),
            ),
            (
                "Standard deviation",
                format_business_value(
                    numeric.std() if len(numeric) > 1 else None,
                    currency=currency,
                ),
            ),
            ("Review flags", f"{int(detect_outliers_iqr(series).sum()):,}"),
            *quality_rows,
        ]

    if profile.semantic_type == SEMANTIC_CATEGORICAL:
        mode = non_null.mode()
        most_common = mode.iloc[0] if not mode.empty else None
        return [
            ("Values available", f"{len(non_null):,}"),
            ("Categories", f"{int(non_null.nunique()):,}"),
            ("Most common value", format_business_value(most_common)),
            (
                "Most common count",
                f"{int((non_null == most_common).sum()) if most_common is not None else 0:,}",
            ),
            *quality_rows,
        ]

    if profile.semantic_type == SEMANTIC_DATE:
        parsed = pd.to_datetime(series, errors="coerce").dropna()
        return [
            ("Values available", f"{len(parsed):,}"),
            ("Distinct dates", f"{int(parsed.dt.normalize().nunique()):,}"),
            (
                "Earliest date",
                format_business_value(parsed.min() if not parsed.empty else None),
            ),
            (
                "Latest date",
                format_business_value(parsed.max() if not parsed.empty else None),
            ),
            *quality_rows,
        ]

    if profile.semantic_type == SEMANTIC_IDENTIFIER:
        unique_count = int(non_null.nunique(dropna=True))
        return [
            ("Total records", f"{len(series):,}"),
            ("Unique values", f"{unique_count:,}"),
            ("Repeated values", f"{max(0, len(non_null) - unique_count):,}"),
            *quality_rows,
        ]

    return [
        ("Values available", f"{len(non_null):,}"),
        *quality_rows,
    ]


def valid_time_columns(
    df: pd.DataFrame,
    profiles: Mapping[object, InsightColumn],
    measure_column: object,
    *,
    minimum_pairs: int = 2,
) -> list[object]:
    """Return date fields with enough complete pairs for a numeric trend."""
    profile = profiles.get(measure_column)
    if profile is None or profile.semantic_type != SEMANTIC_NUMERIC:
        return []
    numeric = pd.to_numeric(df[measure_column], errors="coerce")
    valid = []
    for column, candidate in profiles.items():
        if candidate.semantic_type != SEMANTIC_DATE:
            continue
        dates = pd.to_datetime(df[column], errors="coerce")
        if int((dates.notna() & numeric.notna()).sum()) >= minimum_pairs:
            valid.append(column)
    return valid


def default_time_settings(column_name: object) -> dict[str, str]:
    """Return safe time-view defaults for one numeric measure."""
    normalized = _normalise_name(column_name)
    if normalized == "unit_price" or normalized.endswith("_price"):
        calculation = "Average"
    elif (
        normalized == "quantity"
        or normalized == "total"
        or normalized.endswith("_total")
        or normalized in {"sales", "revenue", "amount"}
    ):
        calculation = "Total"
    else:
        calculation = "Average"
    return {
        "Group by": "Month",
        "Calculation": calculation,
    }


def source_comparison_available(
    df: pd.DataFrame,
    profile: InsightColumn,
    *,
    source_column: str = "source_file",
) -> bool:
    """Return whether a concise comparison can be shown for two source files."""
    return (
        source_column in df.columns
        and int(df[source_column].dropna().nunique()) >= 2
        and profile.name in df.columns
        and profile.name != source_column
        and profile.semantic_type
        in {
            SEMANTIC_NUMERIC,
            SEMANTIC_CATEGORICAL,
            SEMANTIC_DATE,
            SEMANTIC_IDENTIFIER,
        }
    )


def source_file_comparison(
    df: pd.DataFrame,
    profile: InsightColumn,
    *,
    source_column: str = "source_file",
) -> pd.DataFrame:
    """Build one concise, immutable source-file comparison table."""
    if not source_comparison_available(
        df,
        profile,
        source_column=source_column,
    ):
        return pd.DataFrame()

    working = df[[source_column, profile.name]].copy()
    working[source_column] = working[source_column].fillna("—")
    display_name = profile.display_name

    if profile.semantic_type == SEMANTIC_NUMERIC:
        working[profile.name] = pd.to_numeric(
            working[profile.name],
            errors="coerce",
        )
        currency = is_currency_column(profile.name)
        records = []
        for source_name, group in working.groupby(source_column, sort=True):
            values = group[profile.name].dropna()
            records.append({
                "Source File": source_name,
                "Record count": len(values),
                "Average": format_business_value(
                    values.mean() if not values.empty else None,
                    currency=currency,
                ),
                "Median": format_business_value(
                    values.median() if not values.empty else None,
                    currency=currency,
                ),
                "Minimum": format_business_value(
                    values.min() if not values.empty else None,
                    currency=currency,
                ),
                "Maximum": format_business_value(
                    values.max() if not values.empty else None,
                    currency=currency,
                ),
            })
        return pd.DataFrame(records).fillna("—")

    if profile.semantic_type == SEMANTIC_CATEGORICAL:
        counts = (
            working.dropna(subset=[profile.name])
            .groupby([source_column, profile.name], dropna=False)
            .size()
            .rename("Records")
            .reset_index()
            .rename(columns={
                source_column: "Source File",
                profile.name: display_name,
            })
            .sort_values(
                ["Source File", "Records"],
                ascending=[True, False],
            )
            .groupby("Source File", sort=False)
            .head(5)
        )
        return counts.fillna("—").reset_index(drop=True)

    if profile.semantic_type == SEMANTIC_DATE:
        working[profile.name] = pd.to_datetime(
            working[profile.name],
            errors="coerce",
        )
        records = []
        for source_name, group in working.groupby(source_column, sort=True):
            values = group[profile.name].dropna()
            records.append({
                "Source File": source_name,
                "Values available": len(values),
                "Earliest date": format_business_value(
                    values.min() if not values.empty else None
                ),
                "Latest date": format_business_value(
                    values.max() if not values.empty else None
                ),
            })
        return pd.DataFrame(records).fillna("—")

    counts = (
        working.groupby(source_column, dropna=False)
        .size()
        .rename("Record count")
        .reset_index()
        .rename(columns={source_column: "Source File"})
    )
    return counts.fillna("—")


def technical_statistics_by_type(
    df: pd.DataFrame,
    profiles: Mapping[object, InsightColumn],
    *,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    cleaning_audit: Iterable[CleaningAuditEntry | Mapping[str, Any]] = (),
) -> dict[str, pd.DataFrame]:
    """Build narrow all-column technical tables grouped by logical purpose."""
    rows: dict[str, list[dict[str, Any]]] = {
        "Numeric measures": [],
        "Numeric measure quality": [],
        "Categories": [],
        "Category quality": [],
        "Dates": [],
        "Date quality": [],
        "Identifiers": [],
        "Identifier quality": [],
    }
    for column, profile in profiles.items():
        series = df[column]
        non_null = series.dropna()
        approved, unavailable, pending = _missing_status_counts(
            df,
            column,
            source_schemas,
            cleaning_audit,
        )
        base = {"Column": profile.display_name}
        blank_status = {
            "Approved to remain blank": approved,
            "Unavailable from source": unavailable,
            "Pending decisions": pending,
        }

        if profile.semantic_type == SEMANTIC_NUMERIC:
            numeric = pd.to_numeric(non_null, errors="coerce").dropna()
            currency = is_currency_column(column)
            rows["Numeric measures"].append({
                **base,
                "Values available": len(numeric),
                "Average": format_business_value(
                    numeric.mean() if not numeric.empty else None,
                    currency=currency,
                ),
                "Median": format_business_value(
                    numeric.median() if not numeric.empty else None,
                    currency=currency,
                ),
                "Minimum": format_business_value(
                    numeric.min() if not numeric.empty else None,
                    currency=currency,
                ),
                "Maximum": format_business_value(
                    numeric.max() if not numeric.empty else None,
                    currency=currency,
                ),
            })
            rows["Numeric measure quality"].append({
                **base,
                "Standard deviation": format_business_value(
                    numeric.std() if len(numeric) > 1 else None,
                    currency=currency,
                ),
                "Review flags": int(detect_outliers_iqr(series).sum()),
                **blank_status,
            })
        elif profile.semantic_type == SEMANTIC_CATEGORICAL:
            mode = non_null.mode()
            most_common = mode.iloc[0] if not mode.empty else None
            rows["Categories"].append({
                **base,
                "Values available": len(non_null),
                "Categories": int(non_null.nunique()),
                "Most common value": format_business_value(most_common),
                "Most common count": (
                    int((non_null == most_common).sum())
                    if most_common is not None
                    else 0
                ),
            })
            rows["Category quality"].append({
                **base,
                **blank_status,
            })
        elif profile.semantic_type == SEMANTIC_DATE:
            parsed = pd.to_datetime(series, errors="coerce").dropna()
            rows["Dates"].append({
                **base,
                "Values available": len(parsed),
                "Distinct dates": int(parsed.dt.normalize().nunique()),
                "Earliest date": format_business_value(
                    parsed.min() if not parsed.empty else None
                ),
                "Latest date": format_business_value(
                    parsed.max() if not parsed.empty else None
                ),
            })
            rows["Date quality"].append({
                **base,
                **blank_status,
            })
        elif profile.semantic_type == SEMANTIC_IDENTIFIER:
            unique_count = int(non_null.nunique(dropna=True))
            rows["Identifiers"].append({
                **base,
                "Total records": len(series),
                "Unique values": unique_count,
                "Repeated values": max(0, len(non_null) - unique_count),
            })
            rows["Identifier quality"].append({
                **base,
                **blank_status,
            })

    return {
        label: pd.DataFrame(records).fillna("—")
        for label, records in rows.items()
        if records
    }
