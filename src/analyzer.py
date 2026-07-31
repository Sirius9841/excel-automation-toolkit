"""Generate summary statistics and detect outliers."""

from collections.abc import Iterable, Mapping
import re
from typing import Any

import pandas as pd

from src.data_quality import CleaningAuditEntry, missing_status_by_column
from src.logger_setup import setup_logger

logger = setup_logger(__name__)

# ── Column classification ─────────────────────────────────


def detect_column_type(series: pd.Series) -> str:
    """Classify a column into one of: numeric, categorical, date, or empty.

    Args:
        series: A pandas Series.

    Returns:
        One of 'numeric', 'categorical', 'date', or 'empty'.
    """
    if series.isna().all():
        return "empty"

    dtype = series.dtype

    if pd.api.types.is_bool_dtype(dtype):
        return "categorical"

    if pd.api.types.is_numeric_dtype(dtype):
        return "numeric"

    if pd.api.types.is_datetime64_any_dtype(dtype):
        return "date"

    # For object/string columns, try a conservative date check
    if pd.api.types.is_object_dtype(dtype) or pd.api.types.is_string_dtype(dtype):
        non_null = series.dropna()
        if len(non_null) == 0:
            return "empty"

        # Only attempt date parsing if column name suggests a date
        name_lower = str(series.name).lower()
        date_keywords = ["date", "time", "timestamp", "day", "month", "year",
                         "created", "updated", "datetime"]
        has_date_name = any(kw in name_lower for kw in date_keywords)

        if has_date_name:
            try:
                sample = non_null.head(20)
                parsed = pd.to_datetime(sample, errors="coerce")
                success_rate = parsed.notna().mean()
                if success_rate >= 0.7:
                    return "date"
            except (ValueError, TypeError):
                pass

        return "categorical"

    return "categorical"


def friendly_column_name(column: object) -> str:
    """Convert an internal column name into a readable display label."""
    text = re.sub(r"[_\-]+", " ", str(column)).strip()
    words = []
    for word in text.split():
        words.append("ID" if word.lower() == "id" else word.capitalize())
    return " ".join(words) or str(column)


def detect_business_column_type(series: pd.Series) -> str:
    """Classify a column for business reporting, including identifiers."""
    basic_type = detect_column_type(series)
    if basic_type in {"empty", "date"}:
        return basic_type

    name = str(series.name).strip().lower()
    id_like_name = name == "id" or name.endswith("_id")
    non_null = series.dropna()
    near_unique = (
        len(non_null) > 0
        and non_null.nunique(dropna=True) / len(non_null) >= 0.9
    )
    if id_like_name and near_unique:
        return "identifier"
    return basic_type


# ── Summary statistics ────────────────────────────────────


def generate_summary_statistics(
    df: pd.DataFrame,
    *,
    source_schemas: Mapping[str, Iterable[str]] | None = None,
    cleaning_audit: Iterable[CleaningAuditEntry | Mapping[str, Any]] = (),
) -> pd.DataFrame:
    """Return a DataFrame of per-column summary statistics.

    Each row is one column from the input. Columns shown depend on the
    detected data type.

    Args:
        df: A pandas DataFrame (may be empty or have zero rows).

    Returns:
        DataFrame with columns determined by data type.
        Returns empty DataFrame if input is empty or has no columns.
    """
    if len(df.columns) == 0:
        return pd.DataFrame()

    statuses = missing_status_by_column(
        df,
        source_schemas,
        cleaning_audit,
    )
    records = []
    for col in df.columns:
        series = df[col]
        col_type = detect_column_type(series)
        status = statuses.get(col)

        row: dict[str, object] = {
            "Column": col,
            "Detected Type": col_type,
            "Non-Missing": int(series.notna().sum()),
            "Missing": int(series.isna().sum()),
            "Missing %": round(series.isna().mean() * 100, 1),
            "Approved blank": status.approved_blank if status else 0,
            "Unavailable from source": (
                status.unavailable_from_source if status else 0
            ),
            "Decisions pending": status.decisions_pending if status else 0,
            "Unique": int(series.nunique(dropna=False)),
        }

        if col_type == "numeric":
            row["Mean"] = round(series.mean(), 2) if series.notna().any() else None
            row["Median"] = round(series.median(), 2) if series.notna().any() else None
            row["Min"] = round(series.min(), 2) if series.notna().any() else None
            row["Max"] = round(series.max(), 2) if series.notna().any() else None
            row["Std"] = round(series.std(), 2) if series.notna().sum() > 1 else None

        elif col_type == "date":
            non_null = series.dropna()
            if not non_null.empty:
                row["Earliest"] = non_null.min()
                row["Latest"] = non_null.max()
            else:
                row["Earliest"] = None
                row["Latest"] = None

        elif col_type == "categorical":
            non_null = series.dropna()
            if not non_null.empty:
                all_unique = non_null.nunique() == len(non_null)
                if all_unique:
                    row["Most Common"] = "— (all unique)"
                    row["Freq"] = 1
                else:
                    mode_vals = non_null.mode()
                    row["Most Common"] = str(mode_vals.iloc[0]) if not mode_vals.empty else None
                    row["Freq"] = int((non_null == mode_vals.iloc[0]).sum()) if not mode_vals.empty else 0
            else:
                row["Most Common"] = None
                row["Freq"] = 0

        # empty columns have no extra stats beyond the base
        records.append(row)

    result = pd.DataFrame(records)
    logger.info("Generated summary statistics for %d columns", len(result))
    return result


# ── Outlier detection ─────────────────────────────────────


def detect_outliers_iqr(series: pd.Series) -> pd.Series:
    """Return a boolean mask where True indicates a statistical outlier.

    Uses the IQR (Interquartile Range) method:
        outlier = value < Q1 - 1.5*IQR  OR  value > Q3 + 1.5*IQR

    Args:
        series: A numeric pandas Series.

    Returns:
        Boolean Series aligned with the input. True = outlier.
        Returns all-False if the column has too few non-null values (< 4).
    """
    if not pd.api.types.is_numeric_dtype(series):
        logger.warning("Column '%s' is not numeric; cannot detect outliers", series.name)
        return pd.Series(False, index=series.index)

    non_null = series.dropna()
    if len(non_null) < 4:
        logger.warning(
            "Column '%s' has only %d non-null values; too few for IQR outlier detection",
            series.name, len(non_null),
        )
        return pd.Series(False, index=series.index)

    q1 = non_null.quantile(0.25)
    q3 = non_null.quantile(0.75)
    iqr = q3 - q1

    lower = q1 - 1.5 * iqr
    upper = q3 + 1.5 * iqr

    outliers = (series < lower) | (series > upper)
    outlier_count = outliers.sum()
    logger.info("Column '%s': %d outliers detected (IQR method)", series.name, outlier_count)
    return outliers


def summarize_outliers(
    df: pd.DataFrame,
    columns: list[str],
) -> pd.DataFrame:
    """Return a DataFrame listing outlier rows for the specified columns.

    Each row in the output corresponds to a cell that was flagged as an
    outlier. Columns: column_name, row_index, value, q1, q3, iqr.

    Args:
        df: Source DataFrame.
        columns: Numeric column names to check.

    Returns:
        DataFrame with one row per outlier cell.
    """
    records = []
    identifier_columns = [
        column
        for column in df.columns
        if str(column).lower() == "id"
        or str(column).lower().endswith("_id")
    ]
    descriptive_columns = [
        column
        for column in df.columns
        if str(column).lower() in {
            "product",
            "name",
            "description",
            "category",
            "customer",
        }
    ]
    for col in columns:
        if col not in df.columns:
            continue
        series = df[col]
        mask = detect_outliers_iqr(series)
        if mask.any():
            non_null = series.dropna()
            q1 = non_null.quantile(0.25)
            q3 = non_null.quantile(0.75)
            iqr = q3 - q1
            lower = q1 - 1.5 * iqr
            upper = q3 + 1.5 * iqr
            outlier_indices = mask[mask].index
            for idx in outlier_indices:
                identifier = (
                    df.at[idx, identifier_columns[0]]
                    if identifier_columns
                    else f"Row {idx}"
                )
                description = (
                    df.at[idx, descriptive_columns[0]]
                    if descriptive_columns
                    else "—"
                )
                source_file = (
                    df.at[idx, "source_file"]
                    if "source_file" in df.columns
                    else "—"
                )
                records.append({
                    "Column": col,
                    "Row": idx,
                    "Record ID": identifier,
                    "Source File": source_file,
                    "Original Source Row": df.attrs.get(
                        "source_row_numbers", {}
                    ).get(idx, "—"),
                    "Description": description,
                    "Value": series.loc[idx],
                    "Q1": round(q1, 2),
                    "Q3": round(q3, 2),
                    "IQR": round(iqr, 2),
                    "Lower Review Boundary": round(lower, 2),
                    "Upper Review Boundary": round(upper, 2),
                    "Reason": (
                        "Value falls outside the IQR review range; "
                        "it is not automatically incorrect."
                    ),
                })
    return pd.DataFrame(records)
