"""Merge, clean, and transform DataFrames."""

import pandas as pd

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
    cleaned = df.drop_duplicates(subset=subset, keep=keep)
    removed = before - len(cleaned)

    report = {
        "before": before,
        "after": len(cleaned),
        "removed": removed,
        "subset": subset,
    }

    logger.info("Removed %d duplicate rows (%d -> %d)", removed, before, len(cleaned))
    return cleaned, report


def detect_missing_values(df: pd.DataFrame) -> pd.DataFrame:
    """Return a DataFrame summarizing missing values per column.

    Args:
        df: Source DataFrame.

    Returns:
        DataFrame with columns: column, missing_count, missing_pct, dtype.
    """
    records = []
    for col in df.columns:
        missing = int(df[col].isna().sum())
        records.append({
            "column": col,
            "missing_count": missing,
            "missing_pct": round(missing / len(df) * 100, 1),
            "dtype": str(df[col].dtype),
        })
    result = pd.DataFrame(records)
    result = result.sort_values("missing_count", ascending=False)
    return result


def handle_missing_values(
    df: pd.DataFrame,
    strategies: dict[str, str],
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
    cleaned = df.copy()
    actions: list[str] = []

    # ── Phase 1: Fill strategies (non-destructive) ───────────
    fill_strategies = {c: s for c, s in strategies.items() if s != "drop_rows"}
    drop_columns = [c for c, s in strategies.items() if s == "drop_rows"]

    for col, strategy in fill_strategies.items():
        if col not in cleaned.columns:
            continue

        missing_before = int(cleaned[col].isna().sum())
        if missing_before == 0:
            continue

        if strategy == "fill_mean":
            val = cleaned[col].mean()
            cleaned[col] = cleaned[col].fillna(val)

        elif strategy == "fill_median":
            val = cleaned[col].median()
            cleaned[col] = cleaned[col].fillna(val)

        elif strategy == "fill_mode":
            mode_vals = cleaned[col].mode(dropna=True)
            val = mode_vals.iloc[0] if not mode_vals.empty else None
            cleaned[col] = cleaned[col].fillna(val)

        elif strategy == "fill_zero":
            cleaned[col] = cleaned[col].fillna(0)

        elif strategy.startswith("fill_value:"):
            custom_val = strategy.split(":", 1)[1]
            cleaned[col] = cleaned[col].fillna(custom_val)

        missing_after = int(cleaned[col].isna().sum())
        filled = missing_before - missing_after
        actions.append(f"Column '{col}': filled {filled} missing values ({strategy})")

        logger.info(
            "Missing values in '%s': %d -> %d (strategy: %s)",
            col, missing_before, missing_after, strategy,
        )

    # ── Phase 2: Drop rows (destructive, applied last) ──────
    if drop_columns:
        before_drop = len(cleaned)
        cleaned = cleaned.dropna(subset=drop_columns)
        dropped = before_drop - len(cleaned)
        actions.append(
            f"Dropped {dropped} rows with missing values in: {drop_columns}"
        )
        logger.info(
            "Dropped %d rows with missing values in %s",
            dropped, drop_columns,
        )

    return cleaned, actions


def generate_cleaning_report(
    original: pd.DataFrame,
    cleaned: pd.DataFrame,
    duplicate_report: dict | None = None,
    missing_actions: list[str] | None = None,
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
    return {
        "rows_before": len(original),
        "rows_after": len(cleaned),
        "columns": list(cleaned.columns),
        "duplicates_removed": duplicate_report["removed"] if duplicate_report else 0,
        "missing_actions": missing_actions or [],
        "total_missing_before": int(original.isna().sum().sum()),
        "total_missing_after": int(cleaned.isna().sum().sum()),
    }
