"""Small, side-effect-free helpers for Streamlit UI defaults."""

import pandas as pd

from src.insights import classify_column


DEFAULT_DUPLICATE_COLUMNS: tuple[str, ...] = ()
IDENTIFIER_UNIQUENESS_THRESHOLD = 0.95


def duplicate_subset(selected_columns: list[str]) -> list[str] | None:
    """Return a custom duplicate subset, or None for complete-row matching."""
    return selected_columns or None


def classify_insight_column(
    column_name: object,
    series: pd.Series,
    *,
    uniqueness_threshold: float = IDENTIFIER_UNIQUENESS_THRESHOLD,
) -> str:
    """Classify a column for a useful, non-misleading insight presentation."""
    return classify_column(
        column_name,
        series,
        uniqueness_threshold=uniqueness_threshold,
    ).semantic_type
