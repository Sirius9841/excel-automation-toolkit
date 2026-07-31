"""Tests for safe duplicate defaults used by the Streamlit UI."""

import pandas as pd

from src.data_processor import remove_duplicates
from src.ui_helpers import (
    DEFAULT_DUPLICATE_COLUMNS,
    classify_insight_column,
    duplicate_subset,
)


def test_no_duplicate_columns_are_selected_by_default():
    assert DEFAULT_DUPLICATE_COLUMNS == ()


def test_empty_selection_uses_complete_row_matching():
    assert duplicate_subset([]) is None

    df = pd.DataFrame({
        "order_id": [1, 1, 2],
        "product": ["Widget", "Widget", "Widget"],
        "amount": [10.0, 10.0, 20.0],
    })
    cleaned, report = remove_duplicates(df, subset=duplicate_subset([]))

    assert len(cleaned) == 2
    assert report["removed"] == 1


def test_repeated_products_are_preserved_by_default():
    df = pd.DataFrame({
        "order_id": [1, 2, 3],
        "product": ["Widget", "Widget", "Widget"],
        "amount": [10.0, 20.0, 30.0],
    })
    cleaned, report = remove_duplicates(df, subset=duplicate_subset([]))

    pd.testing.assert_frame_equal(cleaned, df)
    assert report["removed"] == 0


def test_identifier_name_takes_priority_over_numeric_metrics():
    series = pd.Series([1001, 1002, 1003, 1004])

    assert classify_insight_column("order_id", series) == "identifier"


def test_almost_unique_column_is_treated_as_identifier():
    series = pd.Series([f"record-{index}" for index in range(19)] + ["record-18"])

    assert classify_insight_column("reference", series) == "identifier"


def test_repeated_numeric_measure_remains_numeric():
    series = pd.Series([10, 10, 20, 20, 30, 30])

    assert classify_insight_column("amount", series) == "numeric"


def test_datetime_column_is_not_misclassified_as_identifier():
    series = pd.Series(pd.date_range("2026-01-01", periods=10))

    assert classify_insight_column("event_id", series) == "date"


def test_repeated_text_column_is_categorical():
    series = pd.Series(["North", "South", "North", "West"])

    assert classify_insight_column("region", series) == "categorical"
