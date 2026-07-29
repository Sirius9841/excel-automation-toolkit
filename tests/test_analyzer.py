"""Tests for analyzer.py — summary statistics and outlier detection."""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import pandas as pd
import pytest

from src.analyzer import (
    detect_column_type,
    generate_summary_statistics,
    detect_outliers_iqr,
    summarize_outliers,
)


# ── Fixtures ──────────────────────────────────────────────


@pytest.fixture
def df_mixed_types() -> pd.DataFrame:
    return pd.DataFrame({
        "price": [10.0, 20.0, 30.0, None, 50.0],
        "product": ["A", "B", "C", "D", "E"],
        "order_date": pd.to_datetime(["2024-01-01", "2024-01-02", "2024-01-03",
                                      "2024-01-04", "2024-01-05"]),
        "is_active": [True, False, True, True, False],
        "empty_col": [None] * 5,
    })


@pytest.fixture
def df_with_outliers() -> pd.DataFrame:
    """10 normal values + 2 extreme outliers."""
    return pd.DataFrame({
        "normal": list(range(1, 11)),
        "with_outlier": [10, 12, 11, 13, 9, 100, 11, 12, 10, 14],
    })


# ── Test: column type detection ───────────────────────────


class TestDetectColumnType:

    def test_numeric_column(self, df_mixed_types):
        assert detect_column_type(df_mixed_types["price"]) == "numeric"

    def test_categorical_column(self, df_mixed_types):
        assert detect_column_type(df_mixed_types["product"]) == "categorical"

    def test_date_column(self, df_mixed_types):
        assert detect_column_type(df_mixed_types["order_date"]) == "date"

    def test_boolean_as_categorical(self, df_mixed_types):
        assert detect_column_type(df_mixed_types["is_active"]) == "categorical"

    def test_empty_column(self, df_mixed_types):
        assert detect_column_type(df_mixed_types["empty_col"]) == "empty"

    def test_string_column_matching_date_keyword(self):
        s = pd.Series(["2024-01-01", "2024-01-02", "invalid date", "2024-01-04"],
                      name="timestamp")
        assert detect_column_type(s) == "date"

    def test_string_column_with_date_keyword_but_bad_values(self):
        s = pd.Series(["hello", "world", "foo", "bar"], name="date_column")
        assert detect_column_type(s) == "categorical"

    def test_column_with_all_unique_values(self):
        s = pd.Series([f"id_{i}" for i in range(100)])
        assert detect_column_type(s) == "categorical"

    def test_mixed_type_column(self):
        s = pd.Series([1, "two", 3.0, None, "five"])
        assert detect_column_type(s) == "categorical"


# ── Test: summary statistics ──────────────────────────────


class TestGenerateSummaryStatistics:

    def test_basic_statistics(self, df_mixed_types):
        stats = generate_summary_statistics(df_mixed_types)
        assert len(stats) == 5  # one row per column
        assert "Column" in stats.columns
        assert "Detected Type" in stats.columns
        assert "Non-Missing" in stats.columns
        assert "Missing %" in stats.columns

    def test_numeric_stats_correct(self, df_mixed_types):
        stats = generate_summary_statistics(df_mixed_types)
        price_stats = stats[stats["Column"] == "price"].iloc[0]
        assert price_stats["Mean"] == 27.5  # (10+20+30+50)/4
        assert price_stats["Median"] == 25.0  # sorted [10,20,30,50] -> (20+30)/2
        assert price_stats["Min"] == 10.0
        assert price_stats["Max"] == 50.0

    def test_empty_dataframe(self):
        stats = generate_summary_statistics(pd.DataFrame())
        assert stats.empty

    def test_zero_rows(self):
        df = pd.DataFrame({"a": pd.Series(dtype="float64"),
                           "b": pd.Series(dtype="object")})
        stats = generate_summary_statistics(df)
        assert len(stats) == 2
        assert stats["Detected Type"].iloc[0] == "empty"

    def test_single_row(self):
        df = pd.DataFrame({"val": [42.0]})
        stats = generate_summary_statistics(df)
        row = stats.iloc[0]
        assert row["Mean"] == 42.0
        assert row["Median"] == 42.0
        assert row["Std"] is None  # can't compute std with 1 row

    def test_column_where_every_value_is_unique(self):
        df = pd.DataFrame({"ids": [f"id_{i}" for i in range(50)]})
        stats = generate_summary_statistics(df)
        row = stats.iloc[0]
        assert row["Unique"] == 50
        assert row["Most Common"] == "— (all unique)"
        assert row["Freq"] == 1

    def test_date_column_stats(self, df_mixed_types):
        stats = generate_summary_statistics(df_mixed_types)
        date_stats = stats[stats["Column"] == "order_date"].iloc[0]
        assert "Earliest" in date_stats.index
        assert "Latest" in date_stats.index


# ── Test: outlier detection ───────────────────────────────


class TestDetectOutliersIQR:

    def test_outliers_detected(self, df_with_outliers):
        mask = detect_outliers_iqr(df_with_outliers["with_outlier"])
        assert mask.sum() == 1  # only 100 is an outlier

    def test_no_outliers_in_normal_data(self, df_with_outliers):
        mask = detect_outliers_iqr(df_with_outliers["normal"])
        assert mask.sum() == 0

    def test_non_numeric_column_returns_false(self):
        s = pd.Series(["a", "b", "c"])
        mask = detect_outliers_iqr(s)
        assert mask.sum() == 0

    def test_too_few_values(self):
        s = pd.Series([1.0, 2.0, 3.0])
        mask = detect_outliers_iqr(s)
        assert mask.sum() == 0  # < 4 values

    def test_zero_iqr_all_identical(self):
        s = pd.Series([5.0, 5.0, 5.0, 5.0, 5.0])
        mask = detect_outliers_iqr(s)
        assert mask.sum() == 0

    def test_zero_iqr_with_extreme_value(self):
        s = pd.Series([10.0, 10.0, 10.0, 10.0, 100.0])
        mask = detect_outliers_iqr(s)
        assert mask.sum() == 1  # 100 should be flagged

    def test_all_nan_column(self):
        s = pd.Series([None, None, None, None])
        mask = detect_outliers_iqr(s)
        assert mask.sum() == 0


class TestSummarizeOutliers:

    def test_summarize_returns_outlier_rows(self, df_with_outliers):
        result = summarize_outliers(df_with_outliers, ["with_outlier"])
        assert len(result) == 1
        assert result.iloc[0]["Column"] == "with_outlier"
        assert result.iloc[0]["Value"] == 100

    def test_summarize_no_outliers(self, df_with_outliers):
        result = summarize_outliers(df_with_outliers, ["normal"])
        assert result.empty

    def test_summarize_multiple_columns(self, df_with_outliers):
        result = summarize_outliers(df_with_outliers, ["normal", "with_outlier"])
        assert len(result) == 1  # only one outlier across both columns

    def test_summarize_nonexistent_column(self, df_with_outliers):
        result = summarize_outliers(df_with_outliers, ["not_a_column"])
        assert result.empty

    def test_summarize_empty_selection(self, df_with_outliers):
        result = summarize_outliers(df_with_outliers, [])
        assert result.empty
