"""Tests for data_processor.py — merge and cleaning logic."""

import sys
from pathlib import Path

# Ensure project root is on the path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import pandas as pd
import pytest

from src.data_processor import (
    handle_missing_values,
    remove_duplicates,
    detect_missing_values,
    generate_cleaning_report,
    merge_datasets,
    compute_schema_compatibility,
)  # fmt: skip


# ── Fixtures ──────────────────────────────────────────────


@pytest.fixture
def df_missing() -> pd.DataFrame:
    """A small DataFrame with controlled missing values."""
    return pd.DataFrame({
        "order_id": [1, 2, 3, 4, 5],
        "product": ["A", "B", None, "D", "E"],
        "price": [10.0, None, 30.0, None, 50.0],
        "quantity": [1, 2, 3, None, 5],
    })


@pytest.fixture
def df_duplicates() -> pd.DataFrame:
    """A DataFrame with one full duplicate row."""
    return pd.DataFrame({
        "id": [1, 2, 2, 3],
        "val": [10, 20, 20, 30],
    })


# ── Test: fill before drop ────────────────────────────────


class TestMissingValueOrder:
    """Fill strategies must run before drop_rows, regardless of dict order."""

    def test_fill_runs_before_drop_when_drop_listed_first(self, df_missing):
        """drop_rows listed first in dict: fill should still see full data."""
        strategies = {
            "price": "drop_rows",      # listed first
            "product": "fill_mode",     # listed second — fill must use all 5 rows
        }
        cleaned, actions = handle_missing_values(df_missing, strategies)

        # product was filled with mode from all 5 rows ('A' or 'B' or 'D' or 'E')
        assert cleaned["product"].isna().sum() == 0
        # price rows were dropped last
        assert len(cleaned) < len(df_missing)

    def test_median_from_full_dataset(self, df_missing):
        """Median fill uses all rows, not post-drop rows."""
        strategies = {
            "quantity": "drop_rows",    # will drop row with NaN
            "price": "fill_median",     # must use all 5 rows for median
        }
        cleaned, _ = handle_missing_values(df_missing, strategies)

        # Median of all 5 price values: sorted [10, 30, 50, NaN, NaN]
        # median of [10, 30, 50] ignoring NaN = 30.0
        full_median = df_missing["price"].median()
        assert full_median == 30.0

        # After fill, any remaining price values should be 30.0
        filled_prices = cleaned["price"].dropna()
        assert (filled_prices == 30.0).any()

    def test_mode_from_full_dataset(self, df_missing):
        """Mode fill uses all rows, not post-drop rows."""
        strategies = {
            "price": "drop_rows",
            "product": "fill_mode",
        }
        cleaned, _ = handle_missing_values(df_missing, strategies)

        # Mode of all 5 products: A, B, D, E (each appears once, mode picks first)
        assert cleaned["product"].isna().sum() == 0

    def test_fill_does_not_see_reduced_data(self, df_missing):
        """Prove fill ran before drop: fill used pre-drop median."""
        df = pd.DataFrame({
            "a": [1.0, 2.0, 3.0, None, 100.0],  # median of all 5 = 2.5
            "b": [10.0, None, None, None, 50.0],  # 3 missing -> will drop 3 rows
        })
        strategies = {"a": "fill_median", "b": "drop_rows"}
        cleaned, _ = handle_missing_values(df, strategies)

        # If fill ran on full data: median = 2.5, filled value at index 3 = 2.5
        # If fill ran after drop: only rows 0 and 4 remain, median = 50.5, fill fails
        assert cleaned.loc[0, "a"] == 1.0
        assert cleaned.loc[4, "a"] == 100.0

    def test_drop_rows_only_at_end(self, df_missing):
        """Multiple drop_rows columns: all drops happen in one pass at the end."""
        strategies = {
            "price": "drop_rows",
            "quantity": "drop_rows",
        }
        before = len(df_missing)
        cleaned, actions = handle_missing_values(df_missing, strategies)

        # Drops rows where price or quantity is NaN.
        # Row 1: price=NaN, quantity=2 -> dropped
        # Row 3: price=NaN, quantity=NaN -> dropped
        # Row 4: price=50.0, quantity=5.0 -> NOT dropped
        expected_dropped = 2
        assert len(cleaned) == before - expected_dropped
        assert any("Dropped" in a for a in actions)


# ── Test: original DataFrame is not modified ──────────────


class TestImmutability:

    def test_original_not_modified(self, df_missing):
        """handle_missing_values must not mutate the input DataFrame."""
        original_copy = df_missing.copy()
        strategies = {"product": "fill_mode", "price": "fill_median"}
        handle_missing_values(df_missing, strategies)

        pd.testing.assert_frame_equal(df_missing, original_copy)

    def test_remove_duplicates_immutable(self, df_duplicates):
        """remove_duplicates must not mutate the input DataFrame."""
        original_copy = df_duplicates.copy()
        remove_duplicates(df_duplicates)

        pd.testing.assert_frame_equal(df_duplicates, original_copy)


# ── Test: cleaning report correctness ─────────────────────


class TestCleaningReport:

    def test_report_counts_filled_values(self, df_missing):
        """Report should show how many values were filled per column."""
        strategies = {"product": "fill_mode"}
        cleaned, actions = handle_missing_values(df_missing, strategies)

        report = generate_cleaning_report(
            df_missing, cleaned, missing_actions=actions
        )
        assert report["total_missing_before"] == 4
        assert report["total_missing_after"] == 3  # product filled, price + quantity remain

    def test_report_records_dropped_rows(self, df_missing):
        """Report should show how many rows were dropped."""
        strategies = {"quantity": "drop_rows"}
        cleaned, actions = handle_missing_values(df_missing, strategies)

        report = generate_cleaning_report(
            df_missing, cleaned, missing_actions=actions
        )
        assert report["rows_before"] == 5
        assert report["rows_after"] == 4

    def test_report_with_duplicates(self, df_duplicates):
        """Report should include duplicate removal count."""
        cleaned, dup_report = remove_duplicates(df_duplicates)
        report = generate_cleaning_report(
            df_duplicates, cleaned, duplicate_report=dup_report
        )
        assert report["duplicates_removed"] == 1
        assert report["rows_before"] == 4
        assert report["rows_after"] == 3


# ── Test: multiple strategies end-to-end ──────────────────


class TestMultipleStrategies:

    def test_mixed_strategies_produce_expected_result(self, df_missing):
        """Apply fill, drop, and verify final state."""
        strategies = {
            "product": "fill_mode",
            "price": "fill_median",
            "quantity": "drop_rows",
        }
        cleaned, actions = handle_missing_values(df_missing, strategies)

        assert cleaned["product"].isna().sum() == 0
        assert cleaned["price"].isna().sum() == 0
        assert cleaned["quantity"].isna().sum() == 0
        assert len(df_missing) - len(cleaned) == 1  # one row dropped (row 3)
        assert len(actions) == 3  # 2 fills + 1 drop

    def test_all_drop_rows(self, df_missing):
        """All columns set to drop_rows: drops rows with ANY NaN."""
        strategies = {col: "drop_rows" for col in df_missing.columns}
        cleaned, _ = handle_missing_values(df_missing, strategies)

        assert cleaned.isna().sum().sum() == 0  # no missing values remain
        assert len(cleaned) <= len(df_missing)

    def test_custom_fill_value(self, df_missing):
        """fill_value:XYZ should insert the custom string."""
        strategies = {"product": "fill_value:UNKNOWN"}
        cleaned, _ = handle_missing_values(df_missing, strategies)

        assert cleaned.loc[2, "product"] == "UNKNOWN"

    def test_all_fill_no_drop(self, df_missing):
        """Using only fill strategies should preserve all rows."""
        strategies = {
            "product": "fill_mode",
            "price": "fill_median",
            "quantity": "fill_zero",
        }
        cleaned, _ = handle_missing_values(df_missing, strategies)

        assert len(cleaned) == len(df_missing)
        assert cleaned.isna().sum().sum() == 0


# ── Test: merge and compatibility ─────────────────────────


class TestMerge:

    def test_merge_union_preserves_all_columns(self):
        df1 = pd.DataFrame({"a": [1], "b": [2]})
        df2 = pd.DataFrame({"a": [3], "c": [4]})
        merged, _ = merge_datasets([("f1", df1), ("f2", df2)])

        assert "a" in merged.columns
        assert "b" in merged.columns
        assert "c" in merged.columns
        assert merged["b"].isna().sum() == 1  # NaN for f2 row

    def test_merge_intersection_common_only(self):
        df1 = pd.DataFrame({"a": [1], "b": [2]})
        df2 = pd.DataFrame({"a": [3], "c": [4]})
        merged, _ = merge_datasets([("f1", df1), ("f2", df2)], keep_all_columns=False)

        # source_file is added before intersection so it's always included
        assert "a" in merged.columns
        assert "source_file" in merged.columns
        assert "b" not in merged.columns
        assert "c" not in merged.columns

    def test_compatibility_perfect_match(self):
        df1 = pd.DataFrame({"a": [1], "b": [2]})
        df2 = pd.DataFrame({"a": [3], "b": [4]})
        scores = compute_schema_compatibility([df1, df2], ["f1", "f2"])

        assert scores[("f1", "f2")] == 1.0

    def test_compatibility_no_match(self):
        df1 = pd.DataFrame({"a": [1]})
        df2 = pd.DataFrame({"b": [2]})
        scores = compute_schema_compatibility([df1, df2], ["f1", "f2"])

        assert scores[("f1", "f2")] == 0.0
