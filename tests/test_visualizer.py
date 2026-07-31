"""Tests for visualizer.py — chart generation."""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd
import pytest
from matplotlib.figure import Figure

from src.visualizer import (
    plot_histogram,
    plot_boxplot,
    plot_bar_chart,
    plot_line_chart,
)


@pytest.fixture(autouse=True)
def close_figures_after_test():
    """Keep chart tests isolated and avoid retaining pyplot figures."""
    yield
    plt.close("all")


@pytest.fixture
def df_numeric() -> pd.DataFrame:
    return pd.DataFrame({"price": [10, 20, 30, 40, 50, 100]})


@pytest.fixture
def df_categorical() -> pd.DataFrame:
    categories = ["A", "A", "B", "B", "B", "C", "D", "E", "F", "G"]
    return pd.DataFrame({"product": categories})


@pytest.fixture
def df_dates() -> pd.DataFrame:
    return pd.DataFrame({
        "date": pd.to_datetime(["2024-01-01", "2024-01-15", "2024-02-01", "2024-02-15",
                                "2024-03-01", "2024-03-15"]),
        "sales": [100, 150, 130, 170, 160, 200],
    })


# ── Test: all chart functions return a Figure ─────────────


class TestAllCharts:

    def test_histogram_returns_figure(self, df_numeric):
        fig = plot_histogram(df_numeric, "price")
        assert isinstance(fig, Figure)

    def test_boxplot_returns_figure(self, df_numeric):
        fig = plot_boxplot(df_numeric, "price")
        assert isinstance(fig, Figure)

    def test_bar_chart_returns_figure(self, df_categorical):
        fig = plot_bar_chart(df_categorical, "product")
        assert isinstance(fig, Figure)

    def test_line_chart_returns_figure(self, df_dates):
        fig = plot_line_chart(df_dates, "date", "sales")
        assert isinstance(fig, Figure)


# ── Test: histogram ───────────────────────────────────────


class TestHistogram:

    def test_bins_parameter(self, df_numeric):
        fig = plot_histogram(df_numeric, "price", bins=5)
        assert isinstance(fig, Figure)

    def test_empty_column_returns_figure(self, df_numeric):
        empty = df_numeric.copy()
        empty["price"] = None
        fig = plot_histogram(empty, "price")
        assert isinstance(fig, Figure)

    def test_dataframe_not_modified(self, df_numeric):
        original = df_numeric.copy()
        plot_histogram(df_numeric, "price")
        pd.testing.assert_frame_equal(df_numeric, original)


# ── Test: bar chart ──────────────────────────────────────


class TestBarChart:

    def test_top_n_limits_categories(self, df_categorical):
        fig = plot_bar_chart(df_categorical, "product", top_n=3)
        assert isinstance(fig, Figure)

    def test_high_cardinality_shows_top_n(self):
        many_cats = {f"cat_{i}": [f"val_{j}" for j in range(100)] for i in range(1)}
        df = pd.DataFrame(many_cats)
        fig = plot_bar_chart(df, "cat_0", top_n=5)
        assert isinstance(fig, Figure)

    def test_empty_column_returns_figure(self, df_categorical):
        empty = df_categorical.copy()
        empty["product"] = None
        fig = plot_bar_chart(empty, "product")
        assert isinstance(fig, Figure)


# ── Test: line chart ─────────────────────────────────────


class TestLineChart:

    def test_sum_aggregation(self, df_dates):
        fig = plot_line_chart(df_dates, "date", "sales", agg="sum")
        assert isinstance(fig, Figure)

    def test_mean_aggregation(self, df_dates):
        fig = plot_line_chart(df_dates, "date", "sales", agg="mean")
        assert isinstance(fig, Figure)

    def test_clear_title_for_total_quantity_by_month(self, df_dates):
        renamed = df_dates.rename(columns={"sales": "quantity"})

        fig = plot_line_chart(
            renamed,
            "date",
            "quantity",
            agg="sum",
            group_by="Month",
        )

        assert fig.axes[0].get_title(loc="left") == "Total Quantity by Month"

    def test_average_unit_price_title_and_week_grouping(self, df_dates):
        renamed = df_dates.rename(columns={"sales": "unit_price"})

        fig = plot_line_chart(
            renamed,
            "date",
            "unit_price",
            agg="mean",
            group_by="Week",
        )

        assert fig.axes[0].get_title(loc="left") == (
            "Average Unit Price by Week"
        )

    def test_sales_total_uses_client_facing_title(self, df_dates):
        renamed = df_dates.rename(columns={"sales": "total"})

        fig = plot_line_chart(
            renamed,
            "date",
            "total",
            agg="sum",
            group_by="Month",
        )

        assert fig.axes[0].get_title(loc="left") == "Sales Total by Month"

    def test_string_date_conversion(self, df_dates):
        df_str = df_dates.copy()
        df_str["date"] = df_str["date"].astype(str)
        fig = plot_line_chart(df_str, "date", "sales")
        assert isinstance(fig, Figure)

    def test_empty_data_returns_figure(self):
        df = pd.DataFrame({"d": pd.Series(dtype="datetime64[ns]"),
                           "v": pd.Series(dtype="float64")})
        fig = plot_line_chart(df, "d", "v")
        assert isinstance(fig, Figure)


# ── Test: immutability ───────────────────────────────────


class TestImmutability:

    def test_none_modified(self, df_numeric, df_categorical, df_dates):
        orig_num = df_numeric.copy()
        orig_cat = df_categorical.copy()
        orig_dates = df_dates.copy()

        plot_histogram(df_numeric, "price")
        plot_bar_chart(df_categorical, "product")
        plot_line_chart(df_dates, "date", "sales")

        pd.testing.assert_frame_equal(df_numeric, orig_num)
        pd.testing.assert_frame_equal(df_categorical, orig_cat)
        pd.testing.assert_frame_equal(df_dates, orig_dates)
