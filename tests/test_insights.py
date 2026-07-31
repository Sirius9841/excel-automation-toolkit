"""Tests for the client-facing semantic Insights model."""

from pathlib import Path

import pandas as pd

from src.insights import (
    SEMANTIC_CATEGORICAL,
    SEMANTIC_DATE,
    SEMANTIC_IDENTIFIER,
    SEMANTIC_NUMERIC,
    blank_status_summary,
    chart_config_is_compatible,
    classify_column,
    classify_columns,
    compatible_chart_columns,
    default_time_settings,
    default_chart_config,
    grouped_column_options,
    insight_interpretation,
    selected_column_metrics,
    selected_column_technical_details,
    source_comparison_available,
    source_file_comparison,
    technical_statistics_by_type,
    valid_time_columns,
)
from src.visualizer import plot_bar_chart, plot_histogram


APP_SOURCE = Path(__file__).resolve().parents[1] / "app.py"


def insight_frame() -> pd.DataFrame:
    return pd.DataFrame({
        "order_id": [1001, 1002, 1003, 1004],
        "quantity": [1, 2, 2, 4],
        "unit_price": [10.0, 12.5, 10.0, 18.0],
        "region": ["North", "West", "North", None],
        "order_date": pd.to_datetime([
            "2026-01-01",
            "2026-01-02",
            "2026-01-03",
            "2026-01-04",
        ]),
        "source_file": ["north.xlsx", "north.xlsx", "south.xlsx", "south.xlsx"],
    })


def test_order_id_uses_identifier_presentation():
    profiles = classify_columns(insight_frame())

    assert profiles["order_id"].semantic_type == SEMANTIC_IDENTIFIER
    assert profiles["order_id"].display_name == "Order ID"
    assert "averages and distributions are not meaningful" not in (
        insight_interpretation(insight_frame(), profiles["order_id"])
    )


def test_configured_semantic_override_is_authoritative():
    profile = classify_column(
        "account_code",
        pd.Series(["A", "B", "C"]),
        configured_overrides={"account_code": SEMANTIC_IDENTIFIER},
    )

    assert profile.semantic_type == SEMANTIC_IDENTIFIER


def test_identifier_columns_cannot_create_distribution_charts():
    profiles = classify_columns(insight_frame())

    assert "order_id" not in compatible_chart_columns(profiles)["Distribution"]
    assert default_chart_config(profiles["order_id"]) is None
    assert not chart_config_is_compatible(
        {"type": "Distribution", "column": "order_id"},
        profiles,
    )


def test_all_chart_selectors_expose_only_compatible_columns():
    profiles = classify_columns(insight_frame())
    compatible = compatible_chart_columns(profiles)

    assert compatible["Distribution"] == ["quantity", "unit_price"]
    assert compatible["Range and review flags"] == [
        "quantity",
        "unit_price",
    ]
    assert compatible["Category comparison"] == ["region"]
    assert compatible["Trend date"] == ["order_date"]
    assert compatible["Trend value"] == ["quantity", "unit_price"]
    assert "source_file" not in compatible["Category comparison"]


def test_selector_options_are_grouped_by_semantic_type():
    frame = insight_frame()
    profiles = classify_columns(frame)
    options = grouped_column_options(profiles)

    assert options == [
        "quantity",
        "unit_price",
        "region",
        "order_date",
        "order_id",
        "source_file",
    ]
    assert profiles["quantity"].semantic_type == SEMANTIC_NUMERIC
    assert profiles["region"].semantic_type == SEMANTIC_CATEGORICAL
    assert profiles["order_date"].semantic_type == SEMANTIC_DATE


def test_client_facing_chart_titles_use_friendly_labels():
    frame = insight_frame()
    numeric_figure = plot_histogram(frame, "unit_price")
    category_figure = plot_bar_chart(frame, "region")

    numeric_title = numeric_figure.axes[0].get_title(loc="left")
    category_title = category_figure.axes[0].get_title(loc="left")
    assert numeric_title == "Unit Price Distribution"
    assert category_title == (
        "Most Frequent Region Values"
    )
    assert "unit_price" not in numeric_title


def test_technical_statistics_are_split_and_do_not_expose_null_markers():
    frame = insight_frame()
    frame["empty_measure"] = [None] * len(frame)
    profiles = classify_columns(
        frame,
        configured_overrides={"empty_measure": SEMANTIC_NUMERIC},
    )
    tables = technical_statistics_by_type(frame, profiles)

    assert {
        "Numeric measures",
        "Numeric measure quality",
        "Categories",
        "Category quality",
        "Dates",
        "Date quality",
        "Identifiers",
        "Identifier quality",
    }.issubset(tables)
    assert list(tables["Identifiers"]["Column"]) == [
        "Order ID",
        "Source File",
    ]
    assert all(len(table.columns) <= 6 for table in tables.values())
    assert list(tables["Numeric measures"].columns) == [
        "Column",
        "Values available",
        "Average",
        "Median",
        "Minimum",
        "Maximum",
    ]
    assert list(tables["Numeric measure quality"].columns) == [
        "Column",
        "Standard deviation",
        "Review flags",
        "Approved to remain blank",
        "Unavailable from source",
        "Pending decisions",
    ]
    combined_text = " ".join(
        table.astype(str).to_string(index=False)
        for table in tables.values()
    )
    assert "None" not in combined_text
    assert "NaT" not in combined_text
    assert "nan" not in combined_text.lower()
    empty_row = tables["Numeric measures"].loc[
        tables["Numeric measures"]["Column"].eq("Empty Measure")
    ].iloc[0]
    assert empty_row["Average"] == "—"


def test_technical_statistics_separate_blank_decision_states():
    frame = pd.DataFrame({
        "region": ["North", None, "South"],
        "customer_city": ["Paris", None, None],
        "source_file": ["north.xlsx", "north.xlsx", "south.xlsx"],
    })
    schemas = {
        "north.xlsx": frozenset({"region", "customer_city"}),
        "south.xlsx": frozenset({"region"}),
    }
    audit = [
        {
            "action_type": "approved_unchanged",
            "decision_state": "approved_unchanged",
            "column": "region",
            "affected_row_count": 1,
        },
        {
            "action_type": "approved_unchanged",
            "decision_state": "approved_unchanged",
            "column": "customer_city",
            "affected_row_count": 1,
        },
    ]
    profiles = classify_columns(frame)
    tables = technical_statistics_by_type(
        frame,
        profiles,
        source_schemas=schemas,
        cleaning_audit=audit,
    )
    categories = tables["Category quality"].set_index("Column")

    assert categories.loc["Region", "Approved to remain blank"] == 1
    assert categories.loc["Region", "Unavailable from source"] == 0
    assert categories.loc["Region", "Pending decisions"] == 0
    assert categories.loc["Customer City", "Approved to remain blank"] == 1
    assert categories.loc["Customer City", "Unavailable from source"] == 1
    assert categories.loc["Customer City", "Pending decisions"] == 0


def test_interpretations_are_factual_and_friendly():
    frame = insight_frame()
    profiles = classify_columns(frame)

    category_text = insight_interpretation(frame, profiles["region"])
    identifier_text = insight_interpretation(frame, profiles["order_id"])

    assert "Region contains 2 categories" in category_text
    assert "North is the most frequent with 2 records" in category_text
    assert "Order ID contains 4 unique values" in identifier_text
    assert "order_id" not in identifier_text


def test_selected_technical_details_default_to_one_column():
    frame = insight_frame()
    profiles = classify_columns(frame)

    details = selected_column_technical_details(
        frame,
        profiles["quantity"],
    )

    assert details[0] == ("Values available", "4")
    assert ("Average", "2.25") in details
    assert ("Standard deviation", "1.26") in details
    detail_text = " ".join(f"{label} {value}" for label, value in details)
    assert "Unit Price" not in detail_text
    assert "Region" not in detail_text


def test_metric_layout_has_no_placeholder_metrics():
    frame = insight_frame()
    profiles = classify_columns(frame)

    numeric_metrics = selected_column_metrics(frame, profiles["quantity"])
    identifier_metrics = selected_column_metrics(frame, profiles["order_id"])
    category_metrics = selected_column_metrics(frame, profiles["region"])
    date_metrics = selected_column_metrics(frame, profiles["order_date"])

    assert [label for label, _value in numeric_metrics] == [
        "Average",
        "Median",
        "Minimum",
        "Maximum",
        "Review flags",
    ]
    assert [label for label, _value in identifier_metrics] == [
        "Unique values",
        "Repeated values",
        "Total records",
        "Pending decisions",
    ]
    assert [label for label, _value in category_metrics] == [
        "Categories",
        "Most common value",
        "Most common count",
        "Review flags",
    ]
    assert [label for label, _value in date_metrics] == [
        "Earliest date",
        "Latest date",
        "Distinct dates",
        "Review flags",
    ]
    assert all(
        label and value != ""
        for metrics in (
            numeric_metrics,
            identifier_metrics,
            category_metrics,
            date_metrics,
        )
        for label, value in metrics
    )


def test_blank_status_uses_friendly_labels():
    summary = blank_status_summary(insight_frame(), "quantity")

    assert summary == (
        "0 approved to remain blank · "
        "0 unavailable from source · "
        "0 pending decisions"
    )
    assert "Approved blank" not in summary
    assert "Decisions pending" not in summary


def test_time_view_requires_numeric_measure_and_enough_dated_records():
    frame = insight_frame()
    profiles = classify_columns(frame)

    assert valid_time_columns(frame, profiles, "quantity") == ["order_date"]
    assert valid_time_columns(frame, profiles, "region") == []

    sparse = frame.copy()
    sparse.loc[1:, "order_date"] = pd.NaT
    sparse_profiles = classify_columns(sparse)
    assert valid_time_columns(sparse, sparse_profiles, "quantity") == []


def test_time_defaults_match_measure_semantics():
    assert default_time_settings("quantity") == {
        "Group by": "Month",
        "Calculation": "Total",
    }
    assert default_time_settings("unit_price") == {
        "Group by": "Month",
        "Calculation": "Average",
    }
    assert default_time_settings("total") == {
        "Group by": "Month",
        "Calculation": "Total",
    }


def test_source_comparison_requires_two_source_files():
    frame = insight_frame()
    profiles = classify_columns(frame)

    assert source_comparison_available(frame, profiles["quantity"])
    assert not source_comparison_available(frame, profiles["source_file"])

    one_source = frame.loc[frame["source_file"].eq("north.xlsx")].copy()
    one_source_profiles = classify_columns(one_source)
    assert not source_comparison_available(
        one_source,
        one_source_profiles["quantity"],
    )


def test_source_comparison_is_concise_and_does_not_modify_data():
    frame = insight_frame()
    original = frame.copy(deep=True)
    profiles = classify_columns(frame)

    comparison = source_file_comparison(frame, profiles["unit_price"])

    assert list(comparison.columns) == [
        "Source File",
        "Record count",
        "Average",
        "Median",
        "Minimum",
        "Maximum",
    ]
    assert list(comparison["Source File"]) == [
        "north.xlsx",
        "south.xlsx",
    ]
    pd.testing.assert_frame_equal(frame, original)


def test_chart_config_accepts_clear_time_controls():
    frame = insight_frame()
    profiles = classify_columns(frame)

    assert chart_config_is_compatible(
        {
            "type": "Trend over time",
            "date_column": "order_date",
            "value_column": "quantity",
            "group_by": "Month",
            "aggregation": "median",
        },
        profiles,
    )
    assert not chart_config_is_compatible(
        {
            "type": "Trend over time",
            "date_column": "order_date",
            "value_column": "order_id",
            "group_by": "Month",
            "aggregation": "sum",
        },
        profiles,
    )


def test_all_column_statistics_require_explicit_secondary_action():
    source = APP_SOURCE.read_text(encoding="utf-8")

    technical_toggle = source.index('"View technical statistics"')
    all_columns_toggle = source.index('"View statistics for all columns"')
    conditional = source.index("if show_all_statistics:")
    table_builder = source.index(
        "technical_tables = technical_statistics_by_type("
    )

    assert technical_toggle < all_columns_toggle < conditional < table_builder
    assert "if show_technical_details:" in source


def test_insights_ui_uses_friendly_time_and_comparison_labels():
    source = APP_SOURCE.read_text(encoding="utf-8")

    for label in (
        '"Date field"',
        '"Group by"',
        '"Calculation"',
        '"Compare source files"',
        '"View statistics for all columns"',
    ):
        assert label in source
    assert '"Date column"' not in source
    assert '"Combine values by"' not in source


def test_single_time_field_is_rendered_as_read_only():
    source = APP_SOURCE.read_text(encoding="utf-8")

    single_date_condition = source.index("if len(valid_dates) == 1:")
    read_only_field = source.index('class="insight-readonly-field"')
    multiple_date_selector = source.index("date_column = st.selectbox(")

    assert single_date_condition < read_only_field < multiple_date_selector
