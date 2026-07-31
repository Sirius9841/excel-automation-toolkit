"""Tests for explicit workflow routing and state invalidation."""

from io import BytesIO

from docx import Document as DocxDocument
import pandas as pd

from src.report_generator import generate_report
from src.workflow import (
    DEFAULT_REPORT_SECTIONS,
    EXPORT_CONTEXT_SCHEMA_VERSION,
    ExportContext,
    REPORT_SECTION_LABELS,
    SCREEN_ADD,
    SCREEN_CLEAN,
    SCREEN_DOWNLOAD,
    SCREEN_INSIGHTS,
    SCREEN_REVIEW,
    approve_merge_configuration,
    ensure_export_context,
    invalidate_report_output,
    mark_cleaned_data_changed,
    merge_requires_recombine,
    navigate,
    open_insights,
    remove_legacy_state,
    reset_session,
    return_from_insights,
    standard_report_context,
    update_source_signature,
)


def populated_state() -> dict:
    return {
        "current_screen": SCREEN_CLEAN,
        "source_signature": ("files", "v1"),
        "approved_source_signature": ("files", "v1"),
        "approved_merge_mode": "Keep every column",
        "merged_df": pd.DataFrame({"value": [1, 2]}),
        "cleaned_df": pd.DataFrame({"value": [1]}),
        "dup_report": {"removed": 1},
        "mv_actions": ["Column 'value': filled 1 missing values (fill_median)"],
        "cleaning_audit": [{
            "action_type": "fill_blank",
            "audit_event_id": "AUD-1",
        }],
        "chart_config": {"type": "Distribution", "column": "value"},
        "report_bytes": b"report",
        "export_filename": "approved_name",
        "report_sections": list(DEFAULT_REPORT_SECTIONS),
        "report_section_overview": True,
        "report_section_charts": False,
        "missing_action_str:'value'": "Use the middle value",
        "missing_applied_action_str:'value'": "Use the middle value",
        "missing_feedback": "Blank values handled.",
    }


def test_add_files_to_review():
    state = {"current_screen": SCREEN_ADD}
    navigate(state, SCREEN_REVIEW)
    assert state["current_screen"] == SCREEN_REVIEW


def test_review_to_clean_after_combining():
    state = {"current_screen": SCREEN_REVIEW}
    changed = approve_merge_configuration(
        state,
        ("files", "v1"),
        "Keep every column",
    )
    assert changed is True
    assert state["current_screen"] == SCREEN_CLEAN


def test_clean_to_download():
    state = {"current_screen": SCREEN_CLEAN}
    navigate(state, SCREEN_DOWNLOAD)
    assert state["current_screen"] == SCREEN_DOWNLOAD


def test_clean_to_insights_and_back_to_clean():
    state = {"current_screen": SCREEN_CLEAN}
    open_insights(state)
    assert state["current_screen"] == SCREEN_INSIGHTS
    return_from_insights(state)
    assert state["current_screen"] == SCREEN_CLEAN


def test_download_to_insights_and_back_to_download():
    state = {"current_screen": SCREEN_DOWNLOAD}
    open_insights(state)
    assert state["current_screen"] == SCREEN_INSIGHTS
    return_from_insights(state)
    assert state["current_screen"] == SCREEN_DOWNLOAD


def test_data_insights_preferences_survive_normal_navigation():
    state = {
        "current_screen": SCREEN_INSIGHTS,
        "analysis_column": "quantity",
        "insight_numeric_view": "Distribution",
        "insight_measure_over_time": True,
        "insight_trend_date": "order_date",
        "insight_trend_group": "Month",
        "insight_trend_calculation": "Total",
        "insight_compare_sources": True,
        "insight_show_technical": True,
        "insight_show_all_statistics": False,
    }
    expected_preferences = {
        key: value
        for key, value in state.items()
        if key != "current_screen"
    }

    navigate(state, SCREEN_CLEAN)
    navigate(state, SCREEN_INSIGHTS)

    assert {
        key: state[key]
        for key in expected_preferences
    } == expected_preferences


def test_returning_to_review_does_not_destroy_work():
    state = populated_state()
    cleaned_df = state["cleaned_df"]
    navigate(state, SCREEN_REVIEW)
    assert state["cleaned_df"] is cleaned_df
    assert state["chart_config"]["type"] == "Distribution"
    assert state["report_bytes"] == b"report"
    assert state["missing_action_str:'value'"] == "Use the middle value"


def test_normal_navigation_does_not_clear_cleaned_data():
    state = populated_state()
    cleaned_df = state["cleaned_df"]
    navigate(state, SCREEN_DOWNLOAD)
    navigate(state, SCREEN_CLEAN)
    assert state["cleaned_df"] is cleaned_df
    assert state["cleaning_audit"][0]["audit_event_id"] == "AUD-1"
    assert state["missing_action_str:'value'"] == "Use the middle value"


def test_changing_source_files_invalidates_derived_data():
    state = populated_state()
    changed = update_source_signature(state, ("files", "v2"))
    assert changed is True
    assert "merged_df" not in state
    assert "cleaned_df" not in state
    assert "chart_config" not in state
    assert "report_bytes" not in state
    assert "missing_action_str:'value'" not in state
    assert "missing_applied_action_str:'value'" not in state
    assert "missing_feedback" not in state
    assert state["export_filename"] == "approved_name"
    assert state["report_sections"] == list(DEFAULT_REPORT_SECTIONS)
    assert state["report_section_overview"] is True
    assert state["report_section_charts"] is False


def test_unchanged_source_files_preserve_derived_data():
    state = populated_state()
    changed = update_source_signature(state, ("files", "v1"))
    assert changed is False
    assert "cleaned_df" in state
    assert "chart_config" in state
    assert "report_bytes" in state


def test_changing_approved_merge_settings_requires_recombining():
    state = populated_state()
    assert merge_requires_recombine(
        state,
        ("files", "v1"),
        "Keep only matching columns",
    )

    cleaned_df = state["cleaned_df"]
    assert state["cleaned_df"] is cleaned_df

    approve_merge_configuration(
        state,
        ("files", "v1"),
        "Keep only matching columns",
    )
    assert "cleaned_df" not in state
    assert "chart_config" not in state
    assert "report_bytes" not in state


def test_unchanged_approved_merge_preserves_downstream_work():
    state = populated_state()
    cleaned_df = state["cleaned_df"]
    changed = approve_merge_configuration(
        state,
        ("files", "v1"),
        "Keep every column",
    )
    assert changed is False
    assert state["cleaned_df"] is cleaned_df
    assert state["chart_config"]["type"] == "Distribution"
    assert state["report_bytes"] == b"report"


def test_start_over_clears_entire_session():
    state = populated_state()
    reset_session(state)
    assert state == {}


def test_confirmed_legacy_state_is_removed_without_touching_active_state():
    state = {
        "workspace_stage": "Explore",
        "charts_viewed": True,
        "outlier_df": pd.DataFrame({"value": [1]}),
        "cleaned_df": pd.DataFrame({"value": [2]}),
        "explore_view": "Charts",
    }
    remove_legacy_state(state)
    assert "workspace_stage" not in state
    assert "charts_viewed" not in state
    assert "outlier_df" not in state
    assert "cleaned_df" in state
    assert state["explore_view"] == "Charts"


def test_report_generation_does_not_require_insights_state():
    report_df = pd.DataFrame({
        "value": [10.0, 11.0, 12.0, 1000.0],
        "category": ["A", "A", "B", "B"],
    })
    merged_df = report_df.copy()
    state = {
        "current_screen": SCREEN_DOWNLOAD,
        "merged_df": merged_df,
        "cleaned_df": report_df,
    }
    assert "outlier_df" not in state
    assert "chart_config" not in state

    context = standard_report_context(
        report_df,
        merged_df,
        duplicate_report=None,
        missing_actions=None,
    )
    assert isinstance(context, ExportContext)
    assert context.schema_version == EXPORT_CONTEXT_SCHEMA_VERSION
    assert context.integrity_report is not None
    report_bytes = generate_report(
        df=report_df,
        stats_df=context.statistics,
        source_files=["source.xlsx"],
        cleaning_report=context.cleaning_summary,
        outlier_df=context.values_to_review,
        include_sections=DEFAULT_REPORT_SECTIONS,
    )
    document = DocxDocument(BytesIO(report_bytes))
    headings = [
        paragraph.text
        for paragraph in document.paragraphs
        if paragraph.style.name.startswith("Heading")
    ]
    assert "1. Dataset Overview" in headings
    assert "2. Data Quality and Cleaning Summary" in headings
    assert "3. Column Statistics" in headings
    assert "4. Values Worth Reviewing" in headings
    assert "5. Methodology and Limitations" in headings
    assert "Charts" not in " ".join(headings)
    assert "Data Preview" not in " ".join(headings)


def test_export_context_always_contains_an_empty_integrity_report():
    frame = pd.DataFrame({"category": ["A", "B"]})
    context = standard_report_context(
        frame,
        frame.copy(),
        duplicate_report=None,
        missing_actions=None,
        source_files=["source.xlsx"],
    )

    assert context.integrity_report.issues == ()
    assert context.integrity_report.relationships == ()
    assert context.integrity_report.severe_count == 0


def test_export_context_contains_integrity_warnings_when_present():
    merged = pd.DataFrame({
        "quantity": [1, 2, 3, 4, 5, 6, 7, 8],
        "unit_price": [10.0] * 8,
        "total": [10.0, 20.0, 30.0, 40.0, 50.0, 60.0, 70.0, 80.0],
    })
    cleaned = merged.copy()
    cleaned.loc[0, "total"] = 1000.0

    context = standard_report_context(
        cleaned,
        merged,
        duplicate_report=None,
        missing_actions=None,
    )

    assert context.integrity_report.severe_count == 1
    assert len(context.values_to_review) >= 1


def test_stale_cached_export_context_is_rebuilt_without_missing_key_crash():
    frame = pd.DataFrame({"value": [1.0, 2.0]})
    state = {
        "export_context": {
            "stats_df": pd.DataFrame(),
            "cleaning_report": {},
            "outlier_df": pd.DataFrame(),
        }
    }

    context = ensure_export_context(
        state,
        frame,
        frame.copy(),
        duplicate_report=None,
        missing_actions=None,
        source_files=["source.xlsx"],
    )

    assert isinstance(context, ExportContext)
    assert context.integrity_report is not None
    assert state["export_context"] is context


def test_valid_export_context_survives_navigation_and_is_reused():
    frame = pd.DataFrame({"value": [1.0, 2.0]})
    state: dict = {"current_screen": SCREEN_DOWNLOAD}
    context = ensure_export_context(
        state,
        frame,
        frame.copy(),
        duplicate_report=None,
        missing_actions=None,
        source_files=["source.xlsx"],
    )

    navigate(state, SCREEN_INSIGHTS)
    navigate(state, SCREEN_DOWNLOAD)
    reused = ensure_export_context(
        state,
        frame,
        frame.copy(),
        duplicate_report=None,
        missing_actions=None,
        source_files=["source.xlsx"],
    )

    assert reused is context


def test_changed_cleaned_data_rebuilds_export_context():
    original = pd.DataFrame({"value": [1.0, 2.0]})
    changed = pd.DataFrame({"value": [1.0, 3.0]})
    state: dict = {}
    original_context = ensure_export_context(
        state,
        original,
        original.copy(),
        duplicate_report=None,
        missing_actions=None,
    )
    rebuilt = ensure_export_context(
        state,
        changed,
        original.copy(),
        duplicate_report=None,
        missing_actions=None,
    )

    assert rebuilt is not original_context
    assert rebuilt.data_signature != original_context.data_signature


def test_standard_report_sections_exclude_optional_charts_and_preview():
    assert DEFAULT_REPORT_SECTIONS == (
        "overview",
        "quality",
        "statistics",
        "outliers",
        "methodology",
    )


def test_report_checklist_uses_clear_standard_and_optional_labels():
    assert [
        REPORT_SECTION_LABELS[section]
        for section in DEFAULT_REPORT_SECTIONS
    ] == [
        "Dataset overview",
        "Cleaning summary",
        "Key statistics",
        "Values worth reviewing",
        "Notes and limitations",
    ]
    assert REPORT_SECTION_LABELS["charts"] == "Charts"
    assert REPORT_SECTION_LABELS["preview"] == "Data preview"


def test_invalidating_report_preserves_data_and_report_preferences():
    state = populated_state()
    cleaned_df = state["cleaned_df"]
    state["report_input_signature"] = ("signature",)
    state["report_generating"] = True
    state["report_error"] = "old error"

    invalidate_report_output(state)

    assert state["cleaned_df"] is cleaned_df
    assert state["report_sections"] == list(DEFAULT_REPORT_SECTIONS)
    assert "report_bytes" not in state
    assert "report_input_signature" not in state
    assert "report_generating" not in state
    assert "report_error" not in state


def test_cleaned_data_change_advances_revision_and_invalidates_report():
    state = populated_state()
    state["data_revision"] = 4
    state["report_input_signature"] = ("signature",)
    state["export_context"] = {"legacy": True}

    mark_cleaned_data_changed(state)

    assert state["data_revision"] == 5
    assert "export_context" not in state
    assert "report_bytes" not in state
    assert "report_input_signature" not in state
