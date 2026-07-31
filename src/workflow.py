"""Explicit screen routing and state lifecycle for the Streamlit workflow."""

from collections.abc import Hashable, MutableMapping
from dataclasses import dataclass
from hashlib import sha256
from typing import Any

import pandas as pd

from src.analyzer import (
    detect_business_column_type,
    generate_summary_statistics,
    summarize_outliers,
)
from src.data_processor import generate_cleaning_report
from src.integrity import (
    IntegrityReport,
    combine_review_findings,
    detect_validated_relationships,
    validate_integrity,
)

SCREEN_ADD = "add"
SCREEN_REVIEW = "review"
SCREEN_CLEAN = "clean"
SCREEN_DOWNLOAD = "download"
SCREEN_INSIGHTS = "insights"
EXPORT_CONTEXT_SCHEMA_VERSION = 2

VALID_SCREENS = {
    SCREEN_ADD,
    SCREEN_REVIEW,
    SCREEN_CLEAN,
    SCREEN_DOWNLOAD,
    SCREEN_INSIGHTS,
}

REPORT_SECTION_LABELS = {
    "overview": "Dataset overview",
    "quality": "Cleaning summary",
    "statistics": "Key statistics",
    "outliers": "Values worth reviewing",
    "methodology": "Notes and limitations",
    "charts": "Charts",
    "preview": "Data preview",
}

DEFAULT_REPORT_SECTIONS = (
    "overview",
    "quality",
    "statistics",
    "outliers",
    "methodology",
)

LEGACY_STATE_KEYS = {
    "workspace_stage",
    "stats_generated",
    "charts_viewed",
    "exported",
    "report_generated",
    "outlier_df",
}

DOWNSTREAM_STATE_KEYS = {
    "cleaned_df",
    "dup_report",
    "mv_actions",
    "cleaning_audit",
    "export_context",
    "chart_config",
    "report_bytes",
    "report_input_signature",
    "report_generating",
    "report_error",
    "integrity_acknowledged",
    "data_revision",
    "missing_feedback",
    "missing_selections_dirty",
    "missing_pending_override",
    "edit_approved_blanks",
    "missing_integrity_failures",
    "clean_view",
    "duplicate_match_mode",
    "duplicate_identity_columns_v2",
    "analysis_column",
    "insight_numeric_view",
    "insight_measure_over_time",
    "insight_trend_date",
    "insight_trend_aggregation",
    "insight_trend_measure",
    "insight_trend_group",
    "insight_trend_calculation",
    "insight_compare_sources",
    "insight_show_technical",
    "insight_show_all_statistics",
    "view_all_missing_changes",
    "chart_type",
    "chart_numeric_column",
    "chart_category_column",
    "chart_category_count",
    "chart_date_column",
    "chart_value_column",
    "chart_aggregation",
}


@dataclass(frozen=True)
class ExportContext:
    """Complete, versioned contract consumed by the Download screen."""

    schema_version: int
    cleaned_df: pd.DataFrame
    statistics: pd.DataFrame
    cleaning_summary: dict[str, Any]
    cleaning_audit: tuple[dict[str, Any], ...]
    integrity_report: IntegrityReport
    values_to_review: pd.DataFrame
    statistical_values_to_review: pd.DataFrame
    source_schemas: dict[str, frozenset[str]]
    source_files: tuple[str, ...]
    workflow_signature: tuple[Any, ...]
    data_signature: tuple[Any, ...]
    export_signature: tuple[Any, ...]
    report_signature: tuple[Any, ...]

    @property
    def stats_df(self) -> pd.DataFrame:
        """Compatibility alias for report-generation call sites."""
        return self.statistics

    @property
    def cleaning_report(self) -> dict[str, Any]:
        """Compatibility alias for export/report call sites."""
        return self.cleaning_summary

    @property
    def outlier_df(self) -> pd.DataFrame:
        """Compatibility alias for values requiring review."""
        return self.values_to_review

    @property
    def statistical_outlier_df(self) -> pd.DataFrame:
        """Compatibility alias for statistical-only review findings."""
        return self.statistical_values_to_review

    def __getitem__(self, key: str) -> Any:
        """Support legacy read-only access while callers migrate to attributes."""
        aliases = {
            "stats_df": self.statistics,
            "cleaning_report": self.cleaning_summary,
            "outlier_df": self.values_to_review,
            "statistical_outlier_df": self.statistical_values_to_review,
        }
        if key in aliases:
            return aliases[key]
        return getattr(self, key)

SOURCE_DERIVED_STATE_KEYS = DOWNSTREAM_STATE_KEYS | {
    "merged_df",
    "approved_source_signature",
    "approved_merge_mode",
    "source_schemas",
}

DATA_WIDGET_PREFIXES = (
    "missing_editor_",
    "missing_action_",
    "missing_applied_action_",
    "missing_applied_custom_",
    "custom_missing_",
)


def navigate(state: MutableMapping[str, Any], target: str) -> None:
    """Move to a screen without invalidating approved work."""
    if target not in VALID_SCREENS:
        raise ValueError(f"Unknown workflow screen: {target}")
    state["current_screen"] = target


def open_insights(state: MutableMapping[str, Any]) -> None:
    """Open optional insights and remember where the user should return."""
    origin = state.get("current_screen", SCREEN_CLEAN)
    if origin not in {SCREEN_CLEAN, SCREEN_DOWNLOAD}:
        origin = SCREEN_CLEAN
    state["insights_return_screen"] = origin
    state["current_screen"] = SCREEN_INSIGHTS


def return_from_insights(state: MutableMapping[str, Any]) -> None:
    """Return from insights to the screen that opened it."""
    target = state.get("insights_return_screen", SCREEN_CLEAN)
    if target not in {SCREEN_CLEAN, SCREEN_DOWNLOAD}:
        target = SCREEN_CLEAN
    state["current_screen"] = target


def _remove_state(
    state: MutableMapping[str, Any],
    keys: set[str],
) -> None:
    for key in keys:
        state.pop(key, None)
    for key in list(state.keys()):
        if isinstance(key, str) and key.startswith(DATA_WIDGET_PREFIXES):
            state.pop(key, None)


def update_source_signature(
    state: MutableMapping[str, Any],
    signature: Hashable,
) -> bool:
    """Store a source signature and invalidate derived work only if it changed."""
    if state.get("source_signature") == signature:
        return False
    _remove_state(state, SOURCE_DERIVED_STATE_KEYS)
    state["source_signature"] = signature
    return True


def merge_requires_recombine(
    state: MutableMapping[str, Any],
    source_signature: Hashable,
    merge_mode: str,
) -> bool:
    """Return whether the selected configuration differs from the approved merge."""
    if state.get("merged_df") is None:
        return False
    return (
        state.get("approved_source_signature") != source_signature
        or state.get("approved_merge_mode") != merge_mode
    )


def approve_merge_configuration(
    state: MutableMapping[str, Any],
    source_signature: Hashable,
    merge_mode: str,
) -> bool:
    """Approve a completed merge and clear stale downstream work when needed."""
    configuration_changed = (
        state.get("approved_source_signature") != source_signature
        or state.get("approved_merge_mode") != merge_mode
    )
    if configuration_changed:
        _remove_state(state, DOWNSTREAM_STATE_KEYS)
    state["approved_source_signature"] = source_signature
    state["approved_merge_mode"] = merge_mode
    state["current_screen"] = SCREEN_CLEAN
    return configuration_changed


def invalidate_report_output(state: MutableMapping[str, Any]) -> None:
    """Remove a generated report when one of its inputs changes."""
    state.pop("report_bytes", None)
    state.pop("report_input_signature", None)
    state.pop("report_error", None)
    state.pop("report_generating", None)


def mark_cleaned_data_changed(state: MutableMapping[str, Any]) -> None:
    """Advance the cleaned-data revision and invalidate its generated report."""
    state["data_revision"] = int(state.get("data_revision", 0)) + 1
    state.pop("integrity_acknowledged", None)
    state.pop("export_context", None)
    invalidate_report_output(state)


def reset_session(state: MutableMapping[str, Any]) -> None:
    """Clear the entire temporary workflow session."""
    state.clear()


def remove_legacy_state(state: MutableMapping[str, Any]) -> None:
    """Remove session keys confirmed unused by the explicit screen router."""
    for key in LEGACY_STATE_KEYS:
        state.pop(key, None)


def dataframe_signature(df: pd.DataFrame) -> tuple[Any, ...]:
    """Return a stable signature for the current DataFrame contents and shape."""
    columns = tuple(str(column) for column in df.columns)
    dtypes = tuple(str(dtype) for dtype in df.dtypes)
    try:
        hashed = pd.util.hash_pandas_object(df, index=True).values.tobytes()
        digest = sha256(hashed).hexdigest()
    except (TypeError, ValueError):
        digest = sha256(
            df.to_csv(index=True, na_rep="<NA>").encode("utf-8")
        ).hexdigest()
    return (len(df), columns, dtypes, digest)


def _source_schema_signature(
    source_schemas: dict[str, set[str] | frozenset[str]] | None,
) -> tuple[tuple[str, tuple[str, ...]], ...]:
    return tuple(
        sorted(
            (
                str(source),
                tuple(sorted(str(column) for column in columns)),
            )
            for source, columns in (source_schemas or {}).items()
        )
    )


def _normalise_source_schemas(
    source_schemas: dict[str, set[str] | frozenset[str]] | None,
) -> dict[str, frozenset[str]]:
    return {
        str(source): frozenset(str(column) for column in columns)
        for source, columns in (source_schemas or {}).items()
    }


def _audit_signature(
    cleaning_audit: list[dict[str, Any]] | None,
) -> tuple[tuple[str, str, str], ...]:
    return tuple(
        (
            str(record.get("audit_event_id", "")),
            str(record.get("action_type", "")),
            str(record.get("resulting_value", "")),
        )
        for record in (cleaning_audit or [])
    )


def _context_report_signature(
    report_df: pd.DataFrame,
    merged_df: pd.DataFrame,
    duplicate_report: dict | None,
    missing_actions: list[str] | None,
    source_schemas: dict[str, set[str] | frozenset[str]] | None,
    source_files: tuple[str, ...],
    cleaning_audit: list[dict[str, Any]] | None,
    workflow_signature: tuple[Any, ...],
) -> tuple[Any, ...]:
    export_signature = (
        EXPORT_CONTEXT_SCHEMA_VERSION,
        dataframe_signature(report_df),
        _source_schema_signature(source_schemas),
        source_files,
        workflow_signature,
    )
    return (
        export_signature,
        dataframe_signature(merged_df),
        _audit_signature(cleaning_audit),
        tuple(missing_actions or ()),
        repr(duplicate_report),
    )


def standard_report_context(
    report_df: pd.DataFrame,
    merged_df: pd.DataFrame,
    duplicate_report: dict | None,
    missing_actions: list[str] | None,
    *,
    source_schemas: dict[str, set[str] | frozenset[str]] | None = None,
    source_files: list[str] | tuple[str, ...] | None = None,
    cleaning_audit: list[dict[str, Any]] | None = None,
    workflow_signature: tuple[Any, ...] = (),
) -> ExportContext:
    """Build complete standard-report inputs without relying on Insights state."""
    numeric_columns = [
        column
        for column in report_df.columns
        if detect_business_column_type(report_df[column]) == "numeric"
    ]
    outlier_df = (
        summarize_outliers(report_df, numeric_columns)
        if numeric_columns
        else pd.DataFrame()
    )
    relationships = detect_validated_relationships(merged_df)
    integrity_report = validate_integrity(
        report_df,
        relationships,
    )
    review_df = combine_review_findings(outlier_df, integrity_report)
    normalised_schemas = _normalise_source_schemas(source_schemas)
    context_source_files = tuple(source_files or normalised_schemas)
    data_signature = dataframe_signature(report_df)
    report_signature = _context_report_signature(
        report_df,
        merged_df,
        duplicate_report,
        missing_actions,
        source_schemas,
        context_source_files,
        cleaning_audit,
        workflow_signature,
    )
    export_signature = report_signature[0]
    return ExportContext(
        schema_version=EXPORT_CONTEXT_SCHEMA_VERSION,
        cleaned_df=report_df,
        statistics=generate_summary_statistics(
            report_df,
            source_schemas=source_schemas,
            cleaning_audit=cleaning_audit or (),
        ),
        values_to_review=review_df,
        statistical_values_to_review=outlier_df,
        integrity_report=integrity_report,
        cleaning_summary=generate_cleaning_report(
            merged_df,
            report_df,
            duplicate_report,
            missing_actions,
            source_schemas=source_schemas,
            cleaning_audit=cleaning_audit,
            integrity_report=integrity_report,
        ),
        cleaning_audit=tuple(dict(record) for record in (cleaning_audit or [])),
        source_schemas=normalised_schemas,
        source_files=context_source_files,
        workflow_signature=workflow_signature,
        data_signature=data_signature,
        export_signature=export_signature,
        report_signature=report_signature,
    )


def ensure_export_context(
    state: MutableMapping[str, Any],
    report_df: pd.DataFrame,
    merged_df: pd.DataFrame,
    duplicate_report: dict | None,
    missing_actions: list[str] | None,
    *,
    source_schemas: dict[str, set[str] | frozenset[str]] | None = None,
    source_files: list[str] | tuple[str, ...] | None = None,
    cleaning_audit: list[dict[str, Any]] | None = None,
    workflow_signature: tuple[Any, ...] = (),
) -> ExportContext:
    """Return a valid cached context, rebuilding stale or legacy shapes.

    Contexts created before the typed schema (including dictionaries missing
    ``integrity_report``) are deliberately invalidated and rebuilt.
    """
    expected_data_signature = dataframe_signature(report_df)
    expected_schema_signature = _source_schema_signature(source_schemas)
    expected_source_files = tuple(source_files or ())
    if not expected_source_files and source_schemas:
        expected_source_files = tuple(source_schemas)
    expected_report_signature = _context_report_signature(
        report_df,
        merged_df,
        duplicate_report,
        missing_actions,
        source_schemas,
        expected_source_files,
        cleaning_audit,
        workflow_signature,
    )
    cached = state.get("export_context")
    cached_is_current = (
        isinstance(cached, ExportContext)
        and cached.schema_version == EXPORT_CONTEXT_SCHEMA_VERSION
        and cached.data_signature == expected_data_signature
        and cached.report_signature == expected_report_signature
        and _source_schema_signature(cached.source_schemas)
        == expected_schema_signature
        and cached.source_files == expected_source_files
    )
    if cached_is_current:
        return cached

    context = standard_report_context(
        report_df,
        merged_df,
        duplicate_report,
        missing_actions,
        source_schemas=source_schemas,
        source_files=source_files,
        cleaning_audit=cleaning_audit,
        workflow_signature=workflow_signature,
    )
    state["export_context"] = context
    return context
