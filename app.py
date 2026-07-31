"""Business Excel Automation Toolkit — Streamlit entry point."""

from datetime import date
from html import escape
import re

import streamlit as st
import pandas as pd

from config.settings import SCHEMA_THRESHOLD, SAMPLE_DATA_DIR, MAX_FILE_SIZE_MB
from src.logger_setup import setup_logger
from src.file_handler import (
    read_uploaded_file,
    describe_dataframe,
    FileReadError,
    FileSizeError,
)
from src.data_processor import (
    merge_datasets,
    compute_schema_compatibility,
    remove_duplicates,
    detect_missing_values,
    recommend_missing_strategy,
)
from src.data_quality import (
    apply_missing_value_strategies,
    audit_summary,
    build_review_summary,
    build_source_schemas,
    cleaning_change_preview,
    grouped_audit_change_table,
    missing_decision_summary,
    missing_status_by_column,
    recommendation_explanation,
)
from src.analyzer import (
    friendly_column_name,
    summarize_outliers,
)
from src.insights import (
    SEMANTIC_CATEGORICAL,
    SEMANTIC_DATE,
    SEMANTIC_IDENTIFIER,
    SEMANTIC_NUMERIC,
    blank_status_summary,
    chart_config_is_compatible,
    classify_columns,
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
from src.visualizer import (
    plot_histogram,
    plot_boxplot,
    plot_bar_chart,
    plot_date_counts,
    plot_line_chart,
    figure_to_bytes,
)
from src.exporter import export_to_excel, export_to_csv, ExportError
from src.report_generator import generate_report, ReportError
from src.utils import is_allowed_file
from src.ui_helpers import (
    DEFAULT_DUPLICATE_COLUMNS,
    duplicate_subset,
)
from src.workflow import (
    DEFAULT_REPORT_SECTIONS,
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
    update_source_signature,
)

logger = setup_logger(__name__)

# ── Page configuration ───────────────────────────────────
st.set_page_config(
    page_title="Excel Automation Toolkit",
    page_icon=":material/table_chart:",
    layout="wide",
)

# ── Custom CSS ───────────────────────────────────────────
st.markdown("""
<style>
    :root {
        --product-page: #e4ecf6;
        --product-surface: #ffffff;
        --product-surface-secondary: #f4f7fc;
        --product-surface-blue: #eef4ff;
        --product-surface-indigo: #f1f0ff;
        --product-text: #0f172a;
        --product-body: #1e293b;
        --product-helper: #475569;
        --product-metadata: #64748b;
        --product-primary: #2563eb;
        --product-indigo: #4f46e5;
        --product-success: #15803d;
        --product-warning: #b45309;
        --product-border: #c7d2e0;
        --product-border-soft: #d8e1ec;
        --product-radius: 11px;
        --product-shadow: 0 13px 32px rgba(15, 23, 42, 0.09);
        --space-1: 6px;
        --space-2: 10px;
        --space-3: 13px;
        --space-4: 20px;
        --space-5: 27px;
        --density-body: 0.75rem;
        --density-helper: 0.7rem;
        --density-meta: 0.65rem;
        --density-control: 2rem;
    }
    html,
    body,
    #root,
    .stApp,
    div[data-testid="stAppViewContainer"] {
        background-color: var(--product-page) !important;
        background-image:
            radial-gradient(
                circle at 8% 8%,
                rgba(37, 99, 235, 0.28) 0,
                rgba(37, 99, 235, 0) 30%
            ),
            radial-gradient(
                circle at 92% 88%,
                rgba(79, 70, 229, 0.09) 0,
                rgba(79, 70, 229, 0) 32%
            ),
            linear-gradient(
                145deg,
                #f1f5fb 0%,
                var(--product-page) 55%,
                #eef1fa 100%
            ) !important;
        background-attachment: fixed !important;
        background-repeat: no-repeat !important;
        background-size: cover !important;
        width: 100%;
        min-height: 100vh;
        margin: 0;
    }
    section[data-testid="stMain"] {
        width: 100%;
        min-height: 100vh;
        background: transparent !important;
    }
    header[data-testid="stHeader"] {
        height: 38px;
        min-height: 38px;
        background: rgba(228, 236, 246, 0.78) !important;
        backdrop-filter: blur(12px);
    }
    div[data-testid="stMainBlockContainer"]:has(.landing-intro) {
        box-sizing: border-box;
        width: min(92vw, 1248px) !important;
        max-width: none !important;
        min-height: calc(100vh - 38px);
        display: flex;
        flex-direction: column;
        justify-content: center;
        margin-right: auto !important;
        margin-left: auto !important;
        padding:
            clamp(1.2rem, 2vh, 1.8rem)
            0
            clamp(2rem, 3.2vh, 2.8rem) !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.landing-intro) {
        gap: clamp(2rem, 2.25vw, 2.8rem) !important;
    }
    div[data-testid="stMainBlockContainer"]:has(.internal-page-marker) {
        box-sizing: border-box;
        width: min(88vw, 832px) !important;
        max-width: none !important;
        margin-right: auto !important;
        margin-left: auto !important;
        padding-top: 1.4rem;
        padding-right: 0;
        padding-bottom: 2rem;
        padding-left: 0;
    }
    .landing-intro {
        padding: 0.6rem 0;
    }
    .landing-intro h1 {
        margin: 0 0 0.44rem;
        color: var(--product-text);
        font-size: clamp(2.3rem, 2.25vw, 2.7rem);
        font-weight: 700;
        letter-spacing: -0.03em;
        line-height: 1.12;
    }
    .workspace-title {
        margin: 0 0 0.36rem;
        color: var(--product-text);
        font-size: clamp(1.6rem, 1.55vw, 1.8rem);
        font-weight: 700;
        letter-spacing: -0.025em;
        line-height: 1.15;
    }
    .landing-intro p, .workspace-subtitle {
        margin: 0;
        color: var(--product-body);
        font-size: 0.8rem;
        line-height: 1.55;
    }
    .workspace-subtitle {
        font-size: var(--density-body);
    }
    .landing-lead {
        max-width: 31rem;
        margin: 0.8rem 0 0.4rem;
        color: var(--product-text);
        font-size: 0.9rem;
        font-weight: 600;
        line-height: 1.45;
    }
    .landing-support {
        max-width: 32rem;
        color: var(--product-helper);
        font-size: 0.8rem;
        line-height: 1.55;
    }
    .brand-title-row {
        display: flex;
        gap: 0.9rem;
        align-items: center;
    }
    .brand-mark {
        display: grid;
        width: clamp(2.4rem, 2.2vw, 2.6rem);
        height: clamp(2.4rem, 2.2vw, 2.6rem);
        flex: 0 0 clamp(2.4rem, 2.2vw, 2.6rem);
        place-items: center;
        border-radius: 14px;
        background: linear-gradient(145deg, #2563eb 0%, #4f46e5 100%);
        color: #ffffff;
        box-shadow: 0 8px 19px rgba(37, 99, 235, 0.24);
    }
    .brand-mark svg {
        width: 1.35rem;
        height: 1.35rem;
        fill: none;
        stroke: currentColor;
        stroke-linecap: round;
        stroke-linejoin: round;
        stroke-width: 1.8;
    }
    .brand-mark.compact {
        width: 1.72rem;
        height: 1.72rem;
        flex-basis: 1.72rem;
        border-radius: 7px;
        box-shadow: 0 5px 13px rgba(37, 99, 235, 0.18);
    }
    .brand-mark.compact svg {
        width: 0.96rem;
        height: 0.96rem;
    }
    .internal-brand {
        display: flex;
        gap: 0.56rem;
        align-items: center;
        color: var(--product-text);
        font-size: var(--density-helper);
        font-weight: 650;
    }
    .st-key-internal_app_header {
        margin-bottom: 1.2rem;
        padding: 0 0 0.6rem;
        border-bottom: 1px solid rgba(148, 163, 184, 0.38);
    }
    .internal-page-marker {
        display: none;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker) {
        padding: 16px 19px 22px;
        border: 1px solid rgba(199, 210, 224, 0.92);
        border-radius: 14px;
        background: rgba(255, 255, 255, 0.78);
        box-shadow: 0 16px 48px rgba(15, 23, 42, 0.08);
        backdrop-filter: blur(6px);
    }
    .landing-eyebrow {
        display: inline-block;
        margin-bottom: 0.67rem;
        color: var(--product-primary);
        font-size: 0.66rem;
        font-weight: 700;
        letter-spacing: 0.07em;
        text-transform: uppercase;
    }
    .landing-benefits {
        display: grid;
        gap: 0.6rem;
        margin-top: 1.2rem;
    }
    .landing-benefit {
        display: grid;
        grid-template-columns: 1.32rem 1fr;
        gap: 0.67rem;
        align-items: center;
        padding: 0.14rem 0;
        color: var(--product-text);
        font-size: var(--density-body);
        line-height: 1.4;
    }
    .landing-benefit strong {
        font-size: var(--density-body);
        font-weight: 650;
    }
    .landing-benefit-icon {
        display: grid;
        width: 1.2rem;
        height: 1.2rem;
        place-items: center;
        border-radius: 50%;
        background: #eaf6ee;
        color: var(--product-success);
    }
    .landing-benefit-icon svg {
        width: 0.72rem;
        height: 0.72rem;
        stroke: currentColor;
        stroke-width: 2.5;
        fill: none;
    }
    .landing-benefit span {
        display: block;
        color: var(--product-helper);
        font-size: var(--density-helper);
        line-height: 1.5;
    }
    .workflow-strip {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 1.4rem;
        margin-top: 13px;
        padding-top: 22px;
        border-top: 1px solid var(--product-border);
    }
    .workflow-item {
        position: relative;
        display: grid;
        grid-template-columns: 1.52rem 1fr;
        gap: 0.67rem;
        align-items: start;
        color: var(--product-text);
        font-size: var(--density-helper);
        line-height: 1.4;
    }
    .workflow-marker {
        position: relative;
        z-index: 1;
        display: grid;
        width: 1.44rem;
        height: 1.44rem;
        place-items: center;
        border-radius: 50%;
        background: #e9f1ff;
        color: var(--product-primary);
        font-size: 0.66rem;
        font-weight: 700;
    }
    .workflow-item > div:last-child {
        position: relative;
        z-index: 1;
        padding-right: var(--space-2);
        background: transparent;
    }
    .workflow-item strong {
        display: block;
        margin-bottom: 0.16rem;
        font-size: var(--density-body);
    }
    .workflow-item span {
        color: var(--product-helper);
        font-size: var(--density-helper);
        line-height: 1.5;
    }
    .panel-title {
        margin: 0 0 0.3rem;
        color: var(--product-text);
        font-size: 0.92rem;
        font-weight: 650;
    }
    .supporting-copy, .helper-copy {
        color: var(--product-helper);
        font-size: var(--density-helper);
        line-height: 1.55;
        opacity: 1;
    }
    div[data-testid="stMarkdownContainer"] p {
        color: var(--product-body);
        line-height: 1.55;
        opacity: 1;
    }
    div[data-testid="stWidgetLabel"] p {
        color: var(--product-body) !important;
        font-size: var(--density-helper) !important;
        font-weight: 500;
        line-height: 1.45 !important;
        opacity: 1 !important;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-testid="stVerticalBlock"] {
        gap: 0.8rem;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-testid="stHorizontalBlock"] {
        gap: 0.8rem;
    }
    .stButton > button,
    .stDownloadButton > button {
        min-height: var(--density-control);
        padding: 0.3rem 0.75rem;
    }
    .stButton > button p,
    .stDownloadButton > button p,
    div[data-testid="stSegmentedControl"] button p {
        font-size: var(--density-body) !important;
        line-height: 1.35 !important;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-testid="stAlert"] {
        padding: 0.6rem 0.8rem;
        border-radius: 8px;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-testid="stAlert"] p {
        font-size: var(--density-body) !important;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-testid="stRadio"] label {
        min-height: 1.5rem;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-testid="stSegmentedControl"] button {
        min-height: 1.6rem;
        padding: 0.2rem 0.6rem;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-baseweb="select"] > div,
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    input,
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    textarea {
        min-height: var(--density-control);
        font-size: var(--density-body);
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border: 1px solid var(--product-border-soft) !important;
        border-radius: 9px !important;
        background: var(--product-surface-secondary) !important;
        box-shadow: none;
    }
    .st-key-review_workspace
    > div[data-testid="stVerticalBlockBorderWrapper"] {
        position: relative;
        overflow: hidden;
        border: 1px solid var(--product-border) !important;
        border-radius: var(--product-radius) !important;
        background: rgba(255, 255, 255, 0.97) !important;
        box-shadow: var(--product-shadow);
        padding: 0.7rem !important;
    }
    .st-key-review_workspace
    > div[data-testid="stVerticalBlockBorderWrapper"]::before {
        position: absolute;
        z-index: 2;
        top: 0;
        right: 0;
        left: 0;
        height: 3px;
        background: linear-gradient(90deg, #2563eb, #4f46e5);
        content: "";
    }
    .metric-grid {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.48rem;
        margin: 0.52rem 0 0.72rem;
    }
    .metric-card {
        padding: 0.5rem 0.6rem;
        border: 1px solid #d8e3f0;
        border-radius: 8px;
        background: var(--product-surface-secondary);
        box-shadow: none;
    }
    .metric-value {
        color: var(--product-text);
        font-size: 0.96rem;
        font-weight: 700;
        line-height: 1.15;
    }
    .metric-label {
        margin-top: 0.2rem;
        color: var(--product-metadata);
        font-size: 0.61rem;
        line-height: 1.45;
    }
    .status-ready {
        color: var(--product-success);
        font-size: 0.6rem;
        font-weight: 600;
    }
    .dataset-summary {
        margin: 0.28rem 0 0.72rem;
        color: var(--product-metadata);
        font-size: var(--density-helper);
        line-height: 1.5;
        opacity: 1;
    }
    div[data-testid="stCaptionContainer"],
    div[data-testid="stCaptionContainer"] p {
        color: var(--product-helper) !important;
        font-size: var(--density-helper) !important;
        font-weight: 400 !important;
        line-height: 1.5 !important;
        opacity: 1 !important;
    }
    .suggestion-text {
        margin-top: -0.2rem;
        color: #334155;
        font-size: var(--density-helper);
        line-height: 1.5;
        opacity: 1;
    }
    .suggestion-text strong {
        color: var(--product-body);
        font-weight: 650;
    }
    .recommendation-badge,
    .difference-badge {
        display: inline-flex;
        align-items: center;
        margin-right: 0.32rem;
        padding: 0.14rem 0.36rem;
        border-radius: 999px;
        font-size: 0.58rem;
        font-weight: 700;
        letter-spacing: 0.01em;
    }
    .recommendation-badge {
        background: #dbeafe;
        color: #1d4ed8;
    }
    .recommendation-explanation {
        margin-top: 0.28rem;
        color: var(--product-helper);
        font-size: var(--density-helper);
        line-height: 1.5;
    }
    .difference-badge {
        background: #fef3c7;
        color: #92400e;
    }
    .effect-preview {
        margin-top: 0.28rem;
        color: var(--product-helper);
        font-size: 0.67rem;
        line-height: 1.5;
    }
    .status-surface {
        display: flex;
        gap: var(--space-2);
        align-items: flex-start;
        margin: var(--space-3) 0;
        padding: var(--space-3);
        border: 1px solid;
        border-radius: 8px;
        color: var(--product-body);
        font-size: var(--density-helper);
        line-height: 1.5;
    }
    .status-surface.info {
        border-color: #bfdbfe;
        background: #eff6ff;
    }
    .status-surface.success {
        border-color: #bbf7d0;
        background: #f0fdf4;
    }
    .status-surface.warning {
        border-color: #fde68a;
        background: #fffbeb;
    }
    .status-icon {
        display: grid;
        width: 1.24rem;
        height: 1.24rem;
        flex: 0 0 1.24rem;
        place-items: center;
        border-radius: 50%;
        background: rgba(37, 99, 235, 0.12);
        color: var(--product-primary);
        font-weight: 750;
    }
    .status-surface.success .status-icon {
        background: rgba(21, 128, 61, 0.12);
        color: var(--product-success);
    }
    .completion-panel {
        margin: var(--space-3) 0 var(--space-4);
        padding: 1rem;
        border: 1px solid #b9d6c2;
        border-radius: 10px;
        background: linear-gradient(145deg, #f0fdf4 0%, #eef7ff 100%);
    }
    .completion-panel h3 {
        margin: 0 0 0.32rem;
        color: var(--product-text);
        font-size: 0.96rem;
    }
    .completion-panel p {
        margin: 0;
        color: var(--product-body);
    }
    .activity-list {
        display: grid;
        gap: 0.44rem;
        margin-top: var(--space-3);
        color: var(--product-helper);
        font-size: var(--density-helper);
    }
    .activity-item {
        display: flex;
        gap: 0.44rem;
        align-items: center;
    }
    .activity-symbol {
        color: var(--product-success);
        font-weight: 750;
    }
    .activity-symbol.neutral {
        color: var(--product-metadata);
    }
    div[data-testid="stExpander"] {
        margin-bottom: 0.52rem;
        border-color: var(--product-border) !important;
        border-radius: var(--product-radius) !important;
        background: var(--product-surface);
    }
    .st-key-landing_upload_card
    section[data-testid="stFileUploaderDropzone"] {
        padding: 1.2rem;
        border: 2px dashed #5b8de3 !important;
        border-radius: var(--product-radius) !important;
        background:
            linear-gradient(
                180deg,
                #eaf2ff 0%,
                #dceaff 100%
            ) !important;
        box-shadow:
            inset 0 0 0 1px rgba(255, 255, 255, 0.86),
            0 6px 18px rgba(37, 99, 235, 0.10);
    }
    .st-key-landing_upload_card
    section[data-testid="stFileUploaderDropzone"]:has(> span) {
        min-height: 10.8rem;
        flex-direction: column;
        gap: 7px;
        align-items: center;
        justify-content: center;
        text-align: center;
    }
    .st-key-landing_upload_card
    section[data-testid="stFileUploaderDropzone"]:has(> span)::before {
        order: 1;
        width: 2.2rem;
        height: 2.2rem;
        flex: 0 0 2.2rem;
        border-radius: 50%;
        background:
            #dbeafe
            url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%232563eb' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M12 16V4m0 0L7 9m5-5 5 5M5 20h14'/%3E%3C/svg%3E")
            center / 1.24rem 1.24rem no-repeat;
        content: "";
    }
    .st-key-landing_upload_card
    section[data-testid="stFileUploaderDropzone"]:has(> span)::after {
        order: 2;
        color: var(--product-text);
        font-size: 0.8rem;
        font-weight: 650;
        line-height: 1.35;
        content: "Drag and drop files here";
    }
    .st-key-landing_upload_card
    section[data-testid="stFileUploaderDropzone"]:has(> span)
    > span {
        order: 3;
        display: flex;
        flex-direction: column;
        gap: 9px;
        align-items: center;
    }
    .st-key-landing_upload_card
    section[data-testid="stFileUploaderDropzone"]:has(> span)
    > span::before {
        color: var(--product-helper);
        font-size: var(--density-body);
        line-height: 1.45;
        content: "or choose files from your computer";
    }
    .st-key-landing_upload_card
    section[data-testid="stFileUploaderDropzone"]:has(> span)
    > span button {
        min-height: 2.24rem;
        padding-right: 1rem !important;
        padding-left: 1rem !important;
        border: 1px solid var(--product-primary) !important;
        background: var(--product-primary) !important;
        color: #ffffff !important;
        box-shadow: 0 4px 11px rgba(37, 99, 235, 0.2);
    }
    .st-key-landing_upload_card
    section[data-testid="stFileUploaderDropzone"]:has(> span)
    > span button:hover {
        border-color: #1d4ed8 !important;
        background: #1d4ed8 !important;
    }
    .st-key-landing_upload_card
    section[data-testid="stFileUploaderDropzone"]:has(> span)
    > span button * {
        color: #ffffff !important;
    }
    .st-key-landing_upload_card
    div[data-testid="stFileUploaderDropzoneInstructions"] {
        order: 4;
        flex: 0 1 auto;
        justify-content: center;
        align-self: center;
        text-align: center;
    }
    .st-key-landing_upload_card
    div[data-testid="stFileUploaderDropzoneInstructions"] span {
        color: var(--product-metadata) !important;
        font-size: 0.66rem;
        line-height: 1.45;
        opacity: 1 !important;
    }
    .st-key-landing_upload_card
    > div[data-testid="stVerticalBlockBorderWrapper"] {
        position: relative;
        overflow: hidden;
        border-color: #bccbe0 !important;
        border-radius: 12px !important;
        background: rgba(255, 255, 255, 0.76) !important;
        box-shadow:
            0 14px 35px rgba(37, 99, 235, 0.15),
            0 5px 13px rgba(15, 23, 42, 0.07);
        padding: 0.9rem !important;
    }
    .st-key-landing_upload_card
    > div[data-testid="stVerticalBlockBorderWrapper"]::before {
        position: absolute;
        z-index: 2;
        top: 0;
        right: 0;
        left: 0;
        height: 3px;
        background: linear-gradient(90deg, #2563eb, #4f46e5);
        content: "";
    }
    .upload-requirement {
        margin: 0.14rem 0 0.62rem;
        color: var(--product-body);
        font-size: var(--density-body);
        font-weight: 500;
        line-height: 1.5;
    }
    .upload-meta,
    .upload-demo-copy {
        color: var(--product-helper);
        font-size: var(--density-helper);
        line-height: 1.5;
        opacity: 1;
    }
    .upload-meta {
        margin-top: 0.52rem;
    }
    .upload-demo-copy {
        margin-top: 0.64rem;
    }
    .st-key-landing_upload_card .panel-title {
        font-size: 1rem;
    }
    .st-key-landing_upload_card .stButton > button {
        padding-right: 0.88rem;
        padding-left: 0.88rem;
    }
    div[data-testid="stFileUploader"] section[data-testid="stFileUploaderDropzone"]:focus-within {
        outline: 3px solid rgba(37, 99, 235, 0.24);
        outline-offset: 2px;
    }
    .stButton > button, .stDownloadButton > button {
        border-radius: 6px;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    .stButton > button,
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    .stDownloadButton > button {
        min-height: var(--density-control);
    }
    .stButton > button:focus-visible,
    .stDownloadButton > button:focus-visible {
        outline: 3px solid rgba(37, 99, 235, 0.28);
        outline-offset: 2px;
    }
    .st-key-combine_files,
    .st-key-continue_approved_merge {
        width: clamp(10rem, 15vw, 12rem);
        margin-top: 1rem;
        margin-left: auto;
    }
    .st-key-combine_files button,
    .st-key-continue_approved_merge button {
        width: 100%;
        min-height: 2.1rem;
    }
    div[data-testid="stHeaderActionElements"]
    a[aria-label="Link to heading"] {
        display: none !important;
    }
    div[data-testid="stSegmentedControl"] {
        margin-top: -0.16rem;
        margin-bottom: 0.2rem;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.clean-page-marker)
    div[data-testid="stSegmentedControl"] {
        display: inline-flex;
        width: auto;
        margin: 0 0 var(--space-4);
        padding: 3px;
        border: 1px solid #cad6e6;
        border-radius: 9px;
        background: #e8eef7;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.clean-page-marker)
    div[data-testid="stSegmentedControl"]
    button[aria-pressed="true"] {
        border-color: var(--product-primary) !important;
        background: linear-gradient(135deg, #2563eb, #315fca) !important;
        color: #ffffff !important;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.20);
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.clean-page-marker)
    div[data-testid="stSegmentedControl"]
    button[aria-pressed="true"] p {
        color: #ffffff !important;
    }
    .st-key-duplicate_preview {
        margin: var(--space-3) 0;
        padding: var(--space-3);
        border: 1px solid #d6e2f0;
        border-radius: 9px;
        background: var(--product-surface-secondary);
    }
    .st-key-missing_value_form {
        margin-top: var(--space-3);
        padding: var(--space-2) var(--space-3);
        border-top: 1px solid var(--product-border-soft);
        border-bottom: 1px solid var(--product-border-soft);
        background: #fbfcfe;
    }
    .insight-type-label {
        display: inline-flex;
        margin-bottom: var(--space-2);
        padding: 0.2rem 0.44rem;
        border-radius: 999px;
        background: #dbeafe;
        color: #1d4ed8;
        font-size: 0.61rem;
        font-weight: 700;
    }
    .insight-stat-strip {
        display: flex;
        flex-wrap: wrap;
        gap: 0.48rem;
        margin: var(--space-3) 0;
        background: transparent;
    }
    .insight-stat {
        flex: 1 1 7.5rem;
        min-width: 0;
        padding: 0.64rem 0.72rem;
        border: 1px solid #d7e1ef;
        border-radius: 8px;
        background: #f7f9fd;
    }
    .insight-stat strong {
        display: block;
        overflow: hidden;
        color: var(--product-text);
        font-size: 0.8rem;
        text-overflow: ellipsis;
        white-space: nowrap;
    }
    .insight-stat span {
        color: var(--product-metadata);
        font-size: 0.62rem;
        line-height: 1.4;
    }
    .insight-interpretation {
        margin: var(--space-3) 0 var(--space-4);
        color: var(--product-body);
        font-size: 0.74rem;
        line-height: 1.55;
    }
    .insight-blank-status {
        margin: -0.15rem 0 var(--space-3);
        padding: 0.42rem 0.62rem;
        border-left: 3px solid #bfdbfe;
        background: #f7faff;
        color: var(--product-helper);
        font-size: var(--density-helper);
        line-height: 1.45;
    }
    .insight-readonly-field {
        display: grid;
        gap: 0.32rem;
    }
    .insight-readonly-field span {
        color: var(--product-body);
        font-size: var(--density-helper);
        font-weight: 500;
        line-height: 1.45;
    }
    .insight-readonly-field strong {
        display: flex;
        min-height: var(--density-control);
        align-items: center;
        padding: 0.3rem 0.75rem;
        border: 1px solid var(--product-border);
        border-radius: 8px;
        background: var(--product-surface-secondary);
        color: var(--product-body);
        font-size: var(--density-body);
        font-weight: 500;
        line-height: 1.35;
    }
    .technical-definition-list {
        display: grid;
        grid-template-columns: minmax(9rem, 1fr) minmax(8rem, 1fr);
        margin: var(--space-2) 0 var(--space-3);
        overflow: hidden;
        border: 1px solid var(--product-border-soft);
        border-radius: 8px;
        background: var(--product-surface);
    }
    .technical-definition-list dt,
    .technical-definition-list dd {
        margin: 0;
        padding: 0.42rem 0.62rem;
        border-bottom: 1px solid #e5ebf3;
        font-size: var(--density-helper);
        line-height: 1.4;
    }
    .technical-definition-list dt {
        color: var(--product-helper);
        font-weight: 500;
    }
    .technical-definition-list dd {
        color: var(--product-body);
        font-weight: 650;
        text-align: right;
    }
    .technical-definition-list dt:last-of-type,
    .technical-definition-list dd:last-of-type {
        border-bottom: 0;
    }
    .st-key-insight_source_comparison {
        margin-top: var(--space-2);
        padding-top: var(--space-2);
        border-top: 1px solid var(--product-border-soft);
    }
    .st-key-insight_primary_surface {
        margin: var(--space-3) 0;
        padding: 1rem;
        border: 1px solid #c9d9ee;
        border-radius: 10px;
        background: linear-gradient(145deg, #eef4ff 0%, #f7f9fd 100%);
    }
    .st-key-insight_review_surface,
    .st-key-chart_output_surface {
        margin-top: var(--space-3);
        padding: var(--space-3);
        border: 1px solid var(--product-border-soft);
        border-radius: 8px;
        background: var(--product-surface-secondary);
    }
    .st-key-download_data_surface
    > div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-download_report_surface
    > div[data-testid="stVerticalBlockBorderWrapper"] {
        position: relative;
        overflow: hidden;
        min-height: 100%;
        border: 1px solid var(--product-border) !important;
        border-radius: var(--product-radius) !important;
        box-shadow: var(--product-shadow);
        padding: 0.8rem !important;
    }
    .st-key-download_data_surface
    > div[data-testid="stVerticalBlockBorderWrapper"] {
        border-top: 3px solid var(--product-primary) !important;
        background: var(--product-surface) !important;
    }
    .st-key-download_report_surface
    > div[data-testid="stVerticalBlockBorderWrapper"] {
        border-top: 3px solid var(--product-indigo) !important;
        background: linear-gradient(145deg, #f7f6ff 0%, #eeefff 100%) !important;
    }
    .surface-heading {
        display: flex;
        gap: 0.6rem;
        align-items: center;
        margin-bottom: var(--space-2);
    }
    .surface-heading-icon {
        display: grid;
        width: 1.88rem;
        height: 1.88rem;
        flex: 0 0 1.88rem;
        place-items: center;
        border-radius: 7px;
        background: #dbeafe;
        color: var(--product-primary);
    }
    .surface-heading-icon.indigo {
        background: #e0e7ff;
        color: var(--product-indigo);
    }
    .surface-heading-icon svg {
        width: 1rem;
        height: 1rem;
        fill: none;
        stroke: currentColor;
        stroke-linecap: round;
        stroke-linejoin: round;
        stroke-width: 1.8;
    }
    .surface-heading h3 {
        margin: 0;
        color: var(--product-text);
        font-size: 0.9rem;
    }
    .report-checklist {
        display: grid;
        gap: 0.36rem;
        margin: var(--space-3) 0;
        color: var(--product-body);
        font-size: var(--density-helper);
    }
    .report-checklist span::before {
        margin-right: 0.4rem;
        color: var(--product-success);
        font-weight: 750;
        content: "✓";
    }
    .st-key-report_contents_checklist {
        margin-top: var(--space-2);
        padding-top: var(--space-2);
        border-top: 1px solid rgba(199, 210, 224, 0.72);
    }
    .st-key-report_contents_checklist div[data-testid="stCheckbox"] {
        min-height: 1.6rem;
    }
    .st-key-report_contents_checklist div[data-testid="stWidgetLabel"] p {
        color: var(--product-body) !important;
        font-size: var(--density-helper) !important;
        font-weight: 500;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-testid="stMarkdownContainer"] p {
        font-size: var(--density-body);
        line-height: 1.5;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-testid="stMarkdownContainer"] h3 {
        margin-top: 0.52rem;
        margin-bottom: 0.32rem;
        color: var(--product-text);
        font-size: 1.1rem;
        line-height: 1.25;
    }
    .block-container
    > div[data-testid="stVerticalBlock"]:has(.internal-page-marker)
    div[data-testid="stMarkdownContainer"] h4 {
        font-size: 0.9rem;
        line-height: 1.3;
    }
    div[data-testid="stDataFrame"] {
        overflow: hidden;
        border: 1px solid var(--product-border);
        border-radius: var(--product-radius);
        background: var(--product-surface);
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
        font-size: var(--density-helper);
    }
    div[data-testid="stDataFrame"] [role="gridcell"],
    div[data-testid="stDataFrame"] [role="columnheader"] {
        font-size: var(--density-helper);
    }
    div[data-testid="stDataFrame"] [role="columnheader"] {
        background: #f3f6fa;
        color: #334155;
        font-weight: 650;
    }
    div[data-baseweb="popover"] {
        max-width: min(25.6rem, calc(100vw - 1.6rem));
    }
    div[data-baseweb="popover"] [role="listbox"] {
        max-height: 14.4rem;
        background: #ffffff;
    }
    div[data-baseweb="popover"] [role="option"] {
        color: #1f2937;
        opacity: 1;
    }
    div[data-baseweb="popover"] [role="option"]:hover {
        background: #eff6ff;
    }
    div[data-baseweb="popover"] [role="option"][aria-selected="true"] {
        background: #dbeafe;
        color: #1e3a8a;
    }
    span[data-baseweb="tag"] {
        background: #dbeafe;
        color: #1e3a8a;
        opacity: 1;
    }
    hr { margin: 0.8rem 0 !important; opacity: 0.18; }
    @media (max-width: 1100px) {
        div[data-testid="stMainBlockContainer"]:has(.landing-intro) {
            min-height: auto;
            display: block;
            padding: clamp(1.6rem, 3.2vh, 2.4rem) 0 2.4rem !important;
        }
        div[data-testid="stHorizontalBlock"]:has(.landing-intro) {
            flex-wrap: wrap;
            gap: var(--space-4) !important;
        }
        div[data-testid="stHorizontalBlock"]:has(.landing-intro)
        > div[data-testid="stColumn"] {
            width: 100% !important;
            min-width: 0 !important;
            flex: 1 1 100% !important;
        }
        .brand-mark {
            width: clamp(2.3rem, 4vw, 2.5rem);
            height: clamp(2.3rem, 4vw, 2.5rem);
            flex-basis: clamp(2.3rem, 4vw, 2.5rem);
        }
    }
    @media (max-width: 760px) {
        div[data-testid="stMainBlockContainer"]:has(.internal-page-marker) {
            padding-top: 2.8rem;
            padding-left: 0.8rem;
            padding-right: 0.8rem;
        }
        .metric-grid { grid-template-columns: 1fr 1fr; }
        .landing-intro h1 { font-size: 1.7rem; }
        .workspace-title { font-size: 1.5rem; }
        .workflow-strip { grid-template-columns: 1fr; gap: 0.6rem; }
        .insight-stat-strip { grid-template-columns: 1fr 1fr; }
        .st-key-combine_files,
        .st-key-continue_approved_merge {
            width: 100%;
        }
    }
    @media (max-width: 520px) {
        .metric-grid { grid-template-columns: 1fr; }
    }
</style>
""", unsafe_allow_html=True)

SAMPLE_FILES = ["sales_north.xlsx", "sales_south.xlsx"]
CHECK_ICON_SVG = (
    '<svg viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="m5 12 4 4L19 6"></path>'
    "</svg>"
)
APP_ICON_SVG = (
    '<svg viewBox="0 0 24 24" aria-hidden="true">'
    '<rect x="4" y="5" width="16" height="14" rx="2"></rect>'
    '<path d="M4 10h16M9 5v14M14 10v9"></path>'
    "</svg>"
)
DOWNLOAD_ICON_SVG = (
    '<svg viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M12 3v11"></path>'
    '<path d="m7 10 5 5 5-5"></path>'
    '<path d="M5 20h14"></path>'
    "</svg>"
)
REPORT_ICON_SVG = (
    '<svg viewBox="0 0 24 24" aria-hidden="true">'
    '<path d="M6 3h9l3 3v15H6z"></path>'
    '<path d="M14 3v4h4M9 12h6M9 16h6"></path>'
    "</svg>"
)
remove_legacy_state(st.session_state)


def reset_workflow() -> None:
    """Clear uploaded files, validation messages, and generated results."""
    reset_session(st.session_state)


def load_sample_data() -> list[tuple[str, pd.DataFrame]]:
    """Read the bundled sample sales files and return as (name, df) pairs."""
    dfs = []
    for fname in SAMPLE_FILES:
        path = SAMPLE_DATA_DIR / fname
        df = pd.read_excel(path)
        dfs.append((fname, df))
    return dfs


def process_uploaded_files(uploaded_files: list) -> None:
    """Validate a changed upload selection and store human-readable outcomes."""
    valid_dfs: list[tuple[str, pd.DataFrame]] = []
    file_errors: list[dict[str, str]] = []

    for uploaded_file in uploaded_files:
        if not is_allowed_file(uploaded_file.name):
            file_errors.append({
                "name": uploaded_file.name,
                "message": "This file type is not supported. Choose an Excel or CSV file.",
            })
            continue
        try:
            df = read_uploaded_file(uploaded_file)
            valid_dfs.append((uploaded_file.name, df))
            logger.info(
                "Successfully read %s (%d rows, %d cols)",
                uploaded_file.name,
                len(df),
                len(df.columns),
            )
        except FileSizeError:
            file_errors.append({
                "name": uploaded_file.name,
                "message": f"This file is larger than the {MAX_FILE_SIZE_MB} MB limit.",
            })
        except FileReadError:
            file_errors.append({
                "name": uploaded_file.name,
                "message": (
                    "We could not open this file. Check that it is a valid "
                    "Excel or CSV file."
                ),
            })

    st.session_state["valid_dfs"] = valid_dfs
    st.session_state["source_schemas"] = build_source_schemas(valid_dfs)
    st.session_state["file_errors"] = file_errors
    st.session_state["_source_mode"] = "uploads"
    st.session_state["_submitted_count"] = len(uploaded_files)


def render_file_preview(valid_dfs: list[tuple[str, pd.DataFrame]]) -> None:
    """Show file metrics and rows only when the user requests them."""
    with st.expander("Preview files", expanded=False):
        tabs = st.tabs([name for name, _ in valid_dfs])
        for tab, (name, df) in zip(tabs, valid_dfs):
            with tab:
                info = describe_dataframe(df, name)
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Rows", info["rows"])
                c2.metric("Columns", info["columns"])
                c3.metric("Missing cells", info["missing_cells"])
                c4.metric("Duplicate rows", info["duplicate_rows"])
                st.dataframe(df.head(8), width="stretch", hide_index=True)


def render_file_errors(file_errors: list[dict[str, str]]) -> None:
    """Place a clear validation message next to each rejected file."""
    for issue in file_errors:
        st.warning(f"**{issue['name']}**\n\n{issue['message']}")


def render_metric_cards(items: list[tuple[str, object]]) -> None:
    """Render compact summary cards."""
    cards = "".join(
        '<div class="metric-card">'
        f'<div class="metric-value">{escape(str(value))}</div>'
        f'<div class="metric-label">{escape(label)}</div>'
        "</div>"
        for label, value in items
    )
    st.markdown(f'<div class="metric-grid">{cards}</div>', unsafe_allow_html=True)


def render_insight_stat_strip(items: list[tuple[str, object]]) -> None:
    """Render one compact supporting statistics strip."""
    stats = "".join(
        '<div class="insight-stat">'
        f"<strong>{escape(str(value))}</strong>"
        f"<span>{escape(label)}</span>"
        "</div>"
        for label, value in items
    )
    st.markdown(
        f'<div class="insight-stat-strip">{stats}</div>',
        unsafe_allow_html=True,
    )


def render_insight_blank_status(message: str) -> None:
    """Render the selected column's blank-state counts as one concise line."""
    st.markdown(
        f'<div class="insight-blank-status">{escape(message)}</div>',
        unsafe_allow_html=True,
    )


def render_technical_definition_list(
    items: list[tuple[str, object]],
) -> None:
    """Render compact selected-column technical details without a wide table."""
    rows = "".join(
        f"<dt>{escape(label)}</dt><dd>{escape(str(value))}</dd>"
        for label, value in items
    )
    st.markdown(
        f'<dl class="technical-definition-list">{rows}</dl>',
        unsafe_allow_html=True,
    )


def render_status_surface(message: str, *, tone: str = "info") -> None:
    """Render a compact status panel that does not rely on color alone."""
    symbol = "✓" if tone == "success" else "i"
    st.markdown(
        f'<div class="status-surface {escape(tone)}">'
        f'<span class="status-icon">{symbol}</span>'
        f"<div>{escape(message)}</div>"
        "</div>",
        unsafe_allow_html=True,
    )


def render_file_list(valid_dfs: list[tuple[str, pd.DataFrame]]) -> None:
    """Render a compact, non-technical list of accepted files."""
    with st.container(border=True, key="file_list_surface"):
        st.markdown('<div class="panel-title">Files ready</div>', unsafe_allow_html=True)
        for name, df in valid_dfs:
            name_col, rows_col, columns_col, status_col = st.columns([4, 1, 1, 1.3])
            name_col.markdown(f"**{name}**")
            rows_col.caption(f"{len(df):,} rows")
            columns_col.caption(f"{len(df.columns)} columns")
            status_col.markdown(
                '<span class="status-ready">Ready to use</span>',
                unsafe_allow_html=True,
            )


def render_uploader_panel(show_demo: bool) -> list:
    """Render the uploader and optional demo action in one focused panel."""
    with st.container(border=True, key="landing_upload_card"):
        st.markdown(
            '<div class="panel-title">Add your spreadsheet files</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="upload-requirement">'
            "Add at least two files to compare and combine them."
            "</div>",
            unsafe_allow_html=True,
        )
        uploaded_files = st.file_uploader(
            "Upload Excel or CSV files",
            type=["xlsx", "csv"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key=(
                f"file_upload_"
                f"{st.session_state.get('uploader_generation', 0)}"
            ),
            help="You can add, remove, or replace files before combining them.",
        )
        st.markdown(
            '<div class="upload-meta">'
            f"Accepted formats: .xlsx and .csv. Maximum size: "
            f"{MAX_FILE_SIZE_MB} MB per file."
            "</div>",
            unsafe_allow_html=True,
        )
        if show_demo:
            st.markdown(
                '<div class="upload-demo-copy">Or try the demo files.</div>',
                unsafe_allow_html=True,
            )
        if show_demo and st.button("Load demo files", type="secondary"):
            sample_dfs = load_sample_data()
            sample_signature = (
                "samples",
                tuple(
                    (name, len(df), tuple(df.columns))
                    for name, df in sample_dfs
                ),
            )
            update_source_signature(st.session_state, sample_signature)
            st.session_state["valid_dfs"] = sample_dfs
            st.session_state["source_schemas"] = build_source_schemas(sample_dfs)
            st.session_state["file_errors"] = []
            st.session_state["_source_mode"] = "samples"
            st.session_state["_submitted_count"] = len(SAMPLE_FILES)
            st.session_state["_upload_signature"] = ()
            st.session_state["uploader_generation"] = (
                st.session_state.get("uploader_generation", 0) + 1
            )
            navigate(st.session_state, SCREEN_REVIEW)
            st.rerun()
    return uploaded_files


def sync_uploaded_files(uploaded_files: list) -> None:
    """Process uploader changes and clear dependent results when needed."""
    upload_signature = tuple(
        (
            uploaded_file.name,
            uploaded_file.size,
            getattr(uploaded_file, "file_id", None),
        )
        for uploaded_file in uploaded_files
    )
    previous_signature = st.session_state.get("_upload_signature")
    source_mode = st.session_state.get("_source_mode")

    if uploaded_files and (
        upload_signature != previous_signature or source_mode != "uploads"
    ):
        update_source_signature(
            st.session_state,
            ("uploads", upload_signature),
        )
        with st.spinner("Checking your files..."):
            process_uploaded_files(uploaded_files)
        st.session_state["_upload_signature"] = upload_signature
        st.rerun()
    if (
        not uploaded_files
        and source_mode == "uploads"
        and previous_signature not in (None, ())
    ):
        update_source_signature(st.session_state, ("uploads", ()))
        st.session_state["valid_dfs"] = []
        st.session_state["file_errors"] = []
        st.session_state["_submitted_count"] = 0
        st.session_state["_upload_signature"] = ()
        navigate(st.session_state, SCREEN_ADD)
        st.rerun()


def render_reset_control() -> None:
    """Render an immediate reset action."""
    st.button(
        "Start over",
        type="secondary",
        on_click=reset_workflow,
        key="workspace_reset",
        width="stretch",
    )


def render_internal_header(screen: str) -> None:
    """Render the compact product identity and integrated reset action."""
    with st.container(key="internal_app_header"):
        st.markdown(
            f'<span class="internal-page-marker {escape(screen)}-page-marker"></span>',
            unsafe_allow_html=True,
        )
        brand_column, reset_column = st.columns(
            [7, 1.25],
            vertical_alignment="center",
        )
        with brand_column:
            st.markdown(
                '<div class="internal-brand">'
                f'<span class="brand-mark compact">{APP_ICON_SVG}</span>'
                "<span>Excel Automation Toolkit</span>"
                "</div>",
                unsafe_allow_html=True,
            )
        with reset_column:
            render_reset_control()


def go_to_screen(screen: str) -> None:
    """Navigate without clearing approved work."""
    if screen == SCREEN_DOWNLOAD:
        st.session_state["export_feedback"] = (
            "Excel and CSV exports are ready to download."
        )
    navigate(st.session_state, screen)


def open_insights_screen() -> None:
    """Open optional insights from the current workflow screen."""
    open_insights(st.session_state)


def return_from_insights_screen() -> None:
    """Return to the screen that opened optional insights."""
    return_from_insights(st.session_state)


def set_clean_view(view: str) -> None:
    """Open a selected cleaning section without clearing its state."""
    st.session_state["clean_view"] = view


def clear_missing_feedback() -> None:
    """Remove stale success feedback when a blank-cell choice changes."""
    st.session_state.pop("missing_feedback", None)
    st.session_state["missing_selections_dirty"] = True


def review_approved_blank_decisions() -> None:
    """Reopen reviewed physical blanks without silently changing decisions."""
    st.session_state["edit_approved_blanks"] = True
    st.session_state["missing_selections_dirty"] = False


def invalidate_current_report() -> None:
    """Invalidate a generated report after a report input changes."""
    invalidate_report_output(st.session_state)


def save_chart_config(config: dict) -> None:
    """Persist a compatible chart without disrupting unrelated report state."""
    if st.session_state.get("chart_config") == config:
        return
    st.session_state["chart_config"] = config
    if "charts" in st.session_state.get("report_sections", []):
        invalidate_current_report()


def report_input_signature(selected_sections: list[str]) -> tuple:
    """Return the state inputs that determine the current report output."""
    chart_signature = (
        repr(st.session_state.get("chart_config"))
        if "charts" in selected_sections
        else None
    )
    return (
        st.session_state.get("source_signature"),
        st.session_state.get("approved_merge_mode"),
        int(st.session_state.get("data_revision", 0)),
        tuple(selected_sections),
        chart_signature,
        getattr(
            st.session_state.get("export_context"),
            "report_signature",
            None,
        ),
    )


def render_flash_success(key: str) -> None:
    """Render a one-rerun success confirmation, then remove it."""
    message = st.session_state.pop(key, None)
    if message:
        st.success(message)


def summarize_missing_actions(actions: list[str]) -> tuple[int, int]:
    """Return counts of filled blank values and removed rows."""
    filled_values = sum(
        int(match.group(1))
        for action in actions
        if (
            match := re.search(
                r"(?:filled|Filled) (\d+) (?:missing|blank) values?",
                action,
            )
        )
    )
    removed_rows = sum(
        int(match.group(1))
        for action in actions
        if (
            match := re.search(
                r"(?:Dropped|Removed) (\d+) "
                r"(?:incomplete )?rows?",
                action,
            )
        )
    )
    return filled_values, removed_rows


def summarize_cleaning_activity() -> tuple[int, int]:
    """Return reconciled fill and incomplete-row-removal counts."""
    audit = st.session_state.get("cleaning_audit", [])
    if audit:
        summary = audit_summary(audit)
        return (
            int(summary["values_filled"]),
            int(summary["incomplete_rows_removed"]),
        )
    return summarize_missing_actions(st.session_state.get("mv_actions", []))


def short_file_label(file_name: str) -> str:
    """Return a concise human label for a source file."""
    stem = file_name.rsplit(".", 1)[0]
    words = stem.replace("-", "_").split("_")
    return words[-1].replace("_", " ").title()


def format_column_list(columns: list[str]) -> str:
    """Format column names for explanatory interface copy."""
    return ", ".join(f"`{column}`" for column in columns)


def format_value(value: object) -> str:
    """Format a value for a compact summary card."""
    if pd.isna(value):
        return "Not available"
    if isinstance(value, (float, int)):
        return f"{value:,.2f}"
    if isinstance(value, pd.Timestamp):
        return value.strftime("%b %d, %Y")
    return str(value)


def build_chart_figure(df: pd.DataFrame, config: dict):
    """Build one chart from the saved chart creator selection."""
    profiles = classify_columns(df)
    if not chart_config_is_compatible(config, profiles):
        raise ValueError(
            "The saved chart configuration is not compatible with the "
            "current column types."
        )
    chart_type = config["type"]
    if chart_type == "Distribution":
        return plot_histogram(df, config["column"])
    if chart_type == "Range and review flags":
        return plot_boxplot(df, config["column"])
    if chart_type == "Category comparison":
        return plot_bar_chart(df, config["column"], top_n=config.get("top_n", 10))
    if chart_type == "Records over time":
        return plot_date_counts(df, config["column"])
    return plot_line_chart(
        df,
        config["date_column"],
        config["value_column"],
        agg=config.get("aggregation", "sum"),
        group_by=config.get("group_by", "Month"),
    )


# ══════════════════════════════════════════════════════════
# START AND FILE REVIEW
# ══════════════════════════════════════════════════════════

valid_dfs = st.session_state.get("valid_dfs", [])
if "current_screen" not in st.session_state:
    if st.session_state.get("merged_df") is not None:
        st.session_state["current_screen"] = SCREEN_CLEAN
    elif len(valid_dfs) >= 2:
        st.session_state["current_screen"] = SCREEN_REVIEW
    else:
        st.session_state["current_screen"] = SCREEN_ADD

current_screen = st.session_state["current_screen"]
if (
    current_screen in {SCREEN_CLEAN, SCREEN_DOWNLOAD, SCREEN_INSIGHTS}
    and st.session_state.get("merged_df") is None
):
    current_screen = SCREEN_REVIEW if len(valid_dfs) >= 2 else SCREEN_ADD
    st.session_state["current_screen"] = current_screen

if current_screen == SCREEN_ADD:
    intro_column, upload_column = st.columns(
        [44, 56],
        gap="large",
        vertical_alignment="center",
    )
    with intro_column:
        st.markdown(
            '<div class="landing-intro">'
            '<span class="landing-eyebrow">Excel automation</span>'
            '<div class="brand-title-row">'
            f'<span class="brand-mark">{APP_ICON_SVG}</span>'
            "<h1>Excel Automation Toolkit</h1>"
            "</div>"
            '<div class="landing-lead">'
            "Prepare spreadsheet data without the repetitive Excel work."
            "</div>"
            '<div class="landing-support">'
            "Combine and clean spreadsheet files in a guided workflow, then "
            "download a ready-to-use result."
            "</div>"
            '<div class="landing-benefits">'
            '<div class="landing-benefit">'
            f'<div class="landing-benefit-icon">{CHECK_ICON_SVG}</div>'
            "<div><strong>Original files stay unchanged</strong>"
            "<span>Your uploaded files are never overwritten.</span></div></div>"
            '<div class="landing-benefit">'
            f'<div class="landing-benefit-icon">{CHECK_ICON_SVG}</div>'
            "<div><strong>Changes only run after approval</strong>"
            "<span>You decide what cleaning actions to apply.</span></div></div>"
            '<div class="landing-benefit">'
            f'<div class="landing-benefit-icon">{CHECK_ICON_SVG}</div>'
            "<div><strong>Files stay in this session</strong>"
            "<span>Work is processed in the running application.</span></div></div>"
            "</div></div>",
            unsafe_allow_html=True,
        )
    with upload_column:
        uploaded_files = render_uploader_panel(show_demo=True)

    sync_uploaded_files(uploaded_files)
    valid_dfs = st.session_state.get("valid_dfs", [])
    file_errors = st.session_state.get("file_errors", [])
    submitted_count = st.session_state.get("_submitted_count", 0)
    render_file_errors(file_errors)

    if valid_dfs:
        render_file_list(valid_dfs)
    elif submitted_count and file_errors:
        st.error("No files could be loaded. Replace the files above and try again.")

    if len(valid_dfs) == 1:
        st.info("Add one more file to combine datasets.")
    elif len(valid_dfs) >= 2:
        _, review_action_column = st.columns([3, 1])
        with review_action_column:
            st.button(
                "Review files",
                type="primary",
                on_click=go_to_screen,
                args=(SCREEN_REVIEW,),
                key="review_files",
                width="stretch",
            )

    st.markdown(
        '<div class="workflow-strip">'
        '<div class="workflow-item"><div class="workflow-marker">1</div>'
        "<div><strong>Add files</strong>"
        "<span>Choose Excel or CSV files.</span></div></div>"
        '<div class="workflow-item"><div class="workflow-marker">2</div>'
        "<div><strong>Clean data</strong>"
        "<span>Approve only the changes you want.</span></div></div>"
        '<div class="workflow-item"><div class="workflow-marker">3</div>'
        "<div><strong>Download results</strong>"
        "<span>Export data or generate a report.</span></div></div>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.stop()

if current_screen == SCREEN_REVIEW:
    valid_dfs = st.session_state.get("valid_dfs", [])
    if len(valid_dfs) < 2:
        navigate(st.session_state, SCREEN_ADD)
        st.rerun()

    render_internal_header(current_screen)
    review_title_column, review_actions_column = st.columns(
        [7, 1.35],
        vertical_alignment="center",
    )
    with review_title_column:
        st.markdown(
            '<h1 class="workspace-title">Review and combine</h1>'
            '<p class="workspace-subtitle">Confirm how the files should be combined.</p>',
            unsafe_allow_html=True,
        )
    with review_actions_column:
        st.button(
            "Change files",
            type="secondary",
            on_click=go_to_screen,
            args=(SCREEN_ADD,),
            key="change_files",
            width="stretch",
        )

    file_names = [name for name, _ in valid_dfs]
    render_file_list(valid_dfs)

    source_signature = st.session_state.get("source_signature")
    if source_signature is None:
        source_signature = (
            "loaded",
            tuple(
                (name, len(df), tuple(df.columns))
                for name, df in valid_dfs
            ),
        )
        st.session_state["source_signature"] = source_signature

    column_sets = [set(df.columns) for _, df in valid_dfs]
    shared_columns = set.intersection(*column_sets)
    all_columns = set.union(*column_sets)
    total_source_rows = sum(len(df) for _, df in valid_dfs)
    compat_scores = compute_schema_compatibility(
        [df for _, df in valid_dfs],
        file_names,
    )
    lowest_score = min(compat_scores.values(), default=1.0)
    column_locations = {
        column: [
            name for name, df in valid_dfs if column in set(df.columns)
        ]
        for column in all_columns
    }
    columns_in_one_file = {
        column: locations[0]
        for column, locations in column_locations.items()
        if len(locations) == 1
    }
    unique_columns_by_file = [
        (
            short_file_label(name),
            sorted(
                column
                for column, location in columns_in_one_file.items()
                if location == name
            ),
        )
        for name, _df in valid_dfs
    ]
    unique_columns_by_file = [
        (label, columns)
        for label, columns in unique_columns_by_file
        if columns
    ]

    with st.container(border=True, key="review_workspace"):
        st.subheader("Ready to combine")
        render_metric_cards([
            ("Files", len(valid_dfs)),
            ("Rows before combining", f"{total_source_rows:,}"),
            ("Shared columns", len(shared_columns)),
            ("Columns in total", len(all_columns)),
        ])

        if lowest_score >= 0.75:
            st.success(
                "These files have mostly matching columns and appear suitable to combine."
            )
        elif lowest_score >= SCHEMA_THRESHOLD:
            st.warning(
                "These files share some columns, but you should review their "
                "differences before combining."
            )
        else:
            st.warning(
                "These files share only a small number of columns and may contain "
                "different kinds of data. Review the column differences before "
                "combining them."
            )

        st.write(f"{len(shared_columns)} of {len(all_columns)} columns match.")
        if len(unique_columns_by_file) == 2:
            first_label, first_columns = unique_columns_by_file[0]
            second_label, second_columns = unique_columns_by_file[1]
            st.write(
                f"The {first_label} file includes "
                f"{format_column_list(first_columns)}, while the {second_label} "
                f"file includes {format_column_list(second_columns)}."
            )
        elif unique_columns_by_file:
            for label, columns in unique_columns_by_file:
                st.write(
                    f"The {label} file includes "
                    f"{format_column_list(columns)}."
                )
        else:
            st.write("Every column appears in every file.")

        st.write(
            "We recommend keeping every column. This preserves the extra columns, "
            "and cells remain blank when a file did not contain that information."
        )

        selected_merge_mode = st.radio(
            "How should columns be handled?",
            ["Keep every column", "Keep only matching columns"],
            captions=[
                (
                    "Recommended. No information is removed. Blank cells appear "
                    "when a column is missing from one file."
                ),
                (
                    "Creates a smaller table but removes columns that are not "
                    "present in every file."
                ),
            ],
            key="merge_mode_v2",
        )
        if st.toggle(
            "Show technical similarity details",
            value=False,
            key="merge_difference_review",
        ):
            st.write(f"Lowest pairwise column similarity: {lowest_score:.0%}")
            if unique_columns_by_file:
                for label, columns in unique_columns_by_file:
                    st.markdown(
                        f"**Only in the {label} file:** "
                        f"{format_column_list(columns)}"
                    )
            else:
                st.write("No columns appear in only one file.")

        requires_recombine = merge_requires_recombine(
            st.session_state,
            source_signature,
            selected_merge_mode,
        )
        if requires_recombine:
            st.warning(
                "Your merge settings changed. Combine the files again to "
                "update the result."
            )

        approved_result_is_current = (
            st.session_state.get("merged_df") is not None
            and not requires_recombine
        )
        if approved_result_is_current:
            st.success("The current combined result uses these files and settings.")
            st.button(
                "Continue to clean",
                type="primary",
                on_click=go_to_screen,
                args=(SCREEN_CLEAN,),
                key="continue_approved_merge",
            )
        elif st.button("Combine files", type="primary", key="combine_files"):
            keep_all = selected_merge_mode == "Keep every column"
            with st.spinner("Combining your files..."):
                merged_df, _schema_warnings = merge_datasets(
                    valid_dfs,
                    keep_all_columns=keep_all,
                )
            approve_merge_configuration(
                st.session_state,
                source_signature,
                selected_merge_mode,
            )
            st.session_state["merged_df"] = merged_df
            st.session_state["clean_view"] = "Check duplicates"
            st.session_state["combine_feedback"] = (
                f"{len(valid_dfs)} files combined successfully."
            )
            st.rerun()

    render_file_preview(valid_dfs)
    st.stop()


# ══════════════════════════════════════════════════════════
# DATA WORKSPACE
# ══════════════════════════════════════════════════════════

merged_df = st.session_state["merged_df"]
valid_dfs = st.session_state.get("valid_dfs", [])
file_names = [name for name, _ in valid_dfs]
source_schemas = st.session_state.get("source_schemas")
if source_schemas is None and valid_dfs:
    source_schemas = build_source_schemas(valid_dfs)
    st.session_state["source_schemas"] = source_schemas
working_df = st.session_state.get("cleaned_df", merged_df)
workspace_missing_summary = detect_missing_values(
    working_df,
    source_schemas,
)
workspace_row_level_blanks = int(
    workspace_missing_summary["row_level_count"].sum()
)
workspace_structural_blanks = int(
    workspace_missing_summary["structural_count"].sum()
)

render_internal_header(current_screen)
workspace_titles = {
    SCREEN_CLEAN: "Clean data",
    SCREEN_DOWNLOAD: "Download results",
    SCREEN_INSIGHTS: "Data insights",
}
st.markdown(
    f'<h1 class="workspace-title">{workspace_titles[current_screen]}</h1>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="dataset-summary">'
    f"{len(working_df):,} rows &nbsp;·&nbsp; "
    f"{len(working_df.columns)} columns &nbsp;·&nbsp; "
    f"{len(valid_dfs)} source files"
    "</div>",
    unsafe_allow_html=True,
)

if current_screen == SCREEN_CLEAN:
    render_flash_success("combine_feedback")
    clean_header_action, _ = st.columns([1, 3])
    with clean_header_action:
        st.button(
            "Back to review",
            type="secondary",
            on_click=go_to_screen,
            args=(SCREEN_REVIEW,),
            key="clean_back_to_review",
            width="stretch",
        )

    clean_sections = [
        "Check duplicates",
        "Review missing values",
        "Review changes",
    ]
    previous_clean_view = st.session_state.get("clean_view")
    if previous_clean_view not in clean_sections:
        st.session_state["clean_view"] = "Check duplicates"

    clean_view = st.segmented_control(
        "Cleaning section",
        clean_sections,
        key="clean_view",
        label_visibility="collapsed",
    )

    if clean_view == "Check duplicates":
        st.markdown("### Check for repeated records")
        st.caption(
            "By default, only completely identical rows are treated as duplicates."
        )
        render_flash_success("duplicate_feedback")

        duplicate_match_mode = st.radio(
            "How should repeated records be matched?",
            ["Match complete rows", "Match selected columns"],
            captions=[
                (
                    "Recommended. A row is considered repeated only when every "
                    "value is identical."
                ),
                "Use this when specific columns identify one unique record.",
            ],
            key="duplicate_match_mode",
        )

        identity_columns: list[str] = []
        selection_ready = duplicate_match_mode == "Match complete rows"
        if duplicate_match_mode == "Match selected columns":
            identity_columns = st.multiselect(
                "Columns that identify a unique record",
                options=list(working_df.columns),
                default=list(DEFAULT_DUPLICATE_COLUMNS),
                key="duplicate_identity_columns_v2",
                placeholder="Choose one or more columns",
            )
            selection_ready = bool(identity_columns)
            if not selection_ready:
                st.info("Select at least one column to check for repeated records.")

        selected_subset = (
            duplicate_subset(identity_columns)
            if duplicate_match_mode == "Match complete rows"
            else identity_columns
        )
        if selection_ready:
            duplicate_mask = working_df.duplicated(
                subset=selected_subset,
                keep=False,
            )
            matching_count = int(duplicate_mask.sum())
            removable_count = int(
                working_df.duplicated(
                    subset=selected_subset,
                    keep="first",
                ).sum()
            )
        else:
            duplicate_mask = pd.Series(False, index=working_df.index)
            matching_count = 0
            removable_count = 0

        if (
            duplicate_match_mode == "Match selected columns"
            and
            len(identity_columns) == 1
            and identity_columns[0].strip().lower() == "product"
        ):
            st.warning(
                "Rows with the same product may represent separate sales. Select "
                "this only if product truly identifies one unique record."
            )

        if removable_count:
            kept_count = matching_count - removable_count
            render_status_surface(
                f"{matching_count:,} matching records found. "
                f"{removable_count:,} will be removed and "
                f"{kept_count:,} will be kept."
            )
            with st.container(key="duplicate_preview"):
                st.markdown("**Preview of repeated records**")
                st.dataframe(
                    working_df.loc[duplicate_mask].head(8),
                    width="stretch",
                    hide_index=True,
                )
        elif identity_columns:
            st.info("No repeated records were found using those columns.")
        elif selection_ready:
            st.success("No repeated records were found.")

        if st.button(
            "Remove selected duplicates",
            type="primary",
            disabled=removable_count == 0,
            key="remove_selected_duplicates",
        ):
            cleaned_dup, dup_report = remove_duplicates(
                working_df,
                subset=selected_subset,
            )
            removed_in_operation = int(dup_report.get("removed", 0))
            previous_duplicate_report = st.session_state.get("dup_report")
            new_duplicate_audit = list(dup_report.get("audit", []))
            if previous_duplicate_report:
                dup_report["before"] = previous_duplicate_report.get(
                    "before",
                    dup_report["before"],
                )
                dup_report["removed"] += int(
                    previous_duplicate_report.get("removed", 0)
                )
                dup_report["audit"] = list(
                    previous_duplicate_report.get("audit", [])
                ) + new_duplicate_audit
            st.session_state["cleaned_df"] = cleaned_dup
            st.session_state["dup_report"] = dup_report
            st.session_state["cleaning_audit"] = (
                st.session_state.get("cleaning_audit", [])
                + new_duplicate_audit
            )
            mark_cleaned_data_changed(st.session_state)
            st.session_state["duplicate_feedback"] = (
                f"{removed_in_operation:,} repeated "
                f"{'row' if removed_in_operation == 1 else 'rows'} removed."
            )
            st.rerun()

        st.divider()
        _, duplicate_next_column = st.columns([3, 1])
        with duplicate_next_column:
            st.button(
                "Continue to missing values",
                type="primary" if removable_count == 0 else "secondary",
                on_click=set_clean_view,
                args=("Review missing values",),
                key="continue_to_missing_values",
                width="stretch",
            )

    elif clean_view == "Review missing values":
        missing_summary = detect_missing_values(
            working_df,
            source_schemas,
        )
        structural_rows = missing_summary[
            missing_summary["structural_count"] > 0
        ]
        missing_rows = missing_summary[
            missing_summary["row_level_count"] > 0
        ]
        missing_feedback = st.session_state.pop("missing_feedback", None)
        if missing_feedback:
            st.success(missing_feedback)
        cleaning_audit = st.session_state.get("cleaning_audit", [])
        pending_override = (
            int(st.session_state.get(
                "missing_pending_override",
                missing_rows["row_level_count"].sum(),
            ))
            if st.session_state.get("missing_selections_dirty")
            else None
        )
        decision_status = missing_decision_summary(
            missing_summary,
            cleaning_audit,
            pending_review_override=pending_override,
            integrity_failures=int(
                st.session_state.get("missing_integrity_failures", 0)
            ),
        )
        duplicate_rows_removed = int(
            (st.session_state.get("dup_report") or {}).get("removed", 0)
        )
        review_summary = build_review_summary(
            missing_summary,
            cleaning_audit,
            duplicate_rows_removed=duplicate_rows_removed,
            pending_review_override=pending_override,
            integrity_failures=int(
                st.session_state.get("missing_integrity_failures", 0)
            ),
        )
        show_completed_review = (
            review_summary.complete
            and not st.session_state.get("edit_approved_blanks", False)
        )

        if show_completed_review:
            handled_values = review_summary.recovered_or_filled_count
            approved_values = review_summary.approved_unchanged_count
            unavailable_values = (
                review_summary.unavailable_from_source_count
            )
            pending_values = review_summary.pending_decision_count
            st.markdown("### Missing-value review complete")
            st.write("All missing-value decisions have been reviewed.")
            render_insight_stat_strip([
                ("Values recovered or filled", f"{handled_values:,}"),
                ("Approved to remain blank", f"{approved_values:,}"),
                ("Unavailable from source", f"{unavailable_values:,}"),
                ("Decisions pending", f"{pending_values:,}"),
            ])

            st.markdown("#### What changed")
            grouped_changes = grouped_audit_change_table(cleaning_audit)
            if not grouped_changes.empty:
                st.dataframe(
                    grouped_changes,
                    width="stretch",
                    hide_index=True,
                )
            else:
                render_status_surface(
                    "No missing-value changes were needed.",
                    tone="success",
                )

            change_preview = cleaning_change_preview(cleaning_audit)
            if not change_preview.empty:
                view_all_changes = st.toggle(
                    "View reviewed records",
                    value=False,
                    key="view_all_missing_changes",
                )
                if view_all_changes:
                    visible_changes = change_preview.copy()
                    visible_changes["Field"] = visible_changes["Field"].map(
                        friendly_column_name
                    )
                    st.dataframe(
                        visible_changes,
                        width="stretch",
                        hide_index=True,
                    )

            unavailable_records: list[dict[str, object]] = []
            for _, structural_row in structural_rows.iterrows():
                column = structural_row["column"]
                for source, count in structural_row[
                    "structural_by_source"
                ].items():
                    unavailable_records.append({
                        "Column": friendly_column_name(column),
                        "Source file": source,
                        "Cell count": int(count),
                        "Status": "Intentionally unavailable",
                    })
            if unavailable_records:
                with st.expander(
                    "Fields unavailable from their source files",
                    expanded=False,
                ):
                    st.write(
                        "These cells are unavailable because their original "
                        "source files did not contain the field. They are not "
                        "pending cleaning decisions."
                    )
                    st.dataframe(
                        pd.DataFrame(unavailable_records),
                        width="stretch",
                        hide_index=True,
                    )
                    st.caption(
                        "Kept blank because this source file did not include "
                        "the field."
                    )
                    st.markdown("**Why are these fields unavailable?**")
                    st.write(
                        "The combined dataset includes fields that were present "
                        "in one source file but not another. These cells are not "
                        "missing entries in the original records, so estimating "
                        "them could create inaccurate business data."
                    )
            if approved_values:
                st.button(
                    "Review approved blank decisions",
                    type="secondary",
                    on_click=review_approved_blank_decisions,
                    key="review_approved_blank_decisions",
                )
        else:
            st.markdown("### Review blank cells")
            st.write(
                "Review the actions below. Nothing will change until you apply them."
            )
            if not structural_rows.empty:
                st.info(
                    "Some cells are unavailable because a source file did not "
                    "include the corresponding field. They will remain blank "
                    "and do not require an action."
                )
            recommendation_labels = {
                "recover_relationship": "Recover from Total ÷ Unit Price",
                "fill_median": "Use the middle value",
                "fill_mode": "Use the most common value",
                "drop_rows": "Remove affected rows",
            }
            selected_actions: dict[str, str] = {}
            custom_values: dict[str, str] = {}
            applied_action_keys: dict[str, str] = {}
            applied_custom_keys: dict[str, str] = {}
            has_unsaved_changes = False
            unsaved_value_count = 0

            with st.container(key="missing_value_form"):
                for row_number, (_, row) in enumerate(missing_rows.iterrows()):
                    column = row["column"]
                    column_label = str(column)
                    column_state_id = (
                        f"{type(column).__name__}:{column!r}"
                    )
                    series = working_df[column]
                    is_numeric = pd.api.types.is_numeric_dtype(series)
                    has_mode = not series.dropna().mode().empty

                    action_options = ["Leave blank"]
                    if is_numeric:
                        action_options.append("Use the middle value")
                    if has_mode:
                        action_options.append("Use the most common value")
                    if is_numeric:
                        action_options.append("Use the average")
                    action_options.extend([
                        "Use a custom value",
                        "Remove affected rows",
                    ])

                    suggested_strategy = recommend_missing_strategy(
                        working_df,
                        column,
                        source_schemas,
                    )
                    suggested_action = recommendation_labels.get(
                        suggested_strategy,
                        "Leave blank",
                    )
                    if (
                        suggested_action == "Recover from Total ÷ Unit Price"
                        and suggested_action not in action_options
                    ):
                        action_options.insert(1, suggested_action)
                    if suggested_action not in action_options:
                        suggested_action = "Leave blank"

                    action_key = f"missing_action_{column_state_id}"
                    applied_action_key = (
                        f"missing_applied_action_{column_state_id}"
                    )
                    applied_custom_key = (
                        f"missing_applied_custom_{column_state_id}"
                    )
                    if st.session_state.get(action_key) not in action_options:
                        st.session_state[action_key] = suggested_action
                    if applied_action_key not in st.session_state:
                        st.session_state[applied_action_key] = None

                    detail_column, action_column = st.columns(
                        [1, 1.15],
                        gap="large",
                        vertical_alignment="center",
                    )
                    with detail_column:
                        st.markdown(f"**{column_label}**")
                        st.caption(
                            f"{int(row['row_level_count']):,} missing "
                            "values requiring attention"
                        )
                    with action_column:
                        selected_action = st.selectbox(
                            f"Action for {column_label}",
                            options=action_options,
                            key=action_key,
                            on_change=clear_missing_feedback,
                        )
                        if selected_action == "Use a custom value":
                            custom_key = (
                                f"custom_missing_{column_state_id}"
                            )
                            custom_values[column] = st.text_input(
                                f"Custom value for {column_label}",
                                key=custom_key,
                                on_change=clear_missing_feedback,
                            )
                        else:
                            custom_values[column] = ""

                        difference_badge = (
                            '<span class="difference-badge">'
                            "Different from recommendation</span>"
                            if selected_action != suggested_action
                            else ""
                        )
                        (
                            recommendation_title,
                            recommendation_detail,
                        ) = recommendation_explanation(suggested_strategy)
                        st.markdown(
                            '<div class="suggestion-text">'
                            '<span class="recommendation-badge">'
                            f"{escape(recommendation_title)}</span>"
                            f" {difference_badge}"
                            '<div class="recommendation-explanation">'
                            f"{escape(recommendation_detail)}</div>"
                            "</div>",
                            unsafe_allow_html=True,
                        )
                        if selected_action == "Use the most common value":
                            st.warning(
                                "This inserts the most frequent value and may "
                                "create incorrect business information."
                            )

                        missing_in_column = int(row["row_level_count"])
                        if (
                            selected_action
                            == "Recover from Total ÷ Unit Price"
                        ):
                            effect_text = (
                                f"{missing_in_column:,} cells will be recovered "
                                "from the validated relationship between Total, "
                                "Quantity, and Unit Price. Formula inputs will "
                                "be recorded in the Cleaning Audit."
                            )
                        elif selected_action == "Use the middle value":
                            effect_text = (
                                f"{missing_in_column:,} cells will be filled "
                                "with the median from the same source file "
                                "where available. Any fallback will be recorded."
                            )
                        elif selected_action == "Use the most common value":
                            effect_text = (
                                f"{missing_in_column:,} cells will be filled "
                                "with the most common value from the same "
                                "source file where available. Any fallback "
                                "will be recorded."
                            )
                        elif selected_action == "Use the average":
                            effect_text = (
                                f"{missing_in_column:,} cells will be filled "
                                "with the average from the same source file "
                                "where available. Any fallback will be recorded."
                            )
                        elif selected_action == "Use a custom value":
                            custom_preview = custom_values[column] or "your custom value"
                            effect_text = (
                                f"{missing_in_column:,} cells will be filled "
                                f"with {custom_preview}."
                            )
                        elif selected_action == "Remove affected rows":
                            effect_text = (
                                f"Up to {missing_in_column:,} rows will be removed."
                            )
                        else:
                            effect_text = (
                                f"{missing_in_column:,} blank cells will remain "
                                "unchanged."
                            )
                        st.markdown(
                            f'<div class="effect-preview">{escape(effect_text)}</div>',
                            unsafe_allow_html=True,
                        )

                    selected_actions[column] = selected_action
                    applied_action_keys[column] = applied_action_key
                    applied_custom_keys[column] = applied_custom_key
                    if (
                        selected_action
                        != st.session_state.get(applied_action_key)
                    ):
                        has_unsaved_changes = True
                        unsaved_value_count += int(row["row_level_count"])
                    elif selected_action == "Use a custom value" and (
                        custom_values[column]
                        != st.session_state.get(applied_custom_key, "")
                    ):
                        has_unsaved_changes = True
                        unsaved_value_count += int(row["row_level_count"])

                    if row_number < len(missing_rows) - 1:
                        st.divider()

            st.session_state["missing_selections_dirty"] = has_unsaved_changes
            st.session_state["missing_pending_override"] = unsaved_value_count
            if has_unsaved_changes:
                st.warning(
                    "You changed one or more actions. Nothing will change until "
                    "you apply them."
                )

            detected_blank_count = int(
                missing_rows["row_level_count"].sum()
            )
            selected_column_count = len(selected_actions)
            if selected_column_count:
                st.caption(
                    f"{detected_blank_count:,} missing values requiring attention "
                    "were detected across "
                    f"{len(missing_rows)} columns. Selected actions will handle "
                    f"{selected_column_count} columns; the remaining blank-cell "
                    "count will be confirmed after you apply them."
                )
            else:
                st.caption(
                    f"{detected_blank_count:,} missing values requiring attention "
                    "were detected. Every "
                    "column is currently set to remain unchanged."
                )

            has_actions_to_apply = bool(selected_actions)
            if st.button(
                (
                    f"Apply {selected_column_count} selected "
                    f"{'action' if selected_column_count == 1 else 'actions'}"
                ),
                type="primary",
                disabled=not has_actions_to_apply,
                key="apply_missing_changes",
            ):
                strategies: dict[str, str] = {}
                validation_errors: list[str] = []
                for column, selected_action in selected_actions.items():
                    if selected_action == "Leave blank":
                        strategies[column] = "leave_blank"
                    elif (
                        selected_action
                        == "Recover from Total ÷ Unit Price"
                    ):
                        strategies[column] = "recover_relationship"
                    elif selected_action == "Use the middle value":
                        strategies[column] = "fill_median"
                    elif selected_action == "Use the most common value":
                        strategies[column] = "fill_mode"
                    elif selected_action == "Use the average":
                        strategies[column] = "fill_mean"
                    elif selected_action == "Remove affected rows":
                        strategies[column] = "drop_rows"
                    elif selected_action == "Use a custom value":
                        custom_value = custom_values.get(column, "")
                        if not custom_value:
                            validation_errors.append(
                                f"Enter a custom value for {column}."
                            )
                        else:
                            strategies[column] = f"fill_value:{custom_value}"

                if validation_errors:
                    for message in validation_errors:
                        st.error(message)
                else:
                    with st.spinner("Applying blank-cell changes..."):
                        cleaning_result = apply_missing_value_strategies(
                            working_df,
                            strategies,
                            source_schemas,
                        )
                        cleaned_missing = cleaning_result.cleaned
                        new_actions = list(cleaning_result.messages)
                    operation_summary = audit_summary(cleaning_result.audit)
                    filled_values = int(operation_summary["values_filled"])
                    approved_values = int(
                        operation_summary["approved_unchanged"]
                    )
                    removed_rows = int(
                        operation_summary["incomplete_rows_removed"]
                    )
                    remaining_summary = detect_missing_values(
                        cleaned_missing,
                        source_schemas,
                    )
                    remaining_decisions = cleaning_result.decision_summary
                    remaining_structural = int(
                        remaining_summary["structural_count"].sum()
                    )
                    st.session_state["missing_feedback"] = (
                        "Selected blank-cell changes were applied. "
                        f"{filled_values:,} filled, "
                        f"{approved_values:,} approved to remain blank, "
                        f"{removed_rows:,} rows removed, and "
                        f"{remaining_decisions.pending_review:,} decisions "
                        "requiring "
                        f"attention remain. {remaining_structural:,} cells are "
                        "unavailable from their source and remain blank."
                    )
                    for column, selected_action in selected_actions.items():
                        st.session_state[
                            applied_action_keys[column]
                        ] = selected_action
                        st.session_state[
                            applied_custom_keys[column]
                        ] = custom_values.get(column, "")
                    st.session_state["cleaned_df"] = cleaned_missing
                    st.session_state["mv_actions"] = (
                        st.session_state.get("mv_actions", []) + new_actions
                    )
                    st.session_state["cleaning_audit"] = (
                        [
                            record
                            for record in st.session_state.get(
                                "cleaning_audit", []
                            )
                            if not (
                                str(record.get("column")) in {
                                    str(column)
                                    for column in selected_actions
                                }
                                and record.get("decision_state") in {
                                    "approved_unchanged",
                                    "failed_or_unresolved",
                                }
                            )
                        ]
                        + [
                            entry.as_record()
                            for entry in cleaning_result.audit
                        ]
                    )
                    st.session_state["missing_selections_dirty"] = False
                    st.session_state["missing_pending_override"] = 0
                    st.session_state["missing_integrity_failures"] = (
                        cleaning_result.integrity_report.severe_count
                    )
                    st.session_state["edit_approved_blanks"] = False
                    mark_cleaned_data_changed(st.session_state)
                    st.rerun()

        st.divider()
        missing_back_column, _, missing_next_column = st.columns([1, 2, 1])
        with missing_back_column:
            st.button(
                "Back to duplicates",
                type="secondary",
                on_click=set_clean_view,
                args=("Check duplicates",),
                key="back_to_duplicates",
                width="stretch",
            )
        with missing_next_column:
            st.button(
                "Continue to review changes",
                type="primary" if decision_status.complete else "secondary",
                on_click=set_clean_view,
                args=("Review changes",),
                key="continue_to_review",
                width="stretch",
            )

    else:
        duplicate_report = st.session_state.get("dup_report")
        missing_actions = st.session_state.get("mv_actions", [])
        duplicates_removed = (
            int(duplicate_report.get("removed", 0))
            if duplicate_report
            else 0
        )
        remaining_missing = detect_missing_values(
            working_df,
            source_schemas,
        )
        pending_override = (
            int(st.session_state.get(
                "missing_pending_override",
                remaining_missing["row_level_count"].sum(),
            ))
            if st.session_state.get("missing_selections_dirty")
            else None
        )
        review_summary = build_review_summary(
            remaining_missing,
            st.session_state.get("cleaning_audit", []),
            duplicate_rows_removed=duplicates_removed,
            pending_review_override=pending_override,
            integrity_failures=int(
                st.session_state.get("missing_integrity_failures", 0)
            ),
        )
        st.markdown("### Review changes")
        summary_text = (
            f"{review_summary.recovered_or_filled_count:,} missing values were "
            "recovered or filled; "
            f"{review_summary.approved_unchanged_count:,} were approved to "
            "remain blank; "
            f"{review_summary.unavailable_from_source_count:,} are unavailable "
            "from their source; "
            f"{review_summary.pending_decision_count:,} decisions are pending."
        )
        panel_heading = (
            "Cleaning review complete"
            if review_summary.complete
            else "Missing-value review is incomplete"
        )
        activity_html = "".join(
            '<div class="activity-item">'
            '<span class="activity-symbol">✓</span>'
            f"<span>{escape(activity)}</span></div>"
            for activity in review_summary.activity_lines()
        )
        st.markdown(
            '<div class="completion-panel">'
            f"<h3>{escape(panel_heading)}</h3>"
            f"<p>{escape(summary_text)}</p>"
            '<div class="activity-list">'
            f"{activity_html}"
            "</div></div>",
            unsafe_allow_html=True,
        )
        render_insight_stat_strip([
            (
                "Repeated rows removed",
                f"{review_summary.duplicate_rows_removed:,}",
            ),
            (
                "Values recovered",
                f"{review_summary.recovered_or_filled_count:,}",
            ),
            (
                "Approved to remain blank",
                f"{review_summary.approved_unchanged_count:,}",
            ),
            (
                "Unavailable from source",
                f"{review_summary.unavailable_from_source_count:,}",
            ),
            (
                "Decisions pending",
                f"{review_summary.pending_decision_count:,}",
            ),
            (
                "Integrity validation",
                review_summary.integrity_status,
            ),
        ])

        st.divider()
        review_back_column, review_optional_column, review_next_column = st.columns(
            [1, 1, 1]
        )
        with review_back_column:
            st.button(
                "Back to missing values",
                type="secondary",
                on_click=set_clean_view,
                args=("Review missing values",),
                key="back_to_missing_values",
                width="stretch",
            )
        with review_optional_column:
            st.button(
                "View data insights",
                type="secondary",
                on_click=open_insights_screen,
                key="clean_view_insights",
                width="stretch",
            )
        with review_next_column:
            st.button(
                "Continue to download",
                type="primary",
                disabled=not review_summary.complete,
                on_click=go_to_screen,
                args=(SCREEN_DOWNLOAD,),
                key="continue_clean_to_download",
                width="stretch",
            )

if current_screen == SCREEN_INSIGHTS:
    analysis_df = st.session_state.get("cleaned_df", merged_df)
    insight_profiles = classify_columns(analysis_df)
    insight_options = grouped_column_options(insight_profiles)
    if st.session_state.get("analysis_column") not in insight_options:
        st.session_state["analysis_column"] = (
            insight_options[0] if insight_options else None
        )
    selected_column = st.selectbox(
        "Choose a column to inspect",
        options=insight_options,
        format_func=lambda column: insight_profiles[column].selector_label,
        key="analysis_column",
        persist_state="session",
    )
    if selected_column is not None:
        profile = insight_profiles[selected_column]
        series = analysis_df[selected_column]
        non_null = series.dropna()
        semantic_label = {
            SEMANTIC_IDENTIFIER: "Identifier",
            SEMANTIC_NUMERIC: "Numeric measure",
            SEMANTIC_CATEGORICAL: "Category",
            SEMANTIC_DATE: "Date",
        }.get(profile.semantic_type, "Unavailable")

        st.markdown(f"### {profile.display_name}")
        st.markdown(
            f'<span class="insight-type-label">{semantic_label}</span>',
            unsafe_allow_html=True,
        )

        chart_config = default_chart_config(profile)
        if profile.semantic_type == SEMANTIC_NUMERIC:
            numeric_view = st.segmented_control(
                "Insight view",
                ["Distribution", "Range and review flags"],
                default="Distribution",
                key="insight_numeric_view",
                label_visibility="collapsed",
                persist_state="session",
            )
            chart_config = {
                "type": numeric_view,
                "column": selected_column,
            }
        active_report_chart_config = chart_config

        with st.container(key="insight_primary_surface"):
            if profile.semantic_type == SEMANTIC_IDENTIFIER:
                st.markdown("#### Record identifiers")
                st.write(
                    "This column identifies records, so averages and "
                    "distributions are not meaningful."
                )
            elif chart_config is not None:
                try:
                    insight_figure = build_chart_figure(
                        analysis_df,
                        chart_config,
                    )
                    st.pyplot(insight_figure, width="stretch")
                except (KeyError, TypeError, ValueError) as exc:
                    logger.exception("Unable to render the selected insight")
                    st.error(
                        "We could not display this insight. Choose another "
                        f"column and try again. ({exc})"
                    )
            else:
                st.info(
                    "This column does not contain supported values to analyze."
                )

        cleaning_audit = st.session_state.get("cleaning_audit", [])
        render_insight_stat_strip(
            selected_column_metrics(
                analysis_df,
                profile,
                source_schemas=source_schemas,
                cleaning_audit=cleaning_audit,
            )
        )
        render_insight_blank_status(
            blank_status_summary(
                analysis_df,
                selected_column,
                source_schemas=source_schemas,
                cleaning_audit=cleaning_audit,
            )
        )

        if profile.semantic_type == SEMANTIC_NUMERIC:
            numeric_values = pd.to_numeric(
                non_null,
                errors="coerce",
            ).dropna()
            column_outliers = (
                summarize_outliers(analysis_df, [selected_column])
                if not numeric_values.empty
                else pd.DataFrame()
            )
            if column_outliers.empty:
                render_status_surface(
                    "No values in this column need extra review.",
                    tone="success",
                )
            else:
                with st.container(key="insight_review_surface"):
                    st.markdown("#### Values worth reviewing")
                    st.write(
                        "These values are far from most others in the same "
                        "column. They are not automatically incorrect."
                    )
                    concise_columns = [
                        column
                        for column in [
                            "Record ID",
                            "Source File",
                            "Value",
                            "Reason",
                        ]
                        if column in column_outliers.columns
                    ]
                    st.dataframe(
                        column_outliers[concise_columns].head(10),
                        width="stretch",
                        hide_index=True,
                    )

            valid_dates = valid_time_columns(
                analysis_df,
                insight_profiles,
                selected_column,
            )
            if valid_dates and not numeric_values.empty:
                compare_over_time = st.toggle(
                    "View this measure over time",
                    value=False,
                    key="insight_measure_over_time",
                    persist_state="session",
                )
                if compare_over_time:
                    defaults = default_time_settings(selected_column)
                    if (
                        st.session_state.get("insight_trend_measure")
                        != selected_column
                    ):
                        st.session_state["insight_trend_measure"] = selected_column
                        st.session_state["insight_trend_date"] = valid_dates[0]
                        st.session_state["insight_trend_group"] = defaults[
                            "Group by"
                        ]
                        st.session_state["insight_trend_calculation"] = defaults[
                            "Calculation"
                        ]
                    if st.session_state.get("insight_trend_date") not in valid_dates:
                        st.session_state["insight_trend_date"] = valid_dates[0]
                    if st.session_state.get("insight_trend_group") not in {
                        "Day",
                        "Week",
                        "Month",
                    }:
                        st.session_state["insight_trend_group"] = defaults[
                            "Group by"
                        ]
                    if st.session_state.get(
                        "insight_trend_calculation"
                    ) not in {
                        "Total",
                        "Average",
                        "Median",
                        "Record count",
                    }:
                        st.session_state["insight_trend_calculation"] = defaults[
                            "Calculation"
                        ]

                    date_control, group_control, calculation_control = st.columns(3)
                    with date_control:
                        if len(valid_dates) == 1:
                            date_column = valid_dates[0]
                            st.session_state["insight_trend_date"] = date_column
                            st.markdown(
                                '<div class="insight-readonly-field">'
                                "<span>Date field</span>"
                                f"<strong>{escape(insight_profiles[date_column].display_name)}</strong>"
                                "</div>",
                                unsafe_allow_html=True,
                            )
                        else:
                            date_column = st.selectbox(
                                "Date field",
                                valid_dates,
                                format_func=lambda column: (
                                    insight_profiles[column].display_name
                                ),
                                key="insight_trend_date",
                                persist_state="session",
                            )
                    with group_control:
                        group_by = st.selectbox(
                            "Group by",
                            ["Day", "Week", "Month"],
                            key="insight_trend_group",
                            persist_state="session",
                        )
                    with calculation_control:
                        calculation_label = st.selectbox(
                            "Calculation",
                            ["Total", "Average", "Median", "Record count"],
                            key="insight_trend_calculation",
                            persist_state="session",
                        )
                    aggregation = {
                        "Total": "sum",
                        "Average": "mean",
                        "Median": "median",
                        "Record count": "count",
                    }[calculation_label]
                    trend_config = {
                        "type": "Trend over time",
                        "date_column": date_column,
                        "value_column": selected_column,
                        "group_by": group_by,
                        "aggregation": aggregation,
                    }
                    active_report_chart_config = trend_config
                    st.pyplot(
                        build_chart_figure(analysis_df, trend_config),
                        width="stretch",
                    )

        if source_comparison_available(analysis_df, profile):
            compare_sources = st.toggle(
                "Compare source files",
                value=False,
                key="insight_compare_sources",
                persist_state="session",
            )
            if compare_sources:
                with st.container(key="insight_source_comparison"):
                    st.markdown("#### Source-file comparison")
                    st.caption(
                        "A concise comparison of the selected column across "
                        "the files that were combined."
                    )
                    comparison_table = source_file_comparison(
                        analysis_df,
                        profile,
                    )
                    st.table(
                        comparison_table,
                        width="stretch",
                        hide_index=True,
                        border="horizontal",
                    )

        if active_report_chart_config is not None:
            save_chart_config(active_report_chart_config)

        st.markdown(
            '<div class="insight-interpretation">'
            f"{escape(insight_interpretation(
                analysis_df,
                profile,
                source_schemas=source_schemas,
                cleaning_audit=st.session_state.get("cleaning_audit", []),
            ))}"
            "</div>",
            unsafe_allow_html=True,
        )

        show_technical_details = st.toggle(
            "View technical statistics",
            value=False,
            key="insight_show_technical",
            persist_state="session",
        )
        if show_technical_details:
            st.markdown(f"#### {profile.display_name} technical details")
            render_technical_definition_list(
                selected_column_technical_details(
                    analysis_df,
                    profile,
                    source_schemas=source_schemas,
                    cleaning_audit=cleaning_audit,
                )
            )
            show_all_statistics = st.toggle(
                "View statistics for all columns",
                value=False,
                key="insight_show_all_statistics",
                persist_state="session",
            )
            if show_all_statistics:
                technical_tables = technical_statistics_by_type(
                    analysis_df,
                    insight_profiles,
                    source_schemas=source_schemas,
                    cleaning_audit=cleaning_audit,
                )
                technical_descriptions = {
                    "Numeric measures": (
                        "Values available, central values, and range."
                    ),
                    "Numeric measure quality": (
                        "Variation, review flags, and blank-status decisions."
                    ),
                    "Categories": (
                        "Category variety and the most common value."
                    ),
                    "Category quality": (
                        "Blank-status decisions for category fields."
                    ),
                    "Dates": (
                        "Available dates, distinct dates, and coverage range."
                    ),
                    "Date quality": (
                        "Blank-status decisions for date fields."
                    ),
                    "Identifiers": (
                        "Record totals, unique values, and repeated values."
                    ),
                    "Identifier quality": (
                        "Blank-status decisions for identifier fields."
                    ),
                }
                for table_name, statistics_table in technical_tables.items():
                    st.markdown(f"#### {table_name}")
                    st.caption(technical_descriptions[table_name])
                    st.table(
                        statistics_table,
                        width="stretch",
                        hide_index=True,
                        border="horizontal",
                    )

    st.divider()
    insights_return_target = st.session_state.get(
        "insights_return_screen",
        SCREEN_CLEAN,
    )
    insights_return_label = (
        "Back to download"
        if insights_return_target == SCREEN_DOWNLOAD
        else "Back to clean"
    )
    insights_return_column, _ = st.columns([1, 3])
    with insights_return_column:
        st.button(
            insights_return_label,
            type="secondary",
            on_click=return_from_insights_screen,
            key="return_from_insights",
            width="stretch",
        )

if current_screen == SCREEN_DOWNLOAD:
    export_df = st.session_state.get("cleaned_df", merged_df)
    export_context = ensure_export_context(
        st.session_state,
        export_df,
        merged_df,
        st.session_state.get("dup_report"),
        st.session_state.get("mv_actions"),
        source_schemas=source_schemas,
        source_files=file_names,
        cleaning_audit=st.session_state.get("cleaning_audit", []),
        workflow_signature=(
            st.session_state.get("source_signature"),
            st.session_state.get("approved_merge_mode"),
            int(st.session_state.get("data_revision", 0)),
        ),
    )
    integrity_report = export_context.integrity_report
    download_allowed = True
    if integrity_report.severe_count:
        st.error(
            f"{integrity_report.severe_count:,} severe relationship "
            f"{'issue requires' if integrity_report.severe_count == 1 else 'issues require'} "
            "review before this dataset can be described as business-ready."
        )
        integrity_rows = pd.DataFrame(integrity_report.issue_records())
        visible_integrity_columns = [
            column
            for column in [
                "affected_record_identifier",
                "source_file",
                "involved_columns",
                "actual_values",
                "expected_relationship",
                "difference",
                "severity",
            ]
            if column in integrity_rows.columns
        ]
        st.dataframe(
            integrity_rows[visible_integrity_columns].rename(columns={
                "affected_record_identifier": "Record ID",
                "source_file": "Source file",
                "involved_columns": "Columns",
                "actual_values": "Actual values",
                "expected_relationship": "Expected relationship",
                "difference": "Difference",
                "severity": "Severity",
            }),
            width="stretch",
            hide_index=True,
        )
        download_allowed = st.checkbox(
            "I acknowledge these relationship issues and want to continue.",
            key="integrity_acknowledged",
        )

    data_column, report_column = st.columns(2)

    with data_column:
        with st.container(border=True, key="download_data_surface"):
            st.markdown(
                '<div class="surface-heading">'
                f'<span class="surface-heading-icon">{DOWNLOAD_ICON_SVG}</span>'
                "<h3>Download cleaned data</h3>"
                "</div>",
                unsafe_allow_html=True,
            )
            st.caption(
                f"{len(export_df):,} rows and {len(export_df.columns)} columns "
                "will be included."
            )
            export_name = st.text_input(
                "File name",
                value=f"cleaned_dataset_{date.today().isoformat()}",
                key="export_filename",
            )
            excel_col, csv_col = st.columns(2)
            excel_export_ready = False
            csv_export_ready = False
            with excel_col:
                try:
                    excel_bytes, excel_name = export_to_excel(
                        export_df,
                        export_name,
                        cleaning_report=export_context.cleaning_summary,
                        cleaning_audit=st.session_state.get(
                            "cleaning_audit",
                            [],
                        ),
                        outlier_df=export_context.values_to_review,
                        source_schemas=source_schemas,
                        source_files=file_names,
                    )
                    excel_export_ready = True
                    st.download_button(
                        "Download Excel",
                        data=excel_bytes,
                        file_name=excel_name,
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        width="stretch",
                        disabled=not download_allowed,
                    )
                except ExportError:
                    logger.exception("Excel download preparation failed")
                    st.error(
                        "We could not prepare the Excel file. Please try again "
                        "or download CSV instead."
                    )
            with csv_col:
                try:
                    csv_bytes, csv_name = export_to_csv(export_df, export_name)
                    csv_export_ready = True
                    st.download_button(
                        "Download CSV",
                        data=csv_bytes,
                        file_name=csv_name,
                        mime="text/csv",
                        width="stretch",
                        disabled=not download_allowed,
                    )
                except ExportError:
                    logger.exception("CSV download preparation failed")
                    st.error("We could not prepare the CSV file. Please try again.")
            export_feedback = st.session_state.pop("export_feedback", None)
            if export_feedback and excel_export_ready and csv_export_ready:
                st.success(export_feedback)
            st.caption(
                "Full cleaning traceability is included in the Excel workbook "
                "and the Word report."
            )

    with report_column:
        with st.container(border=True, key="download_report_surface"):
            st.markdown(
                '<div class="surface-heading">'
                f'<span class="surface-heading-icon indigo">{REPORT_ICON_SVG}</span>'
                "<h3>Generate report</h3>"
                "</div>",
                unsafe_allow_html=True,
            )
            st.write(
                "Create a formatted summary with the dataset overview, cleaning "
                "summary, column statistics, values worth reviewing, and notes."
            )
            st.markdown(
                '<div class="report-checklist">'
                "<strong>Standard report includes</strong>"
                "<span>Dataset overview</span>"
                "<span>Cleaning summary</span>"
                "<span>Key statistics</span>"
                "<span>Values worth reviewing</span>"
                "<span>Notes and limitations</span>"
                "</div>",
                unsafe_allow_html=True,
            )
            saved_report_sections = st.session_state.get(
                "report_sections",
                list(DEFAULT_REPORT_SECTIONS),
            )
            with st.expander("Customize report", expanded=False):
                st.markdown("**Choose report contents**")
                with st.container(key="report_contents_checklist"):
                    checklist_columns = st.columns(2)
                    report_section_order = list(REPORT_SECTION_LABELS.keys())
                    for index, section_key in enumerate(report_section_order):
                        checkbox_key = f"report_section_{section_key}"
                        if checkbox_key not in st.session_state:
                            st.session_state[checkbox_key] = (
                                section_key in saved_report_sections
                            )
                        with checklist_columns[index % 2]:
                            st.checkbox(
                                REPORT_SECTION_LABELS[section_key],
                                key=checkbox_key,
                                on_change=invalidate_current_report,
                                persist_state="session",
                            )

                selected_sections = [
                    section_key
                    for section_key in report_section_order
                    if st.session_state.get(f"report_section_{section_key}", False)
                ]
                st.session_state["report_sections"] = selected_sections
                if "charts" in selected_sections:
                    if st.session_state.get("chart_config"):
                        st.caption("The most recently created chart will be included.")
                    else:
                        st.caption(
                            "Create a chart first if you want a chart in the report."
                        )

            current_report_signature = report_input_signature(selected_sections)
            if (
                st.session_state.get("report_bytes")
                and st.session_state.get("report_input_signature")
                != current_report_signature
            ):
                invalidate_current_report()

            report_error = st.session_state.get("report_error")
            if report_error:
                st.error(report_error)

            report_bytes = st.session_state.get("report_bytes")
            report_is_generating = bool(
                st.session_state.get("report_generating", False)
            )

            if not report_bytes:
                generate_requested = st.button(
                    "Generate report",
                    type="primary",
                    key="generate_report",
                    disabled=report_is_generating or not download_allowed,
                )
                if generate_requested:
                    st.session_state["report_generating"] = True
                    st.session_state.pop("report_error", None)
                    st.rerun()

            if report_is_generating:
                with st.spinner("Generating your report..."):
                    try:
                        report_df = st.session_state.get("cleaned_df", merged_df)
                        chart_images: list[bytes] = []
                        saved_chart_config = st.session_state.get("chart_config")
                        if (
                            "charts" in selected_sections
                            and saved_chart_config
                            and chart_config_is_compatible(
                                saved_chart_config,
                                classify_columns(report_df),
                            )
                        ):
                            report_figure = build_chart_figure(
                                report_df,
                                saved_chart_config,
                            )
                            chart_images.append(figure_to_bytes(report_figure))

                        generated_report = generate_report(
                            df=report_df,
                            stats_df=export_context.statistics,
                            source_files=file_names,
                            cleaning_report=export_context.cleaning_summary,
                            outlier_df=export_context.values_to_review,
                            chart_images=chart_images or None,
                            include_sections=selected_sections,
                            cleaning_audit=st.session_state.get(
                                "cleaning_audit",
                                [],
                            ),
                            source_schemas=source_schemas,
                            dataset_name=st.session_state.get(
                                "export_filename",
                                "Approved cleaned dataset",
                            ),
                        )
                        st.session_state["report_bytes"] = generated_report
                        st.session_state[
                            "report_input_signature"
                        ] = current_report_signature
                    except ReportError as exc:
                        st.session_state["report_error"] = (
                            f"We could not generate the report: {exc}"
                        )
                    except Exception:
                        logger.exception("Report generation failed")
                        st.session_state["report_error"] = (
                            "We could not generate the report. Review the selected "
                            "contents and try again."
                        )
                    finally:
                        st.session_state["report_generating"] = False
                st.rerun()

            report_bytes = st.session_state.get("report_bytes")
            if report_bytes:
                render_status_surface(
                    "Your report is ready.",
                    tone="success",
                )
                st.download_button(
                    "Download report",
                    data=report_bytes,
                    file_name=f"data_report_{date.today().isoformat()}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    ),
                    type="primary",
                    width="stretch",
                    disabled=not download_allowed,
                )

    st.divider()
    download_clean_column, download_explore_column, _ = st.columns([1, 1, 2])
    with download_clean_column:
        st.button(
            "Back to clean",
            type="secondary",
            on_click=go_to_screen,
            args=(SCREEN_CLEAN,),
            key="download_back_to_clean",
            width="stretch",
        )
    with download_explore_column:
        st.button(
            "View data insights",
            type="secondary",
            on_click=open_insights_screen,
            key="download_to_explore",
            width="stretch",
        )
