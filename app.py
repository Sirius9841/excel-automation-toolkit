"""Business Excel Automation Toolkit — Streamlit entry point."""

from datetime import date
from pathlib import Path
from io import BytesIO

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
    get_merge_summary,
    compute_schema_compatibility,
    remove_duplicates,
    detect_missing_values,
    handle_missing_values,
    recommend_strategy,
    generate_cleaning_report,
)
from src.analyzer import (
    generate_summary_statistics,
    detect_outliers_iqr,
    summarize_outliers,
)
from src.visualizer import (
    plot_histogram,
    plot_boxplot,
    plot_bar_chart,
    plot_line_chart,
    figure_to_bytes,
)
from src.exporter import export_to_excel, export_to_csv, ExportError
from src.report_generator import generate_report, ReportError
from src.utils import is_allowed_file, format_bytes

logger = setup_logger(__name__)

# ── Page configuration ───────────────────────────────────
st.set_page_config(
    page_title="Excel Automation Toolkit",
    page_icon="📊",
    layout="wide",
)

# ── Custom CSS ───────────────────────────────────────────
st.markdown("""
<style>
    .block-container { max-width: 920px; padding-top: 1rem; }
    .main-title { font-size: 1.6rem; font-weight: 700; margin-bottom: 0.25rem; }
    .tagline { font-size: 0.95rem; color: #666; margin-bottom: 1.25rem; }
    .card { border: 1px solid rgba(128,128,128,0.2); border-radius: 8px; padding: 1.25rem 1.5rem; margin-bottom: 1rem; background: rgba(128,128,128,0.03); }
    .card-compact { border: 1px solid rgba(128,128,128,0.15); border-radius: 6px; padding: 0.75rem 1rem; margin-bottom: 0.5rem; }
    .trust-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem; font-size: 0.82rem; margin: 0.75rem 0; }
    .trust-item { padding: 0.5rem; border-radius: 4px; background: rgba(128,128,128,0.04); }
    .step-label { font-size: 0.82rem; line-height: 1.3; }
    .file-badge { font-size: 0.82rem; color: #555; }
    .workflow-bar { display: flex; align-items: center; gap: 0.25rem; flex-wrap: wrap; font-size: 0.82rem; margin: 0.5rem 0; }
    .wf-done { color: #2e7d32; font-weight: 500; }
    .wf-active { color: #1565c0; font-weight: 600; }
    .wf-pending { color: #9e9e9e; }
    .wf-arrow { color: #ccc; margin: 0 0.15rem; }
    hr { margin: 0.75rem 0 !important; opacity: 0.3; }
    div[data-testid="stExpander"] { border: 1px solid rgba(128,128,128,0.15); border-radius: 6px; margin-bottom: 0.5rem; }
</style>
""", unsafe_allow_html=True)

SAMPLE_FILES = ["sales_north.xlsx", "sales_south.xlsx"]


# ── Workflow helpers ─────────────────────────────────────
def get_step_index() -> int:
    """Return 0-based index of the current workflow step."""
    if st.session_state.get("merged_df") is None:
        uploaded = st.session_state.get("_has_data", False)
        return 0 if not uploaded else 1
    return 2


STEPS = ["Upload", "Compare", "Merge", "Clean", "Analyze", "Export", "Report"]


def render_workflow_bar(compact: bool = False) -> None:
    """Render a horizontal workflow progress bar."""
    current = get_step_index()
    parts = []
    for i, s in enumerate(STEPS):
        if i < current:
            parts.append(f'<span class="wf-done">\u2713 {s}</span>')
        elif i == current:
            parts.append(f'<span class="wf-active">\u25b6 {s}</span>')
        else:
            parts.append(f'<span class="wf-pending">\u25cb {s}</span>')
        if i < len(STEPS) - 1:
            parts.append(f'<span class="wf-arrow">\u2192</span>')
    st.markdown(f'<div class="workflow-bar">{" ".join(parts)}</div>',
                unsafe_allow_html=True)


def load_sample_data() -> list[tuple[str, pd.DataFrame]]:
    """Read the bundled sample sales files and return as (name, df) pairs."""
    dfs = []
    for fname in SAMPLE_FILES:
        path = SAMPLE_DATA_DIR / fname
        df = pd.read_excel(path)
        dfs.append((fname, df))
    return dfs


# ── Sidebar ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("**Excel Automation Toolkit**")
    st.markdown(
        '<span style="font-size:0.82rem;color:#777;">'
        "Merge, clean, and export Excel &amp; CSV files</span>",
        unsafe_allow_html=True,
    )

    has_data = st.session_state.get("_has_data", False)

    if has_data:
        st.divider()
        render_workflow_bar()

        st.divider()
        confirm = st.session_state.get("_confirm_reset", False)
        if not confirm:
            if st.button("Start Over", use_container_width=True):
                st.session_state["_confirm_reset"] = True
                st.rerun()
        else:
            st.warning("Clear all progress?")
            c1, c2 = st.columns(2)
            if c1.button("Yes, reset", use_container_width=True):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
            if c2.button("Cancel", use_container_width=True):
                st.session_state["_confirm_reset"] = False
                st.rerun()
    else:
        st.divider()
        st.caption("Upload files or load sample data to begin.")

# ══════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════

has_data = st.session_state.get("_has_data", False)

# ── Landing page (no data yet) ───────────────────────────
if not has_data:
    st.markdown('<div class="main-title">Excel Automation Toolkit</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<div class="tagline">'
        "Combine, clean, and export data from multiple Excel and CSV files.</div>",
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Upload files",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="landing_upload",
        help="",
    )

    if uploaded_files and len(uploaded_files) > 0:
        st.session_state["_landing_files"] = uploaded_files
        st.session_state["_has_data"] = True
        st.rerun()

    st.markdown(
        f'<div class="file-badge" style="margin-top:-0.75rem;margin-bottom:1rem">'
        f'Supported: <strong>.xlsx</strong> and <strong>.csv</strong> &middot; '
        f'Max {MAX_FILE_SIZE_MB} MB per file</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div style="text-align:center;margin:0.5rem 0">or</div>',
                unsafe_allow_html=True)
    if st.button("Load sample sales files", type="secondary", use_container_width=True):
        dfs = load_sample_data()
        st.session_state["valid_dfs"] = dfs
        st.session_state["_sample_loaded"] = True
        st.session_state["_has_data"] = True
        st.rerun()

    # Trust points
    st.markdown(
        '<div class="trust-grid">'
        '<div class="trust-item">\u2705 Original files are never modified</div>'
        '<div class="trust-item">\u2705 All cleaning requires your confirmation</div>'
        '<div class="trust-item">\u2705 Files are processed in memory, not stored permanently</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    # Workflow visual
    st.markdown("---")
    wf_labels = ["Upload", "Compare", "Merge", "Clean", "Analyze", "Export", "Report"]
    wf_cols = st.columns(len(wf_labels) * 2 - 1)
    for i, label in enumerate(wf_labels):
        with wf_cols[i * 2]:
            if i == 0:
                st.markdown(f"**{label}**")
            else:
                st.markdown(f'<span style="color:#999">{label}</span>',
                            unsafe_allow_html=True)
        if i < len(wf_labels) - 1:
            with wf_cols[i * 2 + 1]:
                st.markdown(f'<span style="color:#ddd">→</span>',
                            unsafe_allow_html=True)

    st.stop()

# ── Retrieve data from session state ─────────────────────
sample_loaded = st.session_state.get("_sample_loaded", False)

if sample_loaded:
    valid_dfs = st.session_state["valid_dfs"]
else:
    uploaded_files = st.session_state.get("_landing_files", [])

# ══════════════════════════════════════════════════════════
# PROCESS FILES
# ══════════════════════════════════════════════════════════

errors: list[str] = []

if not sample_loaded:
    uploaded_files = st.session_state.get("_landing_files", [])
    progress_bar = st.progress(0, text="Reading files...")
    valid_dfs_batch: list[tuple[str, pd.DataFrame]] = []
    for i, uploaded_file in enumerate(uploaded_files):
        progress_bar.progress(
            (i + 1) / len(uploaded_files),
            text=f"Processing: {uploaded_file.name}",
        )
        if not is_allowed_file(uploaded_file.name):
            errors.append(f"'{uploaded_file.name}' — file type not supported.")
            continue
        try:
            df = read_uploaded_file(uploaded_file)
            valid_dfs_batch.append((uploaded_file.name, df))
            logger.info("Successfully read %s (%d rows, %d cols)",
                         uploaded_file.name, len(df), len(df.columns))
        except FileSizeError as e:
            errors.append(str(e))
        except FileReadError as e:
            errors.append(str(e))
    progress_bar.empty()
    st.session_state["valid_dfs"] = valid_dfs_batch
    st.session_state["errors"] = errors

valid_dfs = st.session_state.get("valid_dfs", [])
errors = st.session_state.get("errors", [])
file_names = [name for name, _ in valid_dfs]

if errors:
    st.subheader("Issues")
    for err in errors:
        st.warning(err)

if not valid_dfs:
    st.error("No files could be loaded.")
    st.stop()

# ══════════════════════════════════════════════════════════
# WORKFLOW BODY
# ══════════════════════════════════════════════════════════

current_step = get_step_index()
has_merged = st.session_state.get("merged_df") is not None

# ── Heading ──────────────────────────────────────────────
st.markdown('<div class="main-title">Excel Automation Toolkit</div>',
            unsafe_allow_html=True)
render_workflow_bar()

# ══════════════════════════════════════════════════════════
# STEP: UPLOAD (collapsible summary when done)
# ══════════════════════════════════════════════════════════
with st.expander("Uploaded files", expanded=not has_merged):
    st.caption(f"{len(valid_dfs)} file(s) loaded.")
    tab_names = [name for name, _ in valid_dfs]
    tabs = st.tabs(tab_names)
    for tab, (name, df) in zip(tabs, valid_dfs):
        with tab:
            info = describe_dataframe(df, name)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Rows", info["rows"])
            c2.metric("Columns", info["columns"])
            c3.metric("Missing cells", info["missing_cells"])
            c4.metric("Duplicate rows", info["duplicate_rows"])
            st.dataframe(df.head(100), use_container_width=True)

    if has_merged:
        st.success(f"Merged from {len(valid_dfs)} files — {file_names}")
    else:
        st.markdown("**Next step:** Review file structures below.")

# ══════════════════════════════════════════════════════════
# STEP: COMPARE & MERGE
# ══════════════════════════════════════════════════════════
merge_expanded = not has_merged

with st.expander("Compare & Merge", expanded=merge_expanded):
    if not has_merged:
        # Schema comparison
        with st.container():
            st.markdown("**File structure similarity**")
            st.caption(
                "How similar the column sets are across your files. "
                "A higher score means more shared columns."
            )
            schema_data = []
            for name, df in valid_dfs:
                for col in df.columns:
                    schema_data.append({"File": name, "Column": col, "Dtype": str(df[col].dtype)})
            schema_df = pd.DataFrame(schema_data)
            pivot = schema_df.pivot_table(
                index="Column", columns="File", values="Dtype",
                aggfunc="first", fill_value="—",
            )
            st.dataframe(pivot, use_container_width=True)

            file_names = [name for name, _ in valid_dfs]
            compat_scores = compute_schema_compatibility(
                [df for _, df in valid_dfs], file_names
            )
            threshold = st.slider(
                "Similarity threshold",
                min_value=0.0, max_value=1.0, value=SCHEMA_THRESHOLD, step=0.05,
                help="Lower values = more tolerant of column differences. "
                     "0.3 means columns present in 30%+ of files are included.",
            )
            compat_rows = []
            for (f1, f2), score in compat_scores.items():
                status = "\u2705 Compatible" if score >= threshold else "\u26a0\ufe0f Low similarity"
                compat_rows.append({"File A": f1, "File B": f2, "Similarity": score, "Status": status})
            compat_df = pd.DataFrame(compat_rows)
            st.dataframe(compat_df, use_container_width=True, hide_index=True)

            low_compat = [r for r in compat_rows if r["Similarity"] < threshold]
            if low_compat:
                for r in low_compat:
                    st.warning(
                        f"**{r['File A']}** and **{r['File B']}** have only "
                        f"{r['Similarity']*100:.0f}% column overlap "
                        f"(below {threshold:.0%} threshold). "
                        "These files may not belong together."
                    )

        # Merge
        st.markdown("**Merge strategy**")
        merge_mode = st.radio(
            "Choose how to combine columns:",
            ["Keep all columns (union)", "Keep matching columns only (intersection)"],
            horizontal=True,
            index=0,
            help="'Keep all columns' preserves every column found in any file. "
                 "'Keep matching only' keeps columns common to all files.",
        )
        keep_all = merge_mode.startswith("Keep all")

        col1, col2 = st.columns([1, 2])
        with col1:
            merge_clicked = st.button(
                "Merge Datasets", type="primary", use_container_width=True,
                disabled=len(valid_dfs) < 2,
            )
        with col2:
            if len(valid_dfs) < 2:
                st.info("At least 2 files are needed for merging.")
            else:
                st.info(
                    "Rows are stacked vertically. "
                    "A 'source_file' column identifies each row's origin.",
                    icon="\u2139\ufe0f",
                )

        if merge_clicked:
            with st.spinner("Merging datasets..."):
                merged_df, schema_warnings = merge_datasets(valid_dfs, keep_all_columns=keep_all)

            if schema_warnings:
                for warning in schema_warnings:
                    st.warning(warning.message)

            summary = get_merge_summary(merged_df, file_names)
            st.success(
                f"Merged {summary['total_rows']:,} rows \u00d7 "
                f"{summary['total_columns']} columns "
                f"from {len(valid_dfs)} files."
            )

            contrib_data = pd.DataFrame(summary["per_file"])
            contrib_data["share"] = (
                contrib_data["rows"] / summary["total_rows"] * 100
            ).round(1).apply(lambda x: f"{x}%")
            st.dataframe(
                contrib_data.rename(columns={"file": "File", "rows": "Rows", "share": "Share"}),
                use_container_width=True, hide_index=True,
            )

            st.dataframe(merged_df.head(100), use_container_width=True)
            st.session_state["merged_df"] = merged_df
            st.rerun()
    else:
        merged_df = st.session_state["merged_df"]
        summary = get_merge_summary(merged_df, file_names)
        c1, c2 = st.columns(2)
        c1.metric("Total rows", summary["total_rows"])
        c2.metric("Total columns", summary["total_columns"])
        st.caption(f"Merged from {len(valid_dfs)} file(s): {', '.join(file_names)}")

# ══════════════════════════════════════════════════════════
# DOWNSTREAM STEPS (only after merge)
# ══════════════════════════════════════════════════════════
if "merged_df" in st.session_state:
    merged_df = st.session_state["merged_df"]

    # ══════════════════════════════════════════════════════
    # STEP: CLEAN
    # ══════════════════════════════════════════════════════
    df_to_clean = merged_df.copy()
    has_cleaned = st.session_state.get("cleaned_df") is not None
    clean_expanded = not has_cleaned

    with st.expander("Clean Data", expanded=clean_expanded):
        st.caption(
            "Remove duplicates and fill or drop missing values. "
            "All actions require confirmation."
        )

        clean_tab1, clean_tab2, clean_tab3 = st.tabs(
            ["Duplicates", "Missing Values", "Cleaning Report"]
        )

        with clean_tab1:
            dup_count = int(df_to_clean.duplicated().sum())
            st.metric("Duplicate rows found", dup_count)
            if dup_count > 0:
                subset_cols = st.multiselect(
                    "Check duplicates only in these columns (leave empty for all columns)",
                    options=list(df_to_clean.columns), default=None,
                    help="Only these columns are used to identify duplicates.",
                )
                if st.button("Remove Duplicates", type="primary"):
                    cleaned_dup, dup_report = remove_duplicates(
                        df_to_clean, subset=subset_cols if subset_cols else None,
                    )
                    st.session_state["cleaned_df"] = cleaned_dup
                    st.session_state["dup_report"] = dup_report
                    st.success(f"Removed {dup_report['removed']} duplicate row(s).")
                    st.rerun()
            else:
                st.info("No duplicate rows found.")

            if "dup_report" in st.session_state:
                r = st.session_state["dup_report"]
                st.caption(f"Removed: {r['removed']} | Before: {r['before']} \u2192 After: {r['after']}")

        with clean_tab2:
            df_for_missing = st.session_state.get("cleaned_df", df_to_clean)
            missing_summary = detect_missing_values(df_for_missing)
            missing_with_issues = missing_summary[missing_summary["missing_count"] > 0]

            if missing_with_issues.empty:
                st.info("No missing values found.")
            else:
                st.dataframe(
                    missing_with_issues.rename(columns={
                        "column": "Column", "missing_count": "Missing",
                        "missing_pct": "Missing %", "dtype": "Data Type",
                    }),
                    use_container_width=True, hide_index=True,
                )
                st.markdown("**Strategy per column**")
                strategies = {}
                for _, row in missing_with_issues.iterrows():
                    col = row["column"]
                    recommended = recommend_strategy(df_for_missing[col])
                    chosen = st.selectbox(
                        f"`{col}` ({row['missing_count']} missing, {row['missing_pct']}%)",
                        options=[
                            "drop_rows", "fill_mean", "fill_median",
                            "fill_mode", "fill_zero", "fill_value:custom",
                        ],
                        index=["drop_rows", "fill_mean", "fill_median",
                               "fill_mode", "fill_zero", "fill_value:custom"]
                        .index(recommended),
                        key=f"strategy_{col}",
                    )
                    if chosen == "fill_value:custom":
                        custom_val = st.text_input(
                            f"Custom value for `{col}`", value="",
                            key=f"custom_val_{col}",
                        )
                        strategies[col] = f"fill_value:{custom_val}" if custom_val else "drop_rows"
                    else:
                        strategies[col] = chosen

                if st.button("Apply Missing Value Handling", type="primary"):
                    cleaned_mv, mv_actions = handle_missing_values(df_for_missing, strategies)
                    st.session_state["cleaned_df"] = cleaned_mv
                    st.session_state["mv_actions"] = mv_actions
                    st.success(f"Applied {len(mv_actions)} strategy(ies).")
                    st.rerun()

            if "mv_actions" in st.session_state:
                for action in st.session_state["mv_actions"]:
                    st.caption(f"\u2713 {action}")

        with clean_tab3:
            original = merged_df
            cleaned = st.session_state.get("cleaned_df", original)
            dup_report = st.session_state.get("dup_report")
            mv_actions = st.session_state.get("mv_actions")
            report = generate_cleaning_report(original, cleaned, dup_report, mv_actions)

            c1, c2, c3 = st.columns(3)
            c1.metric("Rows", report["rows_before"],
                       delta=report["rows_after"] - report["rows_before"])
            c2.metric("Duplicates removed", report["duplicates_removed"])
            c3.metric("Missing values", report["total_missing_before"],
                       delta=report["total_missing_after"] - report["total_missing_before"])

            if report["missing_actions"]:
                st.markdown("**Actions taken**")
                for action in report["missing_actions"]:
                    st.caption(f"\u2022 {action}")

            if report["duplicates_removed"] > 0 or report["missing_actions"]:
                with st.expander("Preview before vs after (first 10 rows)"):
                    bc, ac = st.columns(2)
                    with bc:
                        st.caption("Before")
                        st.dataframe(original.head(10), use_container_width=True)
                    with ac:
                        st.caption("After")
                        st.dataframe(cleaned.head(10), use_container_width=True)

    # ══════════════════════════════════════════════════════
    # STEP: ANALYZE
    # ══════════════════════════════════════════════════════
    df_to_analyze = st.session_state.get("cleaned_df", merged_df)

    with st.expander("Analyze", expanded=True):
        stats_df = generate_summary_statistics(df_to_analyze)
        st.dataframe(stats_df, use_container_width=True, hide_index=True)
        st.session_state["stats_generated"] = True

        with st.expander("Unusual values (outliers)"):
            st.markdown(
                "Values that fall far outside the normal range. "
                "**No data is modified here** — unusual values are flagged for review."
            )
            numeric_cols = [
                col for col in df_to_analyze.columns
                if pd.api.types.is_numeric_dtype(df_to_analyze[col])
            ]
            if not numeric_cols:
                st.info("No numeric columns available for this check.")
            else:
                selected_cols = st.multiselect(
                    "Select numeric columns to check",
                    options=numeric_cols, default=numeric_cols[:3],
                )
                if selected_cols and st.button("Find Unusual Values", type="primary"):
                    outlier_df = summarize_outliers(df_to_analyze, selected_cols)
                    st.session_state["outlier_df"] = outlier_df
                    if outlier_df.empty:
                        st.success("No unusual values detected.")
                    else:
                        st.warning(
                            f"Found {len(outlier_df)} unusual value(s) in "
                            f"{outlier_df['Column'].nunique()} column(s)."
                        )
                        st.dataframe(outlier_df, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════
    # STEP: CHARTS
    # ══════════════════════════════════════════════════════
    df_to_chart = st.session_state.get("cleaned_df", merged_df)

    with st.expander("Charts", expanded=True):
        numeric_cols = [c for c in df_to_chart.columns
                        if pd.api.types.is_numeric_dtype(df_to_chart[c])]
        categorical_cols = [c for c in df_to_chart.columns
                            if not pd.api.types.is_numeric_dtype(df_to_chart[c])
                            and not pd.api.types.is_datetime64_any_dtype(df_to_chart[c])]
        date_cols = [c for c in df_to_chart.columns
                     if pd.api.types.is_datetime64_any_dtype(df_to_chart[c])]

        chart_type = st.radio(
            "Chart type",
            ["Histogram", "Box Plot", "Bar Chart", "Line Chart (date + value)"],
            horizontal=True,
        )

        if chart_type in ("Histogram", "Box Plot"):
            if not numeric_cols:
                st.info("No numeric columns available.")
            else:
                col = st.selectbox("Select column", numeric_cols)
                if chart_type == "Histogram":
                    bins = st.slider("Number of bins", 5, 100, 20)
                    fig = plot_histogram(df_to_chart, col, bins=bins)
                else:
                    fig = plot_boxplot(df_to_chart, col)
                st.pyplot(fig)
                st.session_state["charts_viewed"] = True

        elif chart_type == "Bar Chart":
            if not categorical_cols:
                st.info("No categorical columns available.")
            else:
                col = st.selectbox("Select column", categorical_cols)
                top_n = st.slider("Show top N values", 5, 50, 10)
                fig = plot_bar_chart(df_to_chart, col, top_n=top_n)
                st.pyplot(fig)
                st.session_state["charts_viewed"] = True

        elif chart_type == "Line Chart (date + value)":
            if not date_cols:
                st.info("No date columns available.")
            elif not numeric_cols:
                st.info("No numeric columns available to plot over time.")
            else:
                date_col = st.selectbox("Select date column", date_cols)
                val_col = st.selectbox("Select value column", numeric_cols)
                agg = st.radio("Aggregation", ["sum", "mean"], horizontal=True)
                fig = plot_line_chart(df_to_chart, date_col, val_col, agg=agg)
                st.pyplot(fig)
                st.session_state["charts_viewed"] = True

    # ══════════════════════════════════════════════════════
    # STEP: EXPORT
    # ══════════════════════════════════════════════════════
    df_to_export = st.session_state.get("cleaned_df", merged_df)

    with st.expander("Export", expanded=True):
        st.session_state["exported"] = True

        if df_to_export.empty:
            st.warning("The dataset is empty. Nothing to export.")
        else:
            st.caption(
                f"Download the processed data ({len(df_to_export):,} rows, "
                f"{len(df_to_export.columns)} columns)."
            )
            export_name = st.text_input(
                "File name (without extension)",
                value=f"cleaned_dataset_{date.today().isoformat()}",
            )
            c1, c2 = st.columns(2)
            with c1:
                try:
                    excel_bytes, excel_name = export_to_excel(df_to_export, export_name)
                    st.download_button(
                        label=f"Download Excel ({len(excel_bytes) / 1024:.1f} KB)",
                        data=excel_bytes, file_name=excel_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except ExportError as e:
                    st.error(f"Excel export failed: {e}")
                except Exception as e:
                    st.error(f"Unexpected error: {e}")
            with c2:
                try:
                    csv_bytes, csv_name = export_to_csv(df_to_export, export_name)
                    st.download_button(
                        label=f"Download CSV ({len(csv_bytes) / 1024:.1f} KB)",
                        data=csv_bytes, file_name=csv_name,
                        mime="text/csv", use_container_width=True,
                    )
                except ExportError as e:
                    st.error(f"CSV export failed: {e}")
                except Exception as e:
                    st.error(f"Unexpected error: {e}")

    # ══════════════════════════════════════════════════════
    # STEP: REPORT
    # ══════════════════════════════════════════════════════
    df_for_report = st.session_state.get("cleaned_df", merged_df)
    stats_df = generate_summary_statistics(df_for_report)

    with st.expander("Generate Report", expanded=True):
        st.caption("Create a Word document (.docx) with overview, statistics, and charts.")

        section_options = {
            "overview": "Dataset Overview",
            "quality": "Data Quality",
            "statistics": "Summary Statistics",
            "outliers": "Unusual Values",
            "charts": "Charts",
            "preview": "Data Preview",
            "methodology": "Methodology & Limitations",
        }
        selected_sections = st.multiselect(
            "Include sections",
            options=list(section_options.keys()),
            default=list(section_options.keys()),
            format_func=lambda k: section_options[k],
        )

        chart_images: list[bytes] = []
        if "charts" in selected_sections:
            st.markdown("**Select charts to include**")
            num_cols = [c for c in df_for_report.columns
                        if pd.api.types.is_numeric_dtype(df_for_report[c])]
            cat_cols = [c for c in df_for_report.columns
                        if not pd.api.types.is_numeric_dtype(df_for_report[c])
                        and not pd.api.types.is_datetime64_any_dtype(df_for_report[c])]
            date_cols = [c for c in df_for_report.columns
                         if pd.api.types.is_datetime64_any_dtype(df_for_report[c])]
            cc1, cc2 = st.columns(2)
            with cc1:
                selected_hist = st.multiselect("Histograms", num_cols, default=num_cols[:2])
                selected_bar = st.multiselect("Bar charts", cat_cols, default=cat_cols[:2])
            with cc2:
                if date_cols and num_cols:
                    sel_date = st.selectbox("Line chart: date", date_cols)
                    sel_val = st.selectbox("Line chart: value", num_cols)
                    line_chart_selected = st.checkbox("Include line chart", value=True)
                else:
                    line_chart_selected = False

            for col in selected_hist:
                fig = plot_histogram(df_for_report, col)
                chart_images.append(figure_to_bytes(fig))
            for col in selected_bar:
                fig = plot_bar_chart(df_for_report, col)
                chart_images.append(figure_to_bytes(fig))
            if line_chart_selected and date_cols and num_cols:
                fig = plot_line_chart(df_for_report, sel_date, sel_val)
                chart_images.append(figure_to_bytes(fig))

        cleaning_report_data = {
            "rows_before": len(merged_df),
            "rows_after": len(df_for_report),
            "duplicates_removed": st.session_state.get("dup_report", {}).get("removed", 0),
            "total_missing_before": int(merged_df.isna().sum().sum()),
            "total_missing_after": int(df_for_report.isna().sum().sum()),
            "missing_actions": st.session_state.get("mv_actions", []),
        }
        outlier_data = st.session_state.get("outlier_df", None)
        source_files = file_names

        if st.button("Generate Report", type="primary"):
            with st.spinner("Generating report..."):
                try:
                    report_bytes = generate_report(
                        df=df_for_report,
                        stats_df=stats_df,
                        source_files=source_files,
                        cleaning_report=cleaning_report_data,
                        outlier_df=outlier_data,
                        chart_images=chart_images if chart_images else None,
                        include_sections=selected_sections,
                    )
                    st.session_state["report_generated"] = True
                    st.download_button(
                        label=f"Download Report ({len(report_bytes) / 1024:.1f} KB)",
                        data=report_bytes,
                        file_name=f"data_quality_report_{date.today().isoformat()}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                except ReportError as e:
                    st.error(f"Report generation failed: {e}")
                except Exception as e:
                    st.error(f"Unexpected error: {e}")
else:
    st.error("No files could be loaded.")
