"""Business Excel Automation Toolkit — Streamlit entry point."""

from datetime import date

import streamlit as st
import pandas as pd

from config.settings import SCHEMA_THRESHOLD as DEFAULT_SCHEMA_THRESHOLD
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
from src.exporter import (
    export_to_excel,
    export_to_csv,
    ExportError,
)
from src.report_generator import (
    generate_report,
    ReportError,
)
from src.utils import is_allowed_file

logger = setup_logger(__name__)

# ── Page configuration ───────────────────────────────────
st.set_page_config(
    page_title="Excel Automation Toolkit",
    page_icon="📊",
    layout="wide",
)

# ── Sidebar ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📊 Excel Automation Toolkit")
    st.markdown(
        "A professional tool for merging, cleaning, analyzing, and exporting "
        "Excel and CSV files. No data is ever sent to external servers."
    )
    st.divider()

    has_files = st.session_state.get("_has_files", False)
    if has_files:
        workflow_steps = [
            "Upload files",
            "Review & merge",
            "Clean data",
            "Analyze",
            "Charts",
            "Export",
            "Report",
        ]
        completed = sum([
            1,  # files uploaded
            1 if st.session_state.get("merged_df") is not None else 0,
            1 if st.session_state.get("cleaned_df") is not None else 0,
            1 if st.session_state.get("stats_generated") else 0,
            1 if st.session_state.get("charts_viewed") else 0,
            1 if st.session_state.get("exported") else 0,
            1 if st.session_state.get("report_generated") else 0,
        ])
        st.caption(f"Progress: {completed}/{len(workflow_steps)} steps")
        for i, step in enumerate(workflow_steps):
            done = i < completed
            icon = "✅" if done else "⬜"
            st.caption(f"{icon} {step}")
        st.divider()
    else:
        st.caption("Upload files in Step 1 to begin.")

    if st.button("🔄 Start Over", use_container_width=True, type="secondary"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

st.title("📊 Business Excel Automation Toolkit")
st.markdown(
    "This tool helps you combine multiple Excel/CSV files into one clean dataset, "
    "spot data quality issues, create charts, and export professional reports. "
    "Follow the numbered steps below."
)

# ── Step 1: Upload files ─────────────────────────────────
st.divider()
st.subheader("Step 1: Upload Files")
st.caption(
    "Upload one or more Excel (`.xlsx`) or CSV (`.csv`) files. "
    "Each file must be under 50 MB. You can upload files with different "
    "columns — the tool will align them automatically."
)

uploaded_files = st.file_uploader(
    "Choose Excel or CSV files",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Select multiple files at once. Max 50 MB per file. "
         "Accepted formats: .xlsx, .csv",
)

st.session_state["_has_files"] = uploaded_files is not None and len(uploaded_files) > 0

if not uploaded_files:
    st.info("Upload files above to begin.")
    st.stop()

# ── Validate and read files ──────────────────────────────
valid_dfs: list[tuple[str, pd.DataFrame]] = []
errors: list[str] = []

progress_bar = st.progress(0, text="Reading files...")

for i, uploaded_file in enumerate(uploaded_files):
    progress_bar.progress(
        (i + 1) / len(uploaded_files),
        text=f"Processing: {uploaded_file.name}",
    )

    if not is_allowed_file(uploaded_file.name):
        errors.append(f"❌ `{uploaded_file.name}` — file type not supported.")
        continue

    try:
        df = read_uploaded_file(uploaded_file)
        valid_dfs.append((uploaded_file.name, df))
        logger.info("Successfully read %s (%d rows, %d cols)",
                     uploaded_file.name, len(df), len(df.columns))
    except FileSizeError as e:
        errors.append(f"❌ {e}")
    except FileReadError as e:
        errors.append(f"❌ {e}")

progress_bar.empty()

# ── Display errors first ─────────────────────────────────
if errors:
    st.subheader("Issues")
    for err in errors:
        st.warning(err)

# ── Display per-file summaries ───────────────────────────
if valid_dfs:
    st.subheader("Uploaded Files")
    st.caption(f"{len(valid_dfs)} of {len(uploaded_files)} file(s) loaded successfully.")

    tab_names = [name for name, _ in valid_dfs]
    tabs = st.tabs(tab_names)

    for tab, (name, df) in zip(tabs, valid_dfs):
        with tab:
            info = describe_dataframe(df, name)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Rows", info["rows"])
            col2.metric("Columns", info["columns"])
            col3.metric("Missing Cells", info["missing_cells"])
            col4.metric("Duplicate Rows", info["duplicate_rows"])

            st.dataframe(df.head(100), use_container_width=True)

    # ── Step 2: Review & Merge ───────────────────────────────
    st.divider()
    st.subheader("Step 2: Review & Merge Datasets")
    st.caption(
        "Compare column structures across your files, then merge them into "
        "one combined dataset. Rows are stacked vertically. "
        "Columns that exist in some files but not others are preserved "
        "(filled with empty values where missing)."
    )

    # Show schema comparison
    with st.expander("Compare column structure across files", expanded=True):
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

        # Compatibility scores
        file_names = [name for name, _ in valid_dfs]
        compat_scores = compute_schema_compatibility(
            [df for _, df in valid_dfs], file_names
        )

        threshold = st.slider(
            "Compatibility threshold",
            min_value=0.0,
            max_value=1.0,
            value=DEFAULT_SCHEMA_THRESHOLD,
            step=0.05,
            help="Controls how strictly the tool flags files as potentially unrelated. "
                 "Lower = more tolerant (fewer warnings, higher risk of merging unrelated data). "
                 "Higher = stricter (more warnings, safer merging).",
        )

        st.markdown("**Schema Compatibility**")
        st.caption(
            f"Jaccard similarity between column sets. "
            f"1.0 = same columns, 0.0 = no overlap. "
            f"Threshold: {threshold:.2f}"
        )

        compat_rows = []
        for (f1, f2), score in compat_scores.items():
            compat_rows.append({
                "File A": f1,
                "File B": f2,
                "Compatibility": score,
                "Status": "✅ Compatible" if score >= threshold else "⚠️ Low compatibility",
            })
        compat_df = pd.DataFrame(compat_rows)
        st.dataframe(compat_df, use_container_width=True, hide_index=True)

        low_compat = [r for r in compat_rows if r["Compatibility"] < threshold]
        if low_compat:
            for r in low_compat:
                st.warning(
                    f"**{r['File A']}** and **{r['File B']}** have only "
                    f"{r['Compatibility']*100:.0f}% column overlap "
                    f"(below {threshold:.0%} threshold). "
                    "These files might not belong together. Review before merging."
                )

    # Merge controls
    merge_col, warn_col = st.columns([1, 2])
    with merge_col:
        merge_clicked = st.button("Merge Datasets", type="primary", use_container_width=True)
    with warn_col:
        st.info(
            "Files with different columns will be merged as a union — "
            "all columns are preserved. Missing values will be NaN.",
            icon="ℹ️",
        )

    if merge_clicked:
        with st.spinner("Merging datasets..."):
            merged_df, schema_warnings = merge_datasets(valid_dfs)

        # Show schema warnings
        if schema_warnings:
            st.subheader("Schema Warnings")
            for warning in schema_warnings:
                st.warning(warning.message)

        # Show merge summary
        file_names = [name for name, _ in valid_dfs]
        summary = get_merge_summary(merged_df, file_names)

        st.subheader("Merged Dataset")
        col1, col2 = st.columns(2)
        col1.metric("Total Rows", summary["total_rows"])
        col2.metric("Total Columns", summary["total_columns"])

        # Per-file contribution
        contrib_data = pd.DataFrame(summary["per_file"])
        contrib_data["percent"] = (
            contrib_data["rows"] / summary["total_rows"] * 100
        ).round(1).apply(lambda x: f"{x}%")
        st.dataframe(
            contrib_data.rename(
                columns={"file": "Source File", "rows": "Rows", "percent": "Share"}
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.dataframe(merged_df.head(100), use_container_width=True)

        # Store in session state for downstream milestones
        st.session_state["merged_df"] = merged_df

    # ── Step 3: Clean Data (only after merge) ────────────────
    if "merged_df" in st.session_state:
        df_to_clean = st.session_state["merged_df"].copy()

        st.divider()
        st.subheader("Step 3: Clean Data")
        st.caption(
            "Remove duplicate rows and handle missing values before analysis. "
            "For each column with missing values, a recommended strategy is shown "
            "based on the data type and how many values are missing. "
            "**You can override any recommendation.** "
            "Missing-value strategies execute in this order: fill first (e.g. mean, median, "
            "mode, zero, custom), then drop any remaining rows with missing values. "
            "This ensures you keep as much data as possible."
        )

        clean_tab1, clean_tab2, clean_tab3 = st.tabs(
            ["Duplicates", "Missing Values", "Cleaning Report"]
        )

        # ── Duplicates tab ──────────────────────────────────
        with clean_tab1:
            dup_count = int(df_to_clean.duplicated().sum())
            st.metric("Duplicate rows detected", dup_count)

            if dup_count > 0:
                subset_cols = st.multiselect(
                    "Consider duplicates only on these columns (leave empty for all columns)",
                    options=list(df_to_clean.columns),
                    default=None,
                    help="If you select columns here, only these columns are checked for duplicates. "
                         "Useful when a subset of columns should be unique together.",
                )
                if st.button("Remove Duplicates", type="primary"):
                    cleaned_dup, dup_report = remove_duplicates(
                        df_to_clean,
                        subset=subset_cols if subset_cols else None,
                    )
                    st.session_state["cleaned_df"] = cleaned_dup
                    st.session_state["dup_report"] = dup_report
                    st.success(f"Removed {dup_report['removed']} duplicate rows.")
            else:
                st.info("No duplicate rows found.")

            if "dup_report" in st.session_state:
                r = st.session_state["dup_report"]
                st.caption(f"Duplicates removed: {r['removed']} | Before: {r['before']} → After: {r['after']}")

        # ── Missing Values tab ──────────────────────────────
        with clean_tab2:
            df_for_missing = st.session_state.get("cleaned_df", df_to_clean)
            missing_summary = detect_missing_values(df_for_missing)
            missing_with_issues = missing_summary[missing_summary["missing_count"] > 0]

            if missing_with_issues.empty:
                st.info("No missing values found.")
            else:
                st.dataframe(
                    missing_with_issues.rename(columns={
                        "column": "Column",
                        "missing_count": "Missing",
                        "missing_pct": "Missing %",
                        "dtype": "Data Type",
                    }),
                    use_container_width=True,
                    hide_index=True,
                )

                st.markdown("**Configure handling per column**")
                strategies = {}
                for _, row in missing_with_issues.iterrows():
                    col = row["column"]
                    recommended = recommend_strategy(df_for_missing[col])
                    chosen = st.selectbox(
                        f"`{col}` ({row['missing_count']} missing, {row['missing_pct']}%)",
                        options=[
                            "drop_rows",
                            "fill_mean",
                            "fill_median",
                            "fill_mode",
                            "fill_zero",
                            "fill_value:custom",
                        ],
                        index=["drop_rows", "fill_mean", "fill_median",
                               "fill_mode", "fill_zero", "fill_value:custom"]
                        .index(recommended),
                        key=f"strategy_{col}",
                    )
                    if chosen == "fill_value:custom":
                        custom_val = st.text_input(
                            f"Custom value for `{col}`",
                            value="",
                            key=f"custom_val_{col}",
                        )
                        strategies[col] = f"fill_value:{custom_val}" if custom_val else "drop_rows"
                    else:
                        strategies[col] = chosen

                if st.button("Apply Missing Value Handling", type="primary"):
                    cleaned_mv, mv_actions = handle_missing_values(
                        df_for_missing, strategies
                    )
                    st.session_state["cleaned_df"] = cleaned_mv
                    st.session_state["mv_actions"] = mv_actions
                    st.success(f"Applied {len(mv_actions)} strategy(ies).")

            if "mv_actions" in st.session_state:
                for action in st.session_state["mv_actions"]:
                    st.caption(f"✓ {action}")

        # ── Cleaning Report tab ─────────────────────────────
        with clean_tab3:
            original = st.session_state["merged_df"]
            cleaned = st.session_state.get("cleaned_df", original)
            dup_report = st.session_state.get("dup_report")
            mv_actions = st.session_state.get("mv_actions")

            report = generate_cleaning_report(
                original, cleaned, dup_report, mv_actions
            )

            col1, col2, col3 = st.columns(3)
            col1.metric("Rows", report["rows_before"], delta=report["rows_after"] - report["rows_before"])
            col2.metric("Duplicates Removed", report["duplicates_removed"])
            col3.metric("Missing Values", report["total_missing_before"],
                        delta=report["total_missing_after"] - report["total_missing_before"])

            if report["missing_actions"]:
                st.markdown("**Actions taken**")
                for action in report["missing_actions"]:
                    st.caption(f"• {action}")

            if report["duplicates_removed"] > 0 or report["missing_actions"]:
                with st.expander("Preview: Before vs After (first 10 rows)"):
                    before_st, after_st = st.columns(2)
                    with before_st:
                        st.caption("Before cleaning")
                        st.dataframe(original.head(10), use_container_width=True)
                    with after_st:
                        st.caption("After cleaning")
                        st.dataframe(cleaned.head(10), use_container_width=True)

    # ── Step 4: Analyze (only after merge + optional cleaning) ─
    if "merged_df" in st.session_state:
        df_to_analyze = st.session_state.get("cleaned_df", st.session_state["merged_df"])

        st.divider()
        st.subheader("Step 4: Analyze — Summary Statistics")
        st.caption(
            "Per-column statistics to help you understand your data at a glance. "
            "Which statistics are shown depends on the detected data type "
            "(numeric: mean, median; categorical: most common value; date: earliest/latest)."
        )

        stats_df = generate_summary_statistics(df_to_analyze)
        st.dataframe(stats_df, use_container_width=True, hide_index=True)
        st.session_state["stats_generated"] = True

        # ── Outlier detection (separate section) ──────────────
        with st.expander("Outlier Detection"):
            st.markdown(
                "Outliers are values that fall far outside the typical range. "
                "They may be data entry errors or legitimate extreme values. "
                "**No data is modified here** — outliers are flagged for review."
            )

            numeric_cols = [
                col for col in df_to_analyze.columns
                if pd.api.types.is_numeric_dtype(df_to_analyze[col])
            ]

            if not numeric_cols:
                st.info("No numeric columns available for outlier detection.")
            else:
                selected_cols = st.multiselect(
                    "Select numeric columns to inspect",
                    options=numeric_cols,
                    default=numeric_cols[:3],
                )

                if selected_cols and st.button("Find Outliers", type="primary"):
                    outlier_df = summarize_outliers(df_to_analyze, selected_cols)
                    st.session_state["outlier_df"] = outlier_df

                    if outlier_df.empty:
                        st.success("No outliers detected in the selected columns.")
                    else:
                        st.warning(
                            f"Found {len(outlier_df)} outlier(s) in "
                            f"{outlier_df['Column'].nunique()} column(s)."
                        )
                        st.dataframe(outlier_df, use_container_width=True, hide_index=True)

    # ── Step 5: Charts (only after merge) ────────────────────
    if "merged_df" in st.session_state:
        df_to_chart = st.session_state.get("cleaned_df", st.session_state["merged_df"])

        st.divider()
        st.subheader("Step 5: Charts")
        st.caption(
            "Visualize your data. Pick a chart type and column. "
            "Charts are read-only — nothing is modified."
        )

        # Determine column types for the selector
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

    # ── Step 6: Export (only after merge) ────────────────────
    if "merged_df" in st.session_state:
        st.session_state["exported"] = True
        df_to_export = st.session_state.get("cleaned_df", st.session_state["merged_df"])

        st.divider()
        st.subheader("Step 6: Export Cleaned Data")

        if df_to_export.empty:
            st.warning("The dataset is empty. Nothing to export.")
        else:
            st.caption(
                f"Download the cleaned dataset ({len(df_to_export):,} rows, "
                f"{len(df_to_export.columns)} columns)."
            )

            export_name = st.text_input(
                "File name (without extension)",
                value=f"cleaned_dataset_{date.today().isoformat()}",
            )

            col1, col2 = st.columns(2)

            with col1:
                try:
                    excel_bytes, excel_name = export_to_excel(df_to_export, export_name)
                    size_kb = len(excel_bytes) / 1024
                    st.download_button(
                        label=f"Download Excel ({size_kb:.1f} KB)",
                        data=excel_bytes,
                        file_name=excel_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except ExportError as e:
                    st.error(f"Excel export failed: {e}")
                except Exception as e:
                    st.error(f"Unexpected Excel export error: {e}")

            with col2:
                try:
                    csv_bytes, csv_name = export_to_csv(df_to_export, export_name)
                    size_kb = len(csv_bytes) / 1024
                    st.download_button(
                        label=f"Download CSV ({size_kb:.1f} KB)",
                        data=csv_bytes,
                        file_name=csv_name,
                        mime="text/csv",
                        use_container_width=True,
                    )
                except ExportError as e:
                    st.error(f"CSV export failed: {e}")
                except Exception as e:
                    st.error(f"Unexpected CSV export error: {e}")

    # ── Step 7: Generate Report (only after merge) ───────────
    if "merged_df" in st.session_state:
        df_for_report = st.session_state.get("cleaned_df", st.session_state["merged_df"])
        stats_df = generate_summary_statistics(df_for_report)

        st.divider()
        st.subheader("Step 7: Generate Word Report")
        st.caption(
            "Create a professional Word document (.docx) with dataset overview, "
            "statistics, cleaning actions, outliers, and charts."
        )

        # Section selection
        section_options = {
            "overview": "Dataset Overview",
            "quality": "Data Quality",
            "statistics": "Summary Statistics",
            "outliers": "Outlier Summary",
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

        # Chart selection (only if Charts section is selected)
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

            chart_col1, chart_col2 = st.columns(2)
            with chart_col1:
                selected_hist = st.multiselect("Histograms", num_cols, default=num_cols[:2])
                selected_bar = st.multiselect("Bar charts", cat_cols, default=cat_cols[:2])
            with chart_col2:
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

        # Build optional data
        cleaning_report_data = {
            "rows_before": len(st.session_state["merged_df"]),
            "rows_after": len(df_for_report),
            "duplicates_removed": st.session_state.get("dup_report", {}).get("removed", 0),
            "total_missing_before": int(st.session_state["merged_df"].isna().sum().sum()),
            "total_missing_after": int(df_for_report.isna().sum().sum()),
            "missing_actions": st.session_state.get("mv_actions", []),
        }
        outlier_data = st.session_state.get("outlier_df", None)
        source_files = [name for name, _ in valid_dfs]

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
                    size_kb = len(report_bytes) / 1024
                    st.download_button(
                        label=f"Download Report ({size_kb:.1f} KB)",
                        data=report_bytes,
                        file_name=f"data_quality_report_{date.today().isoformat()}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                except ReportError as e:
                    st.error(f"Report generation failed: {e}")
                except Exception as e:
                    st.error(f"Unexpected report error: {e}")
else:
    st.error("No files could be loaded. Check the errors above.")
