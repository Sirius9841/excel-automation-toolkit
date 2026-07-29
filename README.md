# Business Excel Automation Toolkit

Combine, clean, analyze, and report on multiple Excel and CSV files — without leaving your browser. Built with Streamlit.

## Overview

Business teams often receive data scattered across multiple spreadsheets with inconsistent columns. Manually merging them in Excel is error-prone and repetitive.

This toolkit lets you upload files, review their schemas, merge them, clean duplicates and missing values, explore statistics and outliers, create charts, and export a professional Word report — all through a web interface. No data is ever sent to an external server.

## Key Features

- **Multiple file uploads** — Excel (`.xlsx`) and CSV (`.csv`), any number at once
- **Schema validation** — see column types and missing columns side by side
- **Compatibility scoring** — Jaccard similarity between file column sets
- **Union or intersection merging** — keep all columns or only common ones
- **Source-file traceability** — every row is tagged with its origin file
- **Duplicate detection** — preview and remove, optionally by column subset
- **Missing-value handling** — per-column strategy: mean, median, mode, zero, custom value, or row drop
- **Cleaning audit trail** — before/after report of every action
- **Summary statistics** — per-column stats based on detected data type
- **Outlier flagging** — IQR-based detection, never auto-removed
- **Charts** — histogram, box plot, bar chart, line chart (date + value)
- **Excel and CSV export** — in-memory, formatted, UTF-8 BOM for Windows Excel
- **Word report generation** — configurable sections, embedded charts, styled tables

## Workflow

```
Upload -> Validate -> Compare -> Merge -> Clean -> Analyze -> Visualize -> Export -> Report
```

Each step is a numbered section in the app. The sidebar tracks your progress.

## Screenshots

<!-- Screenshot placeholders — replace with actual screenshots -->

| Section | Description |
|---|---|
| `screenshots/upload.png` | File upload widget with three files loaded, showing per-file tabs with row/column metrics |
| `screenshots/schema-comparison.png` | Pivot table comparing column dtypes across files, compatibility scores, and a warning for low-compatibility files |
| `screenshots/merge.png` | Merge button, per-file contribution table, merged data preview |
| `screenshots/cleaning.png` | Duplicate removal controls, missing-value strategy selectors per column, apply button |
| `screenshots/cleaning-report.png` | Before/after row counts, duplicates removed, missing values delta, action log |
| `screenshots/statistics.png` | Summary statistics table showing mean, median, min, max, unique counts per column |
| `screenshots/outliers.png` | Outlier detection results listing column, row index, value, Q1, Q3, IQR |
| `screenshots/charts.png` | Histogram, box plot, bar chart, and line chart examples |
| `screenshots/export.png` | Excel and CSV download buttons with file sizes |
| `screenshots/report.png` | Word report section with section-selection checkboxes and generate button |

To create screenshots: open the app at `http://localhost:8501`, upload the sample files (`sales_north.xlsx` + `sales_south.xlsx`), and step through each section. Capture each relevant area.

## Architecture

```
excel-automation-toolkit/
├── app.py                  # Streamlit entry point — UI layout and workflow orchestration
├── config/
│   ├── __init__.py
│   └── settings.py         # Paths, file-size limits, logging config, defaults
├── src/
│   ├── __init__.py
│   ├── file_handler.py     # Upload, read, validate, describe files
│   ├── data_processor.py   # Merge, schema comparison, duplicate removal, missing values
│   ├── analyzer.py         # Column classification, summary statistics, IQR outlier detection
│   ├── visualizer.py       # Matplotlib charts — histogram, box plot, bar, line
│   ├── exporter.py         # Excel (.xlsx) and CSV (utf-8-sig) export to bytes
│   ├── report_generator.py # Word (.docx) report with tables, charts, methodology
│   ├── logger_setup.py     # Centralised console + file logger
│   └── utils.py            # File-extension check, byte formatting
├── tests/                  # 87 pytest tests across all modules
├── sample_data/            # Demo files + generator script
│   ├── generate_samples.py
│   ├── sales_north.xlsx
│   ├── sales_south.xlsx
│   └── employees.xlsx
├── requirements.txt
└── README.md
```

### Module responsibilities

| Module | Role |
|---|---|
| `app.py` | Streamlit UI, ties all modules into a workflow |
| `file_handler.py` | Reads uploaded files, validates size/type, describes DataFrame stats |
| `data_processor.py` | Merges DataFrames (union/intersection), schema compatibility (Jaccard), duplicate removal, missing-value handling with fill-before-drop execution |
| `analyzer.py` | Classifies column types, generates type-aware summary statistics, detects outliers via IQR |
| `visualizer.py` | Pure-function chart generation (read-only, never modifies input) |
| `exporter.py` | In-memory Excel (openpyxl, styled) and CSV (UTF-8 BOM) export |
| `report_generator.py` | Generates a 7-section Word document with tables, chart images, and factual-only conclusions |
| `logger_setup.py` | Configures console (INFO+) and file (DEBUG+) logging |
| `settings.py` | Centralised constants for paths, limits, and defaults |

## Installation

### Prerequisites

- Python 3.10 or later
- pip

### Setup

```bash
# Clone the repository
git clone https://github.com/your-username/excel-automation-toolkit.git
cd excel-automation-toolkit

# Create and activate a virtual environment (Windows)
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Generate sample data (optional)
python sample_data/generate_samples.py

# Start the app
streamlit run app.py

# In another terminal, run tests
python -m pytest tests/ -v
```

Open http://localhost:8501 in your browser.

## Example usage

1. Click **Browse files** in Step 1 and select `sales_north.xlsx` and `sales_south.xlsx` from `sample_data/`
2. In Step 2, review the schema comparison pivot table and compatibility scores, then click **Merge Datasets**
3. In Step 3, remove duplicate rows and configure missing-value strategies per column
4. In Step 4, view summary statistics and run outlier detection on numeric columns
5. In Step 5, create a histogram, box plot, bar chart, or line chart
6. In Step 6, download the cleaned data as Excel or CSV
7. In Step 7, select report sections and click **Generate Report** to download a `.docx`

## Testing

87 tests across 5 test files, covering all modules:

- `test_analyzer.py` — column typing, statistics, outlier IQR detection
- `test_data_processor.py` — merge, compatibility, duplicate removal, fill-before-drop order, cleaning report
- `test_exporter.py` — filename safety, Excel/CSV round-trips, special characters, immutability
- `test_visualizer.py` — all chart types return figures, edge cases, immutability
- `test_report_generator.py` — valid docx output, section selection, cleaning data, outliers

```bash
python -m pytest tests/ -v
```

## Design decisions

- **Union merge as safe default** — preserving all columns avoids silent data loss. Users can switch to intersection if needed.
- **Jaccard compatibility scoring** — simple, interpretable 0–1 measure of column-set overlap. Threshold is user-adjustable.
- **User-controlled cleaning** — the tool recommends strategies but never auto-executes. Destructive actions require a deliberate button click.
- **Fill before drop** — fill strategies (mean, median, mode, zero, custom) run on the full dataset first; row drops happen last. This maximises the data available for fills.
- **Flagging, not deleting, outliers** — outliers are flagged with their Q1/Q3/IQR values. The user decides whether to investigate. No automatic removal.
- **In-memory exports** — Excel, CSV, and Word files are generated as byte streams. Nothing is written to disk — clean for cloud/container deployments.
- **Factual-only reports** — the Word report states objective facts (row counts, statistics, outlier values) without inventing business conclusions.

## Privacy

- All file processing happens in memory on your machine.
- No data is uploaded to external servers.
- Logs contain file names and row/column counts only — individual cell values are never written to logs.
- The app runs entirely on localhost.

## Limitations

- **Business meaning cannot be inferred** — the tool works with column names and data types, not business context. A column called "ID" is treated the same as "Revenue".
- **Compatibility scores are advisory** — a low Jaccard score does not mean merging is wrong; it means the column sets are different. The user decides.
- **Missing-value replacements are estimates** — mean, median, and mode fills are statistical approximations, not guarantees of the true value.
- **Very large files may need optimisation** — the app loads all data into memory. Files over 50 MB are blocked by default (configurable in `settings.py`).
- **No database or cloud-storage integrations** — the toolkit works with local file uploads only.
- **Single sheet per Excel file** — only the first sheet of each workbook is read.

## Future improvements

- Multi-sheet Excel file support
- Column-level filters and data-type casting
- Drag-and-drop column mapping for heterogeneous schemas
- Optional cloud storage connectors (S3, Azure Blob)
- Larger-file optimisation (chunked reading, lazy loading)
