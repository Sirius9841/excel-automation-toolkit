# Excel Automation Toolkit

A Streamlit application for combining, cleaning, reviewing, and exporting spreadsheet data while keeping each approved change traceable.

## Live Demo

[Open the live demo](https://excel-automation-toolkit-4grkxben6isuewsyosq7mn.streamlit.app/)

The demo ships with sample workbooks — `sales_north.xlsx` and `sales_south.xlsx` (complete records, which is what the built-in integrity check needs) and `employees.xlsx` with a different schema. You don't need to prepare your own files.

## Overview

Spreadsheets come in from different offices, departments, or systems, and they rarely match up: columns differ, rows repeat, and a blank cell can mean "missing" or "this source never collected it". Excel Automation Toolkit combines them into one dataset you can review, and walks you through the duplicate and missing-value decisions instead of merging blindly.

## Screenshots

Screenshots will be added here once captured: upload, schema review, cleaning approvals, insights, downloads, and the Word report.

## Demo Workflow

1. Add two or more Excel or CSV files, or load the demo files.
2. Compare the columns, then combine. Keeping every column is the default.
3. Review duplicates and missing values. Actions are previews until applied.
4. Download the cleaned data or generate the Word report.

## Key Capabilities

- **Import and validation** — multiple `.xlsx`/`.csv` uploads (50 MB per file) with clear rejection of unsupported files; `order_id`-like fields stay text, every row is tagged with its `source_file`, and schemas are compared before combining.
- **Cleaning** — duplicate previews (identical rows by default, optionally on identity columns) and source-aware blank classification. Each affected column gets a conservative recommendation and a user-approved action: leave blank, recover from a validated relationship, or fill with median, mean, mode, or a custom value.
- **Review** — type-aware statistics for identifiers, numbers, categories, and dates; unusual values are flagged, not errors. Integrity checks run after cleaning; optional insights add distributions, comparisons, and source-file views.
- **Outputs** — formatted Excel workbook, UTF-8 CSV, audit CSV, and a customizable Word report (details below).

## Why the Cleaning Logic Is Safer

A basic merge creates blanks wherever one file lacks a column another has. Those blanks look identical to genuine ones, but they mean something different. This toolkit keeps each source schema: a cell in a column a source never contained is marked as unavailable from that source, and is never filled from another file or dropped.

For genuine blanks, recommendations are conservative. When complete records validate a relationship such as `Total = Quantity × Unit Price`, a missing quantity is reconstructed from `Total ÷ Unit Price` — a calculation from known inputs, not an estimate. The formula, inputs, record, and source go into the audit.

Cleaning actions apply only after the user approves them. Statistical replacements are identified as estimates and prefer same-source values; unusual values remain review flags, and every approved action has an audit record.

## Example Use Case

Two regional offices send weekly sales spreadsheets. One collects `customer_city`, the other `discount_code`, and both contain missing or repeated records. The app combines them without dropping either column, and you end up with files that are safe to review or share.

## Output Files

- **Excel workbook** — four sheets: `Cleaned Data`, `Cleaning Summary`, `Values to Review`, and `Cleaning Audit` (each change, with original state, method, source, and timestamp).
- **CSV** — the cleaned dataset in UTF-8 with a byte-order mark, so non-English text opens correctly.
- **Word report** — executive summary, dataset overview, cleaning summary, column statistics, values worth reviewing, methodology, and limitations. Charts and a data preview are optional.

The audit CSV is an example output of `scripts/generate_business_ready_samples.py`; in the app, the audit is inside the workbook and Word report.

## Architecture

```text
excel-automation-toolkit/
|-- app.py
|-- config/
|   `-- settings.py
|-- src/
|   |-- analyzer.py
|   |-- data_processor.py
|   |-- data_quality.py
|   |-- exporter.py
|   |-- file_handler.py
|   |-- insights.py
|   |-- integrity.py
|   |-- logger_setup.py
|   |-- report_generator.py
|   |-- ui_helpers.py
|   |-- utils.py
|   |-- visualizer.py
|   `-- workflow.py
|-- sample_data/
|-- scripts/
|-- tests/
|-- requirements.txt
`-- README.md
```

| Module | Responsibility |
| --- | --- |
| `app.py`, `workflow.py` | Streamlit UI, guided flow, session state |
| `file_handler.py`, `data_processor.py` | Uploads, parsing, schema comparison, combining |
| `data_quality.py`, `integrity.py` | Blank classification, audit records, relationship validation |
| `analyzer.py`, `insights.py`, `visualizer.py` | Statistics, flags, insights, charts |
| `exporter.py`, `report_generator.py` | Excel/CSV outputs and the Word report |
| `ui_helpers.py`, `utils.py`, `logger_setup.py` | UI helpers, shared utilities, logging |

## Technology

Python, Streamlit, pandas, openpyxl, Matplotlib, python-docx, pytest. Runtime dependencies are in `requirements.txt`.

## Running Locally

From PowerShell on Windows:

```powershell
git clone https://github.com/Sirius9841/excel-automation-toolkit.git
cd excel-automation-toolkit
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
streamlit run app.py
```

Run the tests:

```powershell
python -m pip install pytest
python -m pytest tests/ -v
```

Optional: regenerate the samples (`python sample_data\generate_samples.py`) or rebuild the example outputs in `output/` (`python scripts\generate_business_ready_samples.py`).

## Testing

181 tests, all passing — schema-aware merging, duplicate handling, missing-value decisions, source-aware replacements, deterministic recovery, integrity checks, exports, reports, and insights.

## Privacy

Uploaded files are parsed into memory; source workbooks are never overwritten, and outputs are generated in memory. Logs record operational details such as filenames and sizes — never cell values.

The code does not send uploads to an external analytics service. In the hosted demo, files are processed by the hosted Streamlit instance and reach that server.

## Current Limitations

- `.xlsx`/`.csv` only (first worksheet, 50 MB per file); requires at least two files, no key-based joins
- Column meaning is inferred, so domain fields may need review
- Only the built-in `Total = Quantity × Unit Price` relationship is validated
- Statistical fills are estimates; review flags are not errors
- No database or cloud-storage connectors; the audit CSV is not exposed in the UI

## Contact

- Upwork: UPWORK_PROFILE_URL (add your profile link)
- GitHub: [Sirius9841](https://github.com/Sirius9841)
