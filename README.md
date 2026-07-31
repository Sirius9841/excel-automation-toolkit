# Excel Automation Toolkit

A Streamlit application for combining, cleaning, reviewing, and exporting Excel and CSV data while keeping every approved change fully traceable.

## Live Demo

[Open the live demo](https://excel-automation-toolkit-4grkxben6isuewsyosq7mn.streamlit.app/)

The demo includes sample workbooks, so you can try the complete workflow without uploading your own files.



## Overview

Combining spreadsheets sounds simple until the files don't match.

Different departments often use different columns, duplicate records appear, and blank cells can mean either *missing data* or *this source never collected that field*. Treating those situations the same can lead to incorrect cleaning decisions.

Excel Automation Toolkit guides the user through the entire process instead of making automatic assumptions. You review duplicates, missing values, and cleaning recommendations before any changes are applied, then export the results together with a complete audit trail.



## Screenshots

| Upload files | Review and combine |
| --- | --- |
| ![Upload files](Starting_Page) | ![Review and combine](Combine_files) |

| Review duplicates | Handle missing values |
| --- | --- |
| ![Duplicate review](Clean_data_dub) | ![Missing-value options](Clean_empty_rows) |

| Explore insights | Export results |
| --- | --- |
| ![Data Insights](Data_insights) | ![Download results](Download_results) |



## Demo Workflow

1. Upload two or more Excel or CSV files, or load the included demo data.
2. Compare the detected schemas before combining the datasets.
3. Review duplicate records and missing-value recommendations.
4. Apply the changes you approve.
5. Export the cleaned dataset and supporting reports.



## Key Features

### Import

- Supports multiple Excel and CSV files
- Compares schemas before combining data
- Preserves identifier columns as text
- Records the original source file for every row

### Cleaning

The application does not apply cleaning changes without the approval of the user.

Instead, it presents every cleaning decision on review.

Some of the supported actions include:

- duplicate removal
- source-aware missing-value handling
- deterministic value recovery from validated relationships
- statistical replacements (mean, median, mode, or custom values)

Every approved change is written to the audit.

### Review

Before exporting, the application provides:

- statistics on the data
- integrity checks
- unusual-value detection
- optional visualizations and source comparisons

### Export

Generate:

- formatted Excel workbook
- cleaned CSV
- Word data quality report
- cleaning audit included in the Excel workbook

##  Why This Is Different from a Basic Spreadsheet Merge

A basic merge creates blanks wherever one file lacks a column another has. Those blanks look identical to genuine ones, but they mean something different. This toolkit keeps each source schema: a cell in a column a source never contained is marked as unavailable from that source, and is never filled from another file or dropped.

For genuine blanks, recommendations are conservative. When complete records validate a relationship such as `Total = Quantity × Unit Price`, a missing quantity is reconstructed from `Total ÷ Unit Price` — a calculation from known inputs, not an estimate. The formula, inputs, record, and source go into the audit.

Cleaning actions apply only after the user approves them. Statistical replacements are identified as estimates and prefer same-source values; unusual values remain review flags, and every approved action has an audit record.

## Example Use Case

Two regional offices send weekly sales spreadsheets. One collects `customer_city`, the other `discount_code`, and both contain missing or repeated records. The app combines them without dropping either column, and you end up with files that are safe to review or share.

## Output Files

- **Excel workbook** — four sheets: `Cleaned Data`, `Cleaning Summary`, `Values to Review`, and `Cleaning Audit` (each change, with original state, method, source, and timestamp).
- **CSV** — the cleaned dataset in UTF-8 with a byte-order mark, so non-English text opens correctly.
- **Word report** — executive summary, dataset overview, cleaning summary, column statistics, values worth reviewing, methodology, and limitations. Charts and a data preview are optional/selectable.


## Architecture



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

- Supports `.xlsx` and `.csv` files up to 50 MB and reads the first Excel worksheet
- The workflow requires at least two files and combines rows rather than performing key-based joins
- Column meaning is inferred, so domain fields may need review
- Only the built-in `Total = Quantity × Unit Price` relationship is validated
- Statistical fills are estimates; review flags are not errors
- No database or cloud-storage connectors; the audit CSV is not exposed in the UI

## Contact

- Upwork: UPWORK_PROFILE_URL (add your profile link)
- GitHub: [Sirius9841](https://github.com/Sirius9841)
