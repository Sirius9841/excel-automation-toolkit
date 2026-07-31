# Excel Automation Toolkit

A guided tool for combining, cleaning, reviewing, and exporting Excel and CSV files while keeping every approved change traceable.

[**Open the live demo**](https://excel-automation-toolkit-4grkxben6isuewsyosq7mn.streamlit.app/)

The demo includes sample files, so the complete workflow can be tested immediately.

<p align="center">
  <a href="https://excel-automation-toolkit-4grkxben6isuewsyosq7mn.streamlit.app/">
    <img
      src="screenshots/Starting_Page.png"
      alt="Excel Automation Toolkit"
      width="760"
    >
  </a>
</p>

## Overview

Combining spreadsheets becomes difficult when files use different columns, contain duplicate records, or represent missing data differently.

Excel Automation Toolkit combines those files into one reviewable dataset. It shows schema differences, previews cleaning actions, and lets the user approve each change before exporting the result with a complete audit trail.

## Key Features

- Compare spreadsheet schemas before combining files
- Preserve the source of every imported row
- Distinguish genuine missing values from fields a source never collected
- Preview duplicate and missing-value actions before applying them
- Recover values from validated relationships when possible
- Review statistics, integrity checks, and unusual values
- Export cleaned data, audit information, and a Word report

## Demo Workflow

1. Upload two or more Excel or CSV files, or load the included demo data.
2. Compare the detected schemas and choose how to combine them.
3. Review duplicate records and missing-value recommendations.
4. Apply only the changes you approve.
5. Export the cleaned dataset and supporting reports.

<details>
<summary><strong>View the complete workflow</strong></summary>

<br>

### 1. Review and combine

<p align="center">
  <img
    src="screenshots/Combine_files.png"
    alt="Review and combine uploaded files"
    width="720"
  >
</p>

### 2. Review duplicate records

<p align="center">
  <img
    src="screenshots/Clean_data_dub.png"
    alt="Review duplicate records"
    width="720"
  >
</p>

### 3. Handle missing values

<p align="center">
  <img
    src="screenshots/Clean_empty_rows.png"
    alt="Choose missing-value cleaning methods"
    width="720"
  >
</p>

### 4. Explore Data Insights

<p align="center">
  <img
    src="screenshots/Data_insights.png"
    alt="Explore Data Insights"
    width="720"
  >
</p>

### 5. Download the results

<p align="center">
  <img
    src="screenshots/Download_results.png"
    alt="Download cleaned data and generate a report"
    width="720"
  >
</p>

</details>

## Why This Is Different from a Basic Spreadsheet Merge

A basic merge creates blanks wherever one file lacks a column another has. Those blanks look identical to genuine ones, but they mean something different. This toolkit keeps each source schema: a cell in a column a source never contained is marked as unavailable from that source, and is never filled from another file or dropped.

For genuine blanks, recommendations are conservative. When complete records validate a relationship such as `Total = Quantity × Unit Price`, a missing quantity is reconstructed from `Total ÷ Unit Price` — a calculation from known inputs, not an estimate. The formula, inputs, record, and source go into the audit.

Cleaning actions apply only after the user approves them. Statistical replacements are identified as estimates and prefer same-source values; unusual values remain review flags, and every approved action has an audit record.

## Example Use Case

Two regional offices send weekly sales spreadsheets. One collects `customer_city`, the other `discount_code`, and both contain missing or repeated records. The app combines them without dropping either column, and you end up with files that are safe to review or share.

## Output Files

- **Excel workbook** — four sheets: `Cleaned Data`, `Cleaning Summary`, `Values to Review`, and `Cleaning Audit` (each change, with original state, method, source, and timestamp).
- **CSV** — the cleaned dataset in UTF-8 with a byte-order mark, so non-English text opens correctly.
- **Word report** — executive summary, dataset overview, cleaning summary, column statistics, values worth reviewing, methodology, and limitations. Charts and a data preview can be included optionally.


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

Uploaded files are parsed into memory; source workbooks are never overwritten, and outputs are generated in memory. Logs record operational details such as filenames and sizes, never cell values.

The code does not send uploads to an external analytics service. In the hosted demo, files are processed by the hosted Streamlit instance and reach that server.

## Current Limitations

- Supports `.xlsx` and `.csv` files up to 50 MB and reads the first Excel worksheet
- The workflow requires at least two files and combines rows rather than performing key-based joins
- Column meaning is inferred, so domain fields may need review
- Only the built-in `Total = Quantity × Unit Price` relationship is validated
- Statistical fills are estimates; review flags are not errors
- No database or cloud-storage connectors; the audit CSV is not exposed in the UI

## Contact

- Email: [m.sigtermans98@gmail.com](mailto:m.sigtermans98@gmail.com)
- GitHub: [Sirius9841](https://github.com/Sirius9841)