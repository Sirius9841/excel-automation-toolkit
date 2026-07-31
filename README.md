# Excel Automation Toolkit

A Streamlit application for combining, cleaning, reviewing, and exporting spreadsheet data while keeping each approved change traceable.

## Live demo

[Open the live demo](https://excel-automation-toolkit-4grkxben6isuewsyosq7mn.streamlit.app/)

The demo includes sample sales files, so you can test the complete workflow without preparing spreadsheets first.

## Project overview

Teams often receive spreadsheets from different offices, departments, or systems. The files may have different columns, repeated records, genuine blank values, and fields that one source never collected.

Excel Automation Toolkit combines those files into one reviewable dataset, shows the schema differences, and guides the user through duplicate and missing-value decisions. It is useful for operations teams, analysts, small businesses, and developers who need a repeatable process without hiding how the data changed.

## Screenshots

Screenshots will be added here once captured. Planned captures and captions:

- Upload files: add `.xlsx` or `.csv` files, or load the included demo files. Planned path: `screenshots/upload.png`
- Review and combine: compare columns and choose whether to keep every field or only shared fields. Planned path: `screenshots/review.png`
- Clean the data: preview repeated records and approve blank-cell actions. Planned path: `screenshots/cleaning.png`
- Review optional insights: inspect one column, create a compatible chart, or compare source files. Planned path: `screenshots/insights.png`
- Download results: prepare the cleaned Excel or CSV file and generate the Word report. Planned path: `screenshots/download.png`
- Review the report: show the generated data quality summary and review sections. Planned path: `screenshots/report.png`

## Demo workflow

1. Add two or more Excel or CSV files. The demo button loads the two included sales workbooks.
2. Review the column differences and combine the files. Keeping every column is the recommended default.
3. Review repeated records and missing values. Cleaning actions are previews until you apply them.
4. Download the cleaned data as Excel or CSV, or generate a Word data quality report.

Data Insights is optional. You can open it from the cleaning or download screen without making it part of the required workflow.

## Key capabilities

### File handling

- Accepts multiple `.xlsx` and `.csv` uploads, with a 50 MB limit per file
- Reads the first worksheet from each Excel workbook
- Reads UTF-8 CSV files and falls back to Latin-1 when needed
- Preserves identifier-like fields such as `order_id` as text during import
- Compares source schemas before combining and lets the user keep all columns or only shared columns
- Adds `source_file` and original source-row context for traceability
- Rejects unsupported, oversized, or unreadable files with a clear message

### Data cleaning

- Treats completely identical rows as duplicates by default, with an option to match on selected identity columns
- Previews repeated records before the user approves their removal
- Separates genuine row-level blanks from fields that were unavailable in a source file
- Offers conservative recommendations and user-approved actions for each affected column
- Supports leaving values blank, arithmetic recovery, median, average, mode, custom-value replacement, and row removal where applicable
- Uses same-source values first for statistical replacements and records any fallback
- Records approved changes, retained blanks, removed rows, and formulas in a cleaning audit

### Data review

- Produces type-aware statistics for identifiers, numeric fields, categories, and dates
- Flags unusual numeric values for review without treating them as automatic errors
- Checks validated arithmetic relationships after cleaning
- Shows record and source context for statistical and integrity findings
- Provides optional distributions, range views, category comparisons, time views, and source-file comparisons

### Outputs

- Formatted Excel workbook with cleaned data, summaries, review findings, and the audit
- UTF-8 CSV export of the cleaned dataset
- Structured cleaning audit CSV through the export helper and sample-output script
- Customizable Word data quality report

## Why the cleaning logic is safer

A basic merge creates blank cells whenever one file has a column that another file does not. Those cells look the same as a blank inside a column that the source actually supplied, but they mean different things.

This application keeps each source schema and distinguishes between the two cases. If a source never contained `Customer City`, the tool marks those cells as unavailable from that source. It does not fill them with a city taken from another file or remove those rows by default.

For genuine blanks, the app starts with conservative recommendations. If complete records validate a relationship such as `Total = Quantity × Unit Price`, it can reconstruct a missing quantity from `Total ÷ Unit Price`. That is a calculation from known inputs, not an estimate. The formula, inputs, record, and source are written to the audit.

The user reviews and applies cleaning actions before the data changes. Statistical replacements are identified as estimates, unusual values remain review flags, and every approved action has an audit record.

## Output files

### Excel workbook

The Excel export contains four sheets:

- `Cleaned Data` contains the approved dataset as a formatted, filterable table with a frozen header row. Dates, numbers, and identifier fields receive appropriate Excel formats.
- `Cleaning Summary` contains source files, row counts, cleaning totals, blank statuses, warnings, and integrity results.
- `Values to Review` contains statistical review flags and failed integrity checks with record and source context.
- `Cleaning Audit` contains approved changes and documented structural blanks, including the original state, resulting value or state, method, source, record reference, and timestamp.

### CSV

The cleaned CSV contains the same approved dataset in a simple interoperable format. It is encoded as UTF-8 with a byte order mark so non-English text opens correctly in common spreadsheet software.

### Cleaning audit CSV

`src/exporter.py` includes a separate audit CSV export. The verification script at `scripts/generate_business_ready_samples.py` uses it to create `output/business_ready_cleaning_audit.csv` for the included sales example.

The current Streamlit download screen does not expose this as a separate button. In the application workflow, the audit is included in the Excel workbook and summarized in the Word report.

### Word report

The standard `.docx` report contains an executive summary followed by:

- Dataset Overview
- Data Quality and Cleaning Summary
- Column Statistics
- Values Worth Reviewing
- Methodology and Limitations

Charts and a short data preview are optional report sections. The report describes data quality, cleaning decisions, statistics, and review flags. It does not turn those findings into business conclusions.

## Example use case

A company receives weekly sales spreadsheets from two regional offices. One file includes `customer_city`, the other includes `discount_code`, and both contain a few missing or repeated records.

The application combines the files without deleting the unmatched columns, distinguishes unavailable fields from genuine blanks, records approved cleaning decisions, and produces files that can be reviewed or shared.

## Sample data

The repository includes three generated workbooks:

- `sample_data/sales_north.xlsx`: 60 sales rows with `customer_city`, repeated rows, missing values, and an unusual unit price
- `sample_data/sales_south.xlsx`: 40 sales rows with `discount_code`, missing values, and a schema that differs from the north file
- `sample_data/employees.xlsx`: 25 employee rows used as a separate-schema sample

The in-app demo loads the two sales workbooks. Their complete records support the built-in `Total = Quantity × Unit Price` integrity check. The employee workbook is not loaded by the demo button.

## Project architecture

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
|   |-- generate_samples.py
|   `-- *.xlsx
|-- scripts/
|   `-- generate_business_ready_samples.py
|-- tests/
|   `-- test_*.py
|-- requirements.txt
`-- README.md
```

| Module | Responsibility |
| --- | --- |
| `app.py` | Streamlit screens, controls, feedback, and the guided user flow |
| `file_handler.py` | Upload size checks and Excel or CSV parsing |
| `data_processor.py` | Schema comparison, row-wise combining, duplicate handling, and cleaning summaries |
| `data_quality.py` | Source-aware blank classification, cleaning decisions, and audit records |
| `integrity.py` | Relationship validation, deterministic recovery, and integrity findings |
| `logger_setup.py` | Console and file logging setup shared by all modules |
| `analyzer.py` | Column typing, summary statistics, and unusual-value detection |
| `insights.py` | Optional column metrics, chart compatibility, time views, and source comparisons |
| `visualizer.py` | Matplotlib chart creation |
| `exporter.py` | Excel, cleaned CSV, and audit CSV generation |
| `report_generator.py` | Word data quality report generation |
| `workflow.py` | Screen navigation, session-state lifecycle, and export context |
| `ui_helpers.py` | Small UI-facing helpers for duplicate and insight choices |
| `utils.py` | Shared helpers such as file-extension checks and byte formatting |

## Technology

- Python
- Streamlit
- pandas
- openpyxl
- Matplotlib
- python-docx
- pytest

The application dependencies are listed in `requirements.txt`. Pytest is used by the repository but is not currently included in that file, so it is installed separately in the test instructions below.

## Running locally

From PowerShell on Windows:

```powershell
git clone https://github.com/Sirius9841/excel-automation-toolkit.git
cd excel-automation-toolkit
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
streamlit run app.py
```

The sample workbooks are already committed. Regenerate them only if you want to restore their original generated contents:

```powershell
python sample_data\generate_samples.py
```

Install the test runner and run the full suite:

```powershell
python -m pip install pytest
python -m pytest tests/ -v
```

To generate and verify the example Excel, cleaned CSV, audit CSV, and Word outputs:

```powershell
python scripts\generate_business_ready_samples.py
```

That command writes generated files to `output/`.

## Testing

The current suite contains 181 tests, and all 181 pass.

The tests cover schema-aware merging, duplicate handling, missing-value decision order, source-aware replacements, deterministic recovery, integrity checks, immutability, statistics, optional insights, charts, Excel and CSV exports, Word reports, sample data, and workflow state transitions.

## Privacy and data handling

Uploaded files are parsed into memory by the running application. The source workbooks and CSV files are not overwritten, and the downloadable outputs are generated in memory.

Normal application logs record operational details such as filenames, file sizes, row and column counts, selected operations, warnings, and exceptions. The logging calls do not write spreadsheet cell values.

The repository does not contain code that intentionally sends uploaded files to an external analytics service. In the hosted demo, uploads are processed by the hosted Streamlit application instance, so the files do reach that server for processing.

## Current limitations

- The normal workflow requires at least two files and combines them by appending rows. It does not perform key-based joins.
- Only `.xlsx` and `.csv` uploads are accepted.
- Only the first worksheet in an Excel workbook is read.
- Each file is limited to 50 MB. Practical capacity also depends on the memory available to the running Streamlit instance.
- Column meaning is inferred from names, data types, and uniqueness patterns, so domain-specific fields may need review.
- Automatic relationship discovery currently covers the built-in `Total = Quantity × Unit Price` pattern. The UI does not define additional rules.
- Statistical replacements are estimates, and statistical review flags are not automatic errors.
- Imports come from local file uploads. There are no database or cloud-storage connectors in the current application.
- The separate audit CSV helper is not exposed on the Streamlit download screen.

## Possible future improvements

- Read and combine selected sheets from multi-sheet workbooks
- Add key-based joins, column mapping, and saved cleaning templates
- Let users configure additional relationship rules
- Add database and cloud-storage imports

## What I learned

Building this project taught me to treat spreadsheet cleaning as a review workflow rather than a chain of automatic edits. I spent most of the effort on preserving provenance, separating UI state from processing code, testing exports and edge cases, and making the result a complete tool instead of stopping at a data-cleaning script.

## Contact

- Upwork: UPWORK_PROFILE_URL (add your profile link)
- GitHub: [Sirius9841](https://github.com/Sirius9841)
