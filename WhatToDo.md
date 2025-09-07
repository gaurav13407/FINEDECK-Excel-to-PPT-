1) Goal & responsibilities

The Excel reader should:

Load one or more sheets from an Excel file into clean pandas.DataFrame objects.

Normalize headings/types and handle common messy inputs.

Provide options for reading (specific sheets, header row, index column, parsers).

Be robust, testable, and easy to extend (CSV, Google Sheets later).

2) Public API (what functions should do)

Design clear function signatures (conceptual):

read_excel(path, sheets=None, header='infer', dtype=None, parse_dates=None, na_values=None, usecols=None, engine=None) -> dict[str, DataFrame] or DataFrame

sheets=None → read all sheets and return {sheet_name: df}.

sheets='Sheet1' or sheets=['Sheet1','Sheet2'] to read subset.

read_sheet(path, sheet_name, ...) -> DataFrame — convenience wrapper for single sheet.

Optionally: stream_excel(path, sheet_name, chunksize) for very large sheets.

3) What the reader must handle

Multiple sheets — return a mapping of sheet name → DataFrame.

Header detection — some files have metadata rows before the header. Allow header=n or provide heuristic detection to find the header row.

Blank rows / footer rows — trim trailing/leading blank rows.

Merged cells — convert merged header cells into repeated header names or flattened names.

Type inference & overrides — default to pandas inference but allow dtype overrides.

Dates & period columns — support parse_dates or automatic date detection for finance data.

Empty/NaN handling — na_values, drop optionally, or fill defaults.

Columns selection — usecols to limit columns (string Excel-style or list).

Whitespace & formatting — strip column names and string cells; normalize unicode.

Numeric formatting in Excel — currency/percent strings to numeric (strip symbols, convert to float).

Large files — chunked reading or memory-efficient approaches.

Password-protected files — either refuse or require password param.

Different file formats — support .xls, .xlsx, .xlsm, and optionally .csv.

4) Pre-processing & normalization steps

After reading raw sheet into a DataFrame, run a normalization pipeline:

Header normalization

Trim whitespace, lower/normalize case, replace spaces with _ (or keep original and provide slugify_cols option).

If duplicate column names, suffix with _1, _2 or coalesce intelligently.

Drop empty rows/cols

Remove rows/columns where all values are NaN or blank (configurable).

Type normalization

Convert numeric-looking strings ($1,234.56, 1,234%) to numbers using regex cleaning.

Parse dates with the chosen format(s). If ambiguous, let user pass formats.

Categorical detection

For low-cardinality string columns, convert to category dtype optionally (saves memory).

Index handling

If a column looks like a date or ID, expose option to set it as index.

Rename columns according to mapping

Allow a column_map to standardize financial column names (e.g., Revenue → revenue).

5) Error handling & logging

Return informative errors:

File not found, unsupported format, sheet not found, password required, memory error.

Graceful fallback:

If a sheet fails to parse, log the error and continue with other sheets (configurable).

Logging:

Info level: sheets found, rows/columns loaded.

Warning: dropped rows, type mismatches, forced conversions.

Debug: raw preview of first few rows, inferred dtypes.

6) Performance & memory considerations

For big sheets (>100k rows):

Use chunksize reading (if possible) and process in streaming fashion.

Convert large text columns to category if appropriate.

Avoid unnecessary copies; operate in-place when safe.

Cache results if same file read multiple times (checksum-based).

Allow user to specify columns to read (usecols) to reduce memory.

7) Validation & sanity checks

After read, run checks and return a small dataset_summary:

rows, cols, missing_pct_by_col, inferred_dtypes, sample_rows.

Allow a schema check function that validates required columns and types (useful for automated pipelines).

8) Extensibility

Abstract the IO layer so you can add:

CSV reader, Google Sheets reader, database connector.

Pluggable normalizers (middleware pattern):

read -> [normalizer1, normalizer2] -> return.

Expose hooks so PPT generator can request raw or normalized data.

9) Testing strategy

Unit tests and integration tests to cover:

Reading single-sheet and multi-sheet files.

Files with leading metadata rows (header=2 etc).

Merged headers.

Currency/percentage columns converted correctly.

Date parsing for multiple formats.

Very large files (simulate with generated DataFrames).

Error conditions (missing sheets, corrupted files).

Schema validation tests.

10) Example usage flow (conceptual)

User calls read_excel('finance.xlsx', sheets=None, header='infer', parse_dates=['Date']).

Reader loads workbook, lists sheets, reads target sheets.

Normalization pipeline runs: trim columns, parse numeric strings, detect dates, drop empty rows.

Return { 'Income Statement': df1, 'Balance Sheet': df2 } plus summary metadata.

If any sheet had issues, the function logs them and includes them in the returned summary.

11) Configuration & CLI

Provide a config.yaml or function params for typical settings:

default_header_row, drop_empty_rows, currency_symbols, date_formats.

CLI/Script usage: excel_to_ppt read --file finance.xlsx --sheets all --output-dir data/

12) Practical tips (implementation pointers)

Use openpyxl/xlrd/pyxlsb as backends depending on file type; fallback heuristics help.

For detecting header row programmatically: look for a row with many non-null strings and few numeric cells.

For cleaning numeric strings: remove thousands separators, strip currency symbols, then convert.

Keep the reader deterministic — no random sampling or unpredictable heuristics unless in auto_detect mode.

13) Deliverables you should get from this step

A clear function spec / docstring for read_excel.

A small set of example Excel files (clean, messy header, merged headers, currency formatting, big data) to use during implementation and tests.

A test plan listing tests you’ll write.

A normalization checklist that the reader will perform on every sheet.      