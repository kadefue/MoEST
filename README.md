# MoEST — Data Processing Engine

This repository contains code (primarily a Jupyter notebook) to extract and clean tabular data from the BEST Excel workbooks (BEST {year}.xlsx) and produce per-sheet CSVs and combined CSV datasets for analysis.

The main implementation is in `MoEST_Data_Processing.ipynb`. The notebook was written to run in Google Colab, but it can be run locally with minor configuration changes.

---

## What this project does

- Loads BEST Excel files for years (2016–2025) from a configured directory (default in the notebook: Google Drive path `/content/drive/MyDrive/BEST`).
- For each Excel workbook, extracts two target sheets (left/right in a paired mapping per year).
- Applies cleaning logic that:
  - Normalizes merged header cells (duplicates text across merged rows/columns in the header region).
  - Fills down structural columns (e.g., region/council identifiers).
  - Removes summary rows like "Total" and "Grand".
  - Drops very sparse columns (columns > 60% empty/zero).
  - Removes duplicate columns that are identical to the region column.
- Saves:
  - One CSV per extracted sheet: `{year}_{sheet_name}.csv`
  - Two combined CSVs across years:
    - `Combined_Left_Sheets.csv`
    - `Combined_Right_Sheets.csv`

---

## Repository contents (key files)

- `MoEST_Data_Processing.ipynb` — the main notebook containing:
  - helper/cleaning functions (e.g., `is_valid_text`, `normalize_merged_cells`, `process_council_sheet`)
  - extraction functions (`extract_sheets_no_spaces`, `extract_sheets_cleaned`, `extract_and_combine`)
  - example runtime outputs printed in the notebook (processing logs)
- `LICENSE` — project license.

---

## Key functions and what they do

- `is_valid_text(value)`
  - Heuristic that returns True if a value looks like alphabetic/alphanumeric text. Used to detect meaningful label cells.

- `normalize_merged_cells(df, header_rows=15)`
  - For the top `header_rows` rows, forward-fills horizontally and vertically to expand merged header cells so subsequent cleaning can treat headers as repeated text values.
  - Important for Excel sheets that use merged header cells.

- `process_council_sheet(df)`
  - Core cleaning logic for council-style tables.
  - Identifies a `region_col` by scanning headers for typical region keywords (`'region'`, `'mkoa'`) and defaults to the first column if no label is found.
  - Fills structural columns downwards (unmerges vertically).
  - Removes rows from the first occurrence of "Grand" and removes rows containing "Total" or variations of "Sub-Total".
  - Attempts to find the council column (by keywords like `'council'`, `'halmashauri'`, `'district'`, etc.) and remove council-level totals.
  - Drops columns that are too sparse (> 60% missing/zero/empty).
  - Drops columns that are exact duplicates of the region column.

- `extract_sheets_no_spaces()` and `extract_sheets_cleaned()`
  - `extract_sheets_no_spaces()` is a simpler extraction that reads sheets with headers and writes CSVs.
  - `extract_sheets_cleaned()` reads sheets with `header=None`, runs `normalize_merged_cells` and `process_council_sheet`, then writes cleaned sheet CSVs (without writing header rows to CSV because headers are part of the data).

- `extract_and_combine()`
  - Iterates the year→sheet mapping, applies cleaning to each left/right sheet, prefixes each extracted DataFrame with a `Source_Year` column, collects left & right side DataFrames, and concatenates them into combined CSV files (`Combined_Left_Sheets.csv`, `Combined_Right_Sheets.csv`).

---

## Input expectations

- Excel files named `BEST {year}.xlsx` for years in the mapping (2016–2025).
- By default the notebook looks in a Google Drive path:
  - `/content/drive/MyDrive/BEST/BEST {year}.xlsx` (for some functions) or
  - `/content/drive/MyDrive/BEST` (directory used by other functions)
- Sheet names are explicitly mapped per year (examples: `3.26Lab`, `3.37Lab`, `T3.39LabG&NG`, etc.) — the mapping is defined in the code and should match the actual sheet tabs.

---

## Outputs

- Per-sheet CSVs: `{year}_{sheet_name}.csv` or `{year}_{sheet_name}.csv` (cleaned/no-header depending on function).
- Combined CSV files:
  - `Combined_Left_Sheets.csv`
  - `Combined_Right_Sheets.csv`

These combined CSVs include a `Source_Year` column that was inserted before concatenation.

---

## How to run

1. Requirements
   - Python 3.8+ (notebook runs in Colab)
   - pandas
   - openpyxl (for reading .xlsx)
   - (optional) Google Colab to mount Google Drive

   Install dependencies if running locally:
   - pip install pandas openpyxl

2. Using Google Colab
   - Open `MoEST_Data_Processing.ipynb` in Colab (the notebook contains a Colab badge).
   - Mount your Google Drive:
     - from google.colab import drive
     - drive.mount('/content/drive')
   - Ensure your BEST Excel files are in `/content/drive/MyDrive/BEST/` or update the `base_dir` variables in the notebook functions.
   - Run the cells (start with the helper functions, then the extraction cell you want: `extract_sheets_cleaned()` or `extract_and_combine()`).

3. Running locally
   - Copy the relevant cells into a Python script (or export the notebook as .py).
   - Update `base_dir` to a local path where `BEST {year}.xlsx` files are stored.
   - Run the script with `python script.py`.

---

## Troubleshooting & notes

- Missing sheet or missing file:
  - The notebook prints warnings like `[Warning] Sheet 'X' not found` or `Skipping {year}: File '...' not found`. Confirm your filenames and sheet names match the mapping.

- Pandas dtype warnings:
  - When assigning the normalized header subset back into the DataFrame (df.iloc[:limit] = subset), pandas may warn about incompatible dtypes. This is a warning from pandas about mixed types and may be harmless for extraction, but you can:
    - Cast cells explicitly (e.g., subset = subset.astype(object)), or
    - Suppress the warning during assignment, or
    - Normalize column types after cleaning.

- Header handling:
  - Cleaned outputs are saved with `header=False` in the notebook because header rows are part of the data after normalization. If you expect standard header rows, you may need to detect and set header rows programmatically.

- Custom sheet mappings:
  - If your BEST workbooks change sheet names or structure in future years, update the `sheet_mapping` dictionary in the notebook.

- Character encoding & special characters:
  - Some sheet names contain special characters (e.g., `&`, `G&NG`); filenames are currently created from raw sheet names and may include these characters. If you plan to move files across systems, sanitize filenames.

---

## Suggestions / Next steps

- Parameterize the notebook:
  - Move `base_dir` and `sheet_mapping` into a configuration block or an external config file so the extraction code can be reused easily.

- Extract reusable Python module:
  - Move helper functions into a standalone Python module (e.g., `moest_extractor.py`) and wrap the main flows with a simple CLI (argparse) for easier automation and testing.

- Add unit tests:
  - Add tests for `normalize_merged_cells` and `process_council_sheet` using small sample DataFrames that mimic merged headers and council tables.

- Improve logging:
  - Replace print statements with the `logging` module, and add log levels so large batch runs are easier to monitor.

- Robust dtype handling:
  - After header normalization, explicitly coerce column dtypes where numeric data is expected (using `pd.to_numeric(..., errors='coerce')`) to avoid warnings and facilitate analysis.

---

## License

This repository includes a `LICENSE` file. See that file for licensing terms.

