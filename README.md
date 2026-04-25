# CSV to Excel Cleaner

Small Python CLI and Docker image for parsing CSV files, cleaning common formatting problems, and exporting `.xlsx` files.

By default, the script reads CSV files from `.\input-today` and writes results to `.\output-today`.

## What It Cleans

- Detects CSV encoding and delimiter when not provided.
- Trims whitespace from headers and cell values.
- Formats headers as readable names by replacing underscores with spaces.
- Removes empty rows by default.
- Keeps empty columns by default so CSV headers are not lost.
- Converts repeated whitespace inside text to a single space.
- Optionally removes duplicate rows.
- Optionally infers numeric and date columns.

## Run Locally

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python csv_to_excel.py
```

Convert every CSV in a folder to separate Excel files:

```powershell
python csv_to_excel.py .\input-today -o .\output-today
```

Combine every CSV in a folder into one workbook:

```powershell
python csv_to_excel.py .\input-today -o .\output-today\combined.xlsx --combine
```

## Run With Docker

Build the image:

```powershell
docker build -t csv-to-excel .
```

Convert one file:

```powershell
docker run --rm -v "${PWD}:/data" csv-to-excel /data/input-today/report.csv -o /data/output-today/report.xlsx
```

Convert a folder:

```powershell
docker run --rm -v "${PWD}:/data" csv-to-excel /data/input-today -o /data/output-today
```

Combine a folder into one workbook:

```powershell
docker run --rm -v "${PWD}:/data" csv-to-excel /data/input-today -o /data/output-today/combined.xlsx --combine
```

## Useful Options

```text
--delimiter ";"          Use a specific delimiter instead of auto-detecting.
--encoding utf-8-sig     Use a specific CSV encoding.
--dedupe                 Remove duplicate rows.
--infer-types            Convert likely numeric and date columns.
--keep-empty             Keep fully empty rows and columns.
--drop-empty-columns     Drop columns where all values are empty.
--no-normalize-headers   Keep original column names.
```
