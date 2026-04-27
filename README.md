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
docker run --rm `
  -v "${PWD}:/data" `
  -v "${PWD}\vlookup-yesterday:/app/vlookup-yesterday:ro" `
  csv-to-excel python /app/csv_to_excel.py /data/input-today/report.csv -o /data/output-today/report.xlsx --skip-template-refresh
```

Convert a folder:

```powershell
docker run --rm `
  -v "${PWD}:/data" `
  -v "${PWD}\vlookup-yesterday:/app/vlookup-yesterday:ro" `
  csv-to-excel python /app/csv_to_excel.py /data/input-today -o /data/output-today --skip-template-refresh
```

Combine a folder into one workbook:

```powershell
docker run --rm `
  -v "${PWD}:/data" `
  -v "${PWD}\vlookup-yesterday:/app/vlookup-yesterday:ro" `
  csv-to-excel python /app/csv_to_excel.py /data/input-today -o /data/output-today/combined.xlsx --combine
```

## Run API With Docker Compose

Start the FastAPI service:

```powershell
docker compose up --build
```

The API listens on `http://localhost:8000`.

Health check:

```powershell
curl http://localhost:8000/health
```

Convert the mounted `.\input-today` folder and write to `.\output-today`:

```powershell
curl -X POST http://localhost:8000/convert/path `
  -H "Content-Type: application/json" `
  -d "{\"input_path\":\"input-today\",\"output_path\":\"output-today\",\"refresh_template\":false}"
```

Upload one raw CSV plus yesterday's cleaned Excel workbook and receive a JSON success response:

```powershell
curl -X POST http://localhost:8000/convert/upload `
  -F "raw_data=@.\input-today\report.csv;type=text/csv" `
  -F "yesterday_cleaned_data=@.\vlookup-yesterday\Daily Tracking 23 April 2026.xlsx;type=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" `
  -F "refresh_template=false"
```

The response includes a `download_url` field with a full link to download the generated Excel workbook.

The Docker API defaults to portable workbook generation (`refresh_template=false`) because Linux containers cannot automate Microsoft Excel through Windows COM. On Windows, the CLI can still refresh the template by default. Use `--skip-template-refresh` for portable CLI output:

```powershell
python csv_to_excel.py .\input-today -o .\output-today --skip-template-refresh
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
--skip-template-refresh  Skip Windows Excel COM template refresh.
```
