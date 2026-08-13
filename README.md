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

## Latest Repo Notes

- `csv_to_excel.py` still supports the original standalone CLI flow from `input-today` to `output-today`.
- `run_daily_pipeline.py` adds an optional one-command flow that runs the daily tracking, ongoing tracking, and iPhone tracking generators sequentially.
- The one-command pipeline is an additional option only. It does not replace the standalone scripts.
- `csv_to_excel.py` can now receive a specific yesterday/reference workbook internally, so the pipeline can pass `Daily Tracking Yesterday.xlsx` directly without changing the old CLI behavior.
- The daily tracking and iPhone generators now preserve header captions, header styling, and column widths from the reference/template workbook more closely.
- Date-like columns are scanned and formatted as Excel dates using `mm/dd/yyyy`; date-time columns keep time using `mm/dd/yyyy hh:mm`.
- The iPhone generator preserves the column model for both `ALL ORDER` and `ALL ORDER IPHONE`, including dynamic headers such as `TARGET ...` and `Phase ...`.
- `generate_ide_tracking.py` adds a standalone IDE Daily Tracking flow with historical lookup, previous-workbook OTC/MRC mapping, normalized Excel dates, and PivotTable refresh.

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

## Deploy On Windows Server

For native Windows Server deployment as a public FastAPI service on port `8000`, see [docs/windows-server-public-deployment.md](docs/windows-server-public-deployment.md).

## Run Standalone Scripts

Daily Tracking only:

```powershell
python .\csv_to_excel.py .\input-today -o .\output-today
```

Ongoing Tracking only:

```powershell
python .\csv_to_excel_on_going.py .\input-ongoing -o .\output-outgoing
```

iPhone Tracking only:

```powershell
python .\generate_iphone_tracking.py .\input-iphone -r .\vlookup-iphone -o .\output-iphone
```

The standalone scripts keep their original folders and commands. Use these when you only need one report.

## Run IDE Tracking

Required local folders:

```text
input-ide/
  IDE DASHBOARD (*.xlsx)
vlookup-ide/
  Daily Tracking IDE <previous date>.xlsx
output-ide/
```

Run with the default folders:

```powershell
python .\generate_ide_tracking.py
```

Or pass every location explicitly:

```powershell
python .\generate_ide_tracking.py .\input-ide `
  -r .\vlookup-ide `
  -o .\output-ide
```

Use `--report-date YYYY-MM-DD` when the report date cannot be inferred from the raw IDE dashboard filename. The generator keeps every raw row with a nonblank `QUOTE ID`, preserves duplicates, uses the previous workbook for historical fields and OTC/MRC, and writes zero when the previous workbook has no valid OTC/MRC value. Microsoft Excel and `pywin32` are required because the copied IDE PivotTables are refreshed through Excel COM.

## Run Full Daily Pipeline

Use this when you want one command to generate all three reports:

1. Daily Tracking
2. Daily Tracking On Progress
3. Daily Tracking iPhone

Required folder structure:

```text
input-pipeline/
  data-order/
    DataOrderSD-YYYYMMDD-*.csv
  log-update/
    LogUpdateStatusOrderSD-YYYYMMDD-*.csv

reference-pipeline/
  daily-yesterday/
    Daily Tracking <yesterday>.xlsx
  ongoing-yesterday/
    Daily Tracking On Progress <yesterday>.xlsx
  iphone-yesterday/
    Daily Tracking Iphone <yesterday>.xlsx

output-pipeline/
```

Run the pipeline:

```powershell
python .\run_daily_pipeline.py `
  --data-order .\input-pipeline\data-order `
  --daily-reference .\reference-pipeline\daily-yesterday `
  --log-update .\input-pipeline\log-update `
  --ongoing-reference .\reference-pipeline\ongoing-yesterday `
  --iphone-reference .\reference-pipeline\iphone-yesterday `
  -o .\output-pipeline
```

The pipeline runs the Excel COM-based steps one by one for reliability. This is intentional because running multiple Excel COM processes in parallel can cause workbook locks, failed saves, or unstable pivot refreshes.

Optional pipeline flags:

```text
--skip-ongoing           Generate only Daily Tracking and iPhone output.
--skip-iphone            Generate only Daily Tracking and ongoing output.
--skip-template-refresh  Skip Excel COM refresh for the Daily Tracking step.
--ongoing-with-pivot     Also refresh the ongoing ALL ORDER ON PROGRESS pivots.
--keep-temp              Keep temporary pipeline folders for debugging.
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
