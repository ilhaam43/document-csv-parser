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
- `weekly/` generates the Weekly Master plus selected Mobile/M2M workbooks and editable PPTs through one shared Catalyst review flow at `/weekly`.

## Run Locally

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python csv_to_excel.py
```

Weekly development and verification commands are documented in
[docs/weekly-report-development.md](docs/weekly-report-development.md).

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

Use `--report-date YYYY-MM-DD` when the report date cannot be inferred from the raw IDE dashboard filename. The generator keeps every raw row with a nonblank `QUOTE ID`, preserves duplicates, and uses the previous workbook for historical fields and OTC/MRC. Microsoft Excel and `pywin32` are required because the copied IDE PivotTables are refreshed through Excel COM.

IDE output columns follow the raw dashboard alongside the generator's calculated fields. Validation checks required input headers, output header order, the Excel table range, and refreshed pivot cache fields rather than requiring a fixed 61-column output.
`MONTH COMPLETE` is always calculated from `SO COMPLETION DATE` and placed
immediately after it (AG in the standard layout). Existing full-month templates
retain `TEXT(...,"MMMM")`, including January for empty dates. A new-period
rebuild uses `MMM-YY`, blank for missing completion dates; subsequent daily runs
retain that format. In this format, an existing dashboard month column is kept
separately as `MONTH COMPLETE2`, never used as the calculated primary month.
Unlabelled scratch columns after the last declared source header are excluded.
Report 4 still requires only the dashboard, previous daily workbook, and collabs
CSV. No date-correction form or extra review upload is required.
Within an existing business period, the delivery pivots retain the previous
workbook's effective month and status selections, including hidden checkboxes.
Legacy templates with no completion-month page field retain the old generator's
monthly scope when the report month is available, otherwise all months.
On a new business period, On Time Delivery uses all completion months and all
delivery statuses with populated OTD results. RFS aging includes open statuses,
including Waiting AOS, and excludes SO Complete and Cancel. This avoids carrying
the old period's restrictive filters into a new period. The first calendar day
still closes the prior period, like the existing target-bucket rule. The previous
report date comes from its filename; if unavailable, cumulative defaults apply.
Other pivot filters are preserved. No source rows, dates, or month members are
added to make a completion month available.
Every connected IDE pivot preserves cell formatting and disables automatic
reformatting so a shared-cache refresh does not shrink fitted column widths.

The selected RFS text policy is
`legacy-excel`, following the owner's decision to retain the manual report's
historical conversions. Ambiguous RFS text is month-first (`11/05/2026` becomes
5 November), with day-first fallback for unambiguous values such as
`26/02/2026`. This is an explicit compatibility policy, not a claim that mixed
regional conversion is a canonical date standard. Strict day-first parsing is
tested separately but is not the active generator default. ISO/year-first and
typed Excel dates are not regionally reparsed. Output RFS columns display dates as
`dd/mm/yyyy` to make the selected meaning visible, with time retained if present.
The generation metadata records `rfs_text_date_policy`. The supplied FAB column
also contains mixed legacy text; its existing handling is unchanged.

The manual samples use two different preparation workflows, selected from the
previous and current business periods rather than from specific months or quote
IDs. These are explicit legacy-compatibility rules, not a universal date policy:

- **Continuation within a period:** preserve reference identifiers and the
  existing historical status header. Nonblank duplicate ACTUAL RFS dates follow
  the first quote match. Unreadable NEW RFS dates prefer the same previous quote
  occurrence before agreeing current peers and `source`. Year-only date changes
  retain the matching reference year. Existing financial snapshot values,
  including missing values resolved to zero, are retained; collabs can supply
  newly added quotes. The daily pre-installation placeholder advances only when
  recognized as the previous daily placeholder; genuine older starts remain.
- **New-period rebuild:** retain each source row's identifier and typed RFS
  date. Broken NEW RFS lookups prefer agreeing current peers and `source` before
  historical recovery. Legacy text can retain a matching reference year.
  Historical NEW RFS originating as text may use SO completion year under the
  approved rule, but never when that creates a date after the report cutoff;
  missing evidence, native dates and invalid leap-day conversions are unchanged.
  OTC/MRC follow exact reference keys, missing values becoming zero without a
  newer CSV overriding that snapshot. New pre-installation starts require an
  observed Pre-Installation or Installation status. Ancillary date text stays
  text; numeric Excel date serials become typed dates.

Historical status values come from the existing historical baseline, not the
current status column. A period rebuild relabels that baseline with the prior
report date. Intentional historical `#N/A` values remain native errors.
When no reference financial snapshot exists, collabs remains the fallback.
Blank current RFS dates stay blank. No duplicate rows are removed.

An out-of-range `QUOTE DATE` year is corrected only when `source.quote_created`
has one agreeing calendar date for the quote with the same month and day. The
original time is retained; the current/report year is never guessed. Other
invalid years without this evidence are retained and reported, not silently
rewritten. Metadata records `reference_workflow`, `financial_source_policy`,
`historical_status_source`, `month_complete_format`, `source_date_corrections`
and `planned_rfs_year_rollovers`.
Editing and committing text in Excel can trigger locale-dependent conversion;
merely selecting a cell does not change its stored date. Such conversion does
not prove the intended business date. A date already converted into an Excel
numeric date cannot be diagnosed as a previous month/day swap from its display
alone. No per-quote exception or date-correction form is used.
Intentional `#N/A` sentinels use native `=NA()` formulas, avoiding a clipboard
copy/paste step while retaining real Excel errors for pivot filtering.

### Manual Parity Scope

Source snapshots matter: the live November dashboard was changed during date
inspection, while the original upload snapshot still contains date text at that
cell. The original November dashboard and manual reference also disagree on one
ambiguous RFS text interpretation. No single global conversion policy reproduces
that exception automatically.

For reference validation, an input copy was prepared using the owner-approved
manual date to convert only that ambiguous text cell into a typed Excel date.
All other input cells and workbook parts were preserved. This is documented
input normalization, not a hidden generator correction. The production generator
does not read the final manual workbook, infer a date from desired pivot totals,
or contain a quote-specific exception. The reference-validation input copy must
be used to reproduce that run; uploading the unnormalized or subsequently edited
dashboard is not guaranteed to reproduce the same historical results.

The web workflow still needs only three files. Full `ALL ORDER` equality is
separate from numeric PivotTable parity and must not be implied by a successful
generation or pivot-only comparison. Original and manual workbooks were not
modified by the validation.

The November three-file API run with that normalized dashboard completed with
2,048 rows and 63 columns. The strict `ALL ORDER` audit found zero cell differences,
including date versus text representation and column order. Independent native
Excel comparison also found zero numeric differences in all ten pivots.
This is explicitly **normalized-input parity**, not parity from an unchanged
or subsequently edited dashboard. The one ambiguous input cell remains a
source-data limitation, not an automatic quote-specific correction.

The October regression uses the original dashboard and has 2,004 rows and
62 columns. Its one approved detail deviation is `QUOTE DATE` for quote
`2-833355577463`: source-confirmed 2026 replaces the manual's remaining 3036.
That correction is deliberately reported as a difference, not suppressed to
claim literal equality. Both reports use the same source-evidenced correction.
All ten native October pivot value sets match the manual as well. Both API
jobs successfully produced downloadable workbooks, and Excel checks found no
hidden pivot body rows or numeric `####` truncation. Validation left source and
manual files unchanged, and all artifacts remain outside the repository.

Run a read-only detail audit separately from the native pivot check:

```powershell
python .\ide_reference_audit.py --manual .\manual.xlsx `
  --generated .\generated.xlsx --output .\comparison.json
```

The command checks the entire declared `ALL ORDER` schema, row order and cached
values. It exits nonzero on any detail difference, including approved corrections;
it never changes either workbook. `exact_value_matches` is distinct from
`semantic_value_matches` (for example, text versus typed dates). Neither result
proves formula-text, formatting or native PivotTable equality.

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

Daily Tracking copies completion-column styles by locating the actual template
headers, so inserted rows do not break format recovery. Percentage formatting
treats an empty table or Grand Total-only table separately from detail rows.
After refresh, completion blocks recover a missing native traffic-light rule
with the existing report-date thresholds. Unrelated conditional formats and
source data are retained. Completion formulas use the refreshed phase pivot's
body and join denominator rows by their full pivot category path, including
subtotal and Grand Total depth. They do not assume that adjacent pivots have the
same row numbers or order. Empty quote-count totals resolve to zero percent;
missing category matches remain errors. Old standalone percentage blocks are
cleared before refresh so moved or shrunken tables do not retain stale formulas.
The final presentation pass restores template column widths and reveals pivot
result rows without exposing intentionally hidden page-filter rows.
Failed Excel rule updates or completion formula errors still fail the job rather
than publishing a workbook with unchecked thresholds or broken percentages.
Successful generation and presentation checks do not establish data parity with
a manual reference; row values, filters and pivot cache coverage need a separate audit.

The current export's recognized Phase takes precedence over historical Phase;
Status is only a fallback when that source Phase is unavailable. Process Adjustment
is rebuilt from the original Process when the template's target period changes:
New Registration becomes New Reg, and other or blank source values become Non New
Reg. Within the same target period, existing classifications are retained and only
missing classifications are inferred. Historical display fallbacks do not change
the original export used for period rebuilding.

iPhone completion pivots are identified by their New/Non New and Phase fields,
not arbitrary names assigned by Excel. Their caches are rebound to local source
tables and checked against the actual table row counts; external or stale caches
cannot silently supply report values. FAB-year filters reset to All for a new
reporting period, including orders without a valid FAB year; daily continuation
preserves the existing selection. Consequently, an old manual pivot still linked
to an external historical workbook can differ intentionally: the generator uses
the current local data, not the historical external snapshot.

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
