# IDE List Format

## Inputs and Isolation

Report 4 detects the new format by the `List` worksheet and required headers,
not by a report date or quote allowlist. `IDE-list-YYYY-MM-DD.xlsx` supplies the
report date unless the CLI explicitly supplies `--report-date`.

The web form accepts two files: the raw List export and the previous new-format
Daily Tracking IDE workbook containing `ALL ORDER`, `PIVOT`, and `Table1`.
Collabs/GCP is not read, even if supplied. The legacy Dashboard branch continues
to require its CSV and retains its existing calculations and pivot policies.

`ide_list_format.py` owns the new adapter/writer; `generate_ide_tracking.py`
dispatches to it. The upload endpoint and isolated IDE worker use the same entry
point as the CLI. Weekly and Reports 1-3 are not changed.

## Row Rules

| Output | Rule |
| --- | --- |
| Rows and Key | Keep raw order and every row. Key must be present and unique; duplicate Quote IDs remain separate rows. |
| Revenue | Copy raw Revenue, including missing values and actual zero. No historical/CSV fallback. Financial pivots sum Revenue. |
| Revenue (Juta) | Preserve the export's value; do not use its rounded amount as financial pivot input. |
| Status | Current raw status, prefixed as 1.Quotation, 2.Pra-Delivery, 3.Delivery, 4.Waiting AOS, 5.SO Complete, 6.Cancel. Unknown statuses fail explicitly. |
| Yesterday Status | First matching Quote ID's current Status in the previous workbook. Missing quote produces intentional Excel `#N/A` only here. |
| ORDER TYPE | Keep the first matching quote's historical ORDER TYPE. For unmatched quotes, a word starting RENEW in Quote Name gives Renewal; otherwise New Registration. This includes the observed RENEWEL spelling. |
| PROCESS | New Registration gives New Reg; other preserved order types give Non New Reg. |
| Monthly target | Compare Target RFS month/year to the calendar report month. Missing Target RFS gives Target Not Yet Inputted. Missing FAB or Actual RFS does not erase a valid target. |
| Aging Of RFS | Report date minus Target RFS date. Missing target remains truly empty. The formula uses the report date, not volatile NOW, so reruns are reproducible. |
| STATUS ORDER | Aging less than -5: On Track; -5 through -1: Potential Delay; zero and above: Delay. Missing aging remains empty. |
| ON TIME DELIVERY | Actual RFS greater than Target RFS gives DELAY, otherwise ONTIME. Both dates are required. Raw OTD remains a separate source column. |
| YEAR FAB UPLOAD | TEXT(FAB Upload Date,"yyyy"), including the reference's 1900 placeholder for a missing date. |
| MONTH COMPLETE | TEXT(SO Completion Date,"MMM-YY"), including the reference's Jan-00 placeholder for a missing date. Raw Month Complete is kept separately as Month Complete2. |
| Other fields | Preserve source values and timestamps; additional declared source columns are appended. |

The standard sample layout contains 67 columns. Required fields are validated;
it is not a fixed column-count gate. Text cells are explicitly written as text
to prevent Excel converting values such as `Nov 2026` into dates.

New-format dates must be typed Excel dates. Ambiguous text dates fail with the
column and row identified rather than silently choosing a locale. Unusual
typed source years are preserved, not guessed; source warnings are included in
generation metadata. Missing completion dates do not change a valid raw status.

## Pivot and Presentation Rules

The copied reference owns layout, names, row hierarchy, selected service/status
populations, and year filters. Monthly target selections are mapped by bucket
meaning to the new month. No fixed quote list or expected output is loaded by
production code.

The native Excel writer binds every pivot to the current Table1 schema by name,
sharing one live cache. This avoids attaching old field indices to unrelated
columns after the raw/export order changes. Formula blanks are cleared where
required so the PivotCache retains a genuine `(blank)` category rather than a
distinct empty-string category.

Counts, sums, and OTD percentages are independently reconciled against the
current source records and effective filters before publishing. ALL ORDER is
read back from the saved workbook and checked by field and row, along with the
Table1 range and every pivot's current cache record count. Publication is
atomic: a failed generation does not replace an existing output.

Percentage helpers use (SO Complete + Cancel) / the corresponding summary's
quote count, matching rows by category identity rather than cell row number.
A zero denominator gives zero. Native styles, widths, header/data alignment,
and reference legends are retained; refreshed pivot rows are not hidden.

## Sample Reconciliation and Limits

The 16 November sample has 2,069 records; 17 November has 2,075. Both have
67 output columns and 11 native pivots. Validation artifacts are stored outside
the repository under `D:\SECM\ide-list-implementation-validation-20261117`.

| Check | 16 November | 17 November |
| --- | ---: | ---: |
| ALL ORDER values compared | 138,623 | 139,025 |
| ALL ORDER differences | 0 | 2 target labels described below |
| Native pivot grids compared | 11 / 11 match | 11 / 11 match |
| Grid cells, including labels and blanks | 1,327 | 1,377 |
| Independent source/filter total checks | 11 / 11 pass | 11 / 11 pass |

Native grid comparisons use each PivotTable's TableRange1, not the entire PIVOT
worksheet: adjacent percentage helpers are explicitly outside that parity claim.
The 17 November browser test uploads only the two XLSX inputs, waits for the
isolated worker, downloads the workbook, and checks desktop/mobile rendering.
The focused unit/API regression suite contains 149 passing tests. Source/manual
workbook hashes remain unchanged by the validation.

Known manual inconsistencies are not implemented as quote-specific overrides:

- In the 17 November manual, IDE-4096 and IDE-4095 have valid November Target
  RFS dates but retain Target Not Yet Inputted. The generator uses Target
  November. Both have YEAR FAB UPLOAD 1900 and are excluded from the affected
  monthly target pivots, so this difference does not change those sample pivots.
- Target After's manual grand-total helper is 800% because of missing formula
  parentheses and a reference outside the rendered status categories. The
  corrected result is 8/27 = 29.63% on 16 November and 8/30 = 26.67% on
  17 November. These helpers are not native pivot count/sum values.
- Jan-00/1900 placeholders and unusual years remain source/reference-compatible;
  preserving them does not certify the source dates as correct.

The Renewal-name classifier and reference filter selections are evidenced by
these two samples, not a claim that every future business type/filter policy
has been confirmed. A requested filter with no available category fails clearly
instead of silently showing all data. Identical pivot values are not a claim
of byte-identical files or identical formula text (aging is intentionally fixed
to the report date).
