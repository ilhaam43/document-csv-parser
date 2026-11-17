# IDE List Format

## Inputs and Isolation

Report 4 detects the new format by the `List` worksheet and required headers,
not by a report date or quote allowlist. `IDE-list-YYYY-MM-DD.xlsx` supplies the
report date unless the CLI explicitly supplies `--report-date`.

The web form accepts two files: the raw List export and the previous new-format
Daily Tracking IDE workbook containing `ALL ORDER`, `PIVOT`, and `Table1`.
Collabs/GCP is not read, even if supplied. The legacy Dashboard branch continues
to require its CSV and retains its existing calculations and pivot policies.
The public Report 4 web form intentionally has no GCP/Collabs upload control;
the optional CSV API parameter and legacy CLI/backend path remain available.

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

Target After percentage helpers follow the confirmed manual rule:
`SO Complete + Cancel / quote count`, with division performed before addition.
This intentionally permits results above 100%; for example, 8 Complete and no
Cancel gives 800%, not 8 divided by the quote count. Only pivots whose monthly
target filter selects exclusively the After bucket use this rule, regardless of
month, pivot name, or row position. Other completion blocks retain
`(SO Complete + Cancel) / quote count`.
Rows are matched by category identity, and status columns by status name rather
than fixed coordinates. Missing status categories contribute zero; a zero
denominator gives zero. Native styles, widths, header/data alignment, and
reference legends are retained; refreshed pivot rows are not hidden.

### Native Pivot Presentation

`ide_pivot_presentation.py` captures the reference PIVOT cell fonts, sizes,
number formats, full alignment (including wrap and indent), row heights, and
hidden-header selections before replacing the cache. It restores styles by
pivot, row category identity/depth, and measure source/function/calculation.
Matching style cells are applied in bounded multi-area batches to avoid a
separate formatting call for every cell. New categories inherit the same
measure presentation; an empty two-level
Non New pivot retains its own font and relative PM/APM indentation when data
arrives. Percentage helper outlines are repaired without boxing every row.
Section banners resize to current pivot bounds.

Numeric summary measures are centered, following the 17 November manual.
Existing pivot column headings, percentage-helper headings, page-filter captions,
and Department/PM/APM body labels retain their reference alignment. Only a newly
appearing measure heading without a populated reference cell uses the existing
centered fallback. Forced header centering applies only to ALL ORDER; its saved
headings are checked before the output is published.
The 13/16 November references use default alignment in some summary columns;
this display standardization is intentional and does not change values.
Reference widths and styles remain inputs to generation, not values/styles
read from the manual expected output. Widths resized manually on the next day
cannot be inferred exactly from the earlier reference; pixel identity across
different reference layouts is not guaranteed.

The downloaded 17 November sample originally had 881 nonblank font differences
and 178 alignment differences from the manual, plus lost Target After accounting
formats and unwanted percentage-row border lines. Native PivotCache replacement
had reset presentation; value-only comparison did not catch it. The new checks
separate data parity from font/alignment/number-format/geometry comparisons.

### Column Presentation

ALL ORDER presentation is bound by normalized column name, not its previous
position. The reader chooses formats/alignment from populated reference cells;
the first data row alone is not a valid style sample for sparse Progress/date
columns. New raw columns can supply fallback presentation. Source Month Complete
and its reference Month Complete2/3 aliases remain separate from MONTH COMPLETE.
All column headings are centered horizontally and vertically with no indent;
this does not change body-label alignment.

- Progress uses `0.0%` (0.5 displays as 50.0%; the stored value stays 0.5).
  Its native data-bar rule follows the named Progress column after reordering
  and is rebound to the entire current table body, excluding the header. The
  fixed 0-100% scale keeps 81% at 81% bar length even when no row is complete.
  Reference bar appearance is preserved; a missing rule uses the manual's solid
  `#7FD3CA` bar with the percentage visible. A recognizable orphaned progress
  bar left on a date column by older outputs is relocated, not duplicated.
  Saved-file validation rejects missing, partial, duplicate, or auto-scaled
  Progress bars. Other conditional formatting is not deleted or recreated.
- Actual/Target RFS, FAB, CASH, SO, AOS, Created, Updated, Resolved, and Nodin
  dates display as `d mmm yyyy`, like both manuals. Underlying timestamps remain
  unchanged even when time is not displayed.
- Aging is an integer (`0`). Revenue, Revenue Upside, penalties, costs, risks,
  and Revenue (Juta) use the reference's two-decimal accounting format. CEI
  retains its own reference format rather than being guessed as a percentage.
- Temporary text format protects identifiers/month labels during ingestion.
  Final display formats are restored only after values/formulas are written,
  so restoring General does not reinterpret `Nov 2026` as an Excel date.
- Header fills/text colors and alignment follow field identity; a newly added
  Yesterday Status inherits the calculated-month header highlight. Body fonts
  and source-field headers follow the raw export (Calibri 10 in the samples);
  calculated-field headers retain the reference's Indosat Regular 10. Known
  helper fields use the wrap settings shared by both current manual samples.
- Widths follow the named reference column, including monthly/source-month
  aliases. Long unwrapped names can expand automatically (up to 48 characters
  wide); reference widths are never reduced. They are not copied from the
  manual expected output at runtime.

The manual's appearance is not wholly consistent between dates: Actual Cost and
Revenue (Juta) use General on 16 November but accounting on 17 November, and many widths were
manually resized differently. The generator standardizes financial display to
accounting and preserves reference widths; this is not a claim of pixel- or
style-XML identity to both incompatible layouts. No values are rounded or
converted into formatted strings by these display rules.

Progress-bar validation artifacts are stored outside the repository under
`D:\SECM\ide-progress-bar-validation-20261117`. Regenerated 16/17 November
outputs retain all ALL ORDER and PIVOT values and populated-cell base styles
from the preceding outputs. The native bars cover J2:J2070 and J2:J2076,
respectively, with solid teal fill and fixed numeric bounds 0/1. Native Excel
renders confirm 81.0% and 84.3% bars. Separate disposable-copy checks verify
recovery from an older generated reference's orphaned date-column bar and
creation of a missing rule. Original inputs/manuals are not modified.

## Generation Performance

The List adapter now reads source/reference values and column presentation in
one streaming pass per input. PIVOT styles are read in streaming mode; row
geometry comes from the worksheet XML rather than loading every ALL ORDER cell
and deserializing the native PivotCaches. `ide_workbook_package.py` resolves
worksheet/table/cache relationships and validates package metadata without
expanding cache records into Python objects.

Saved-output checks still compare every data value, centered headings, Progress
rule coverage and fixed scale, table bounds, pivot count, and every referenced
cache's source and record count. Independent native pivot/source checks and the
PIVOT formula-error scan remain enabled. Nothing is published before validation.

Native pivot formatting retains its existing bounded range batches. Redundant
alignment capture/restoration was removed from this adapter only; the richer
presentation snapshot already restores those properties. Row-field names are
cached by hierarchy depth within each individual pivot pass, and status-column
metadata is read once per percentage block instead of once per data cell.
Input/style caches are per run; no date, quote, or expected output is cached.
Legacy IDE and the other reports keep their existing generation paths.

The generator result includes `timings_seconds` for input reading, row rules,
native writing/refresh/layout, saved-output validation, and total duration.
There are no new web inputs or user-facing audit files.

Measured locally on the existing samples without the Python profiler:

| Run | Observed Time |
| --- | ---: |
| 16 November CLI, 2,069 records | 63.23 seconds |
| 17 November web worker, 2,075 records | 64.67 seconds |
| 17 November click Generate to file ready, including upload/polling | 66.69 seconds |

These are observations on this machine, not production runtime guarantees.
Profiler runs have substantial overhead and are not directly compared with web
elapsed times. Experimental makepy binding and multi-area ALL ORDER formatting
were discarded: the former was incompatible with existing collection helpers,
and the latter was slower. Neither is used in production code.

Artifacts are outside the repository under
`D:\SECM\ide-performance-validation-20261117`. Both regenerated samples have
zero value, formula, populated-cell style, column-width, row-height/visibility,
or merge differences from the pre-optimization outputs on ALL ORDER and PIVOT.
The 17 November output was generated and downloaded through the actual web form.
This preserves the previously documented differences from the manual; it does
not claim that optimization resolves those business-data discrepancies.

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
worksheet. After the Target After-only formula confirmation, a separate full
PIVOT value comparison also verified every adjacent percentage-helper value
against both manuals. The only remaining PIVOT value difference is the leading
space in the manual's title at A1; native grid and helper numbers match.
Compared with the prior generated outputs, only 11 Target After helper values
change per sample, and populated PIVOT cell styles remain unchanged. Fresh
outputs and audit results are stored outside the repository under
`D:\SECM\ide-manual-formula-validation-20261117`. Native Excel recalculation
checks cover a nonzero Cancel value, absent Cancel, and a zero denominator.
The 17 November browser test uploads only the two XLSX inputs, waits for the
isolated worker, downloads the workbook, and checks desktop/mobile rendering.
The focused unit/API regression suite contains 208 passing tests. Source/manual
workbook hashes remain unchanged by the validation.

Known manual inconsistencies are not implemented as quote-specific overrides:

- In the 17 November manual, IDE-4096 and IDE-4095 have valid November Target
  RFS dates but retain Target Not Yet Inputted. The generator uses Target
  November. Both have YEAR FAB UPLOAD 1900 and are excluded from the affected
  monthly target pivots, so this difference does not change those sample pivots.
- Jan-00/1900 placeholders and unusual years remain source/reference-compatible;
  preserving them does not certify the source dates as correct.

The earlier Target After helper discrepancy (29.63% / 26.67% versus the
manual's 800%) was a formula-policy difference, not a count or Revenue mismatch.
The confirmed Target After-only precedence above replaces that earlier policy;
New Registration and other completion blocks do not inherit it.

The Renewal-name classifier and reference filter selections are evidenced by
these two samples, not a claim that every future business type/filter policy
has been confirmed. A requested filter with no available category fails clearly
instead of silently showing all data. Identical pivot values are not a claim
of byte-identical files or identical formula text (aging is intentionally fixed
to the report date).
