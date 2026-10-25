# Weekly SECM Report Development

## Current scope

The command-line workflow can prepare a Catalyst review package, finalize the
Master, Mobile, and M2M datasets, and write Excel workbooks with refreshed native
PivotTables. It uses the prior week's workbooks as templates and requires Windows,
Microsoft Excel, and pywin32 for the workbook stage. The original templates are
copied; they are never edited in place.

The Catalyst check is a human checkpoint until dashboard access is available.
The Weekly web page and API implement that checkpoint. Both PowerPoint
generators use the prior week's service-specific PPTX as editable templates; they leave the
Power BI dashboard slide content unchanged.

## Weekly web workflow

Open `/weekly` on the running server. Enter the **BQ cutoff/report date, FAB
reporting month, and business report week separately**. For example, W38 uses
24 September 2026 / September 2026 / W38, while W39 uses 2 October 2026 /
September 2026 / W39. The report month selects new FAB rows; the report date
remains the cutoff and the date printed in the PPT. The month defaults to the
report date's month unless explicitly supplied via CLI/API. For example,
Excel `WEEKNUM(date, 14)` evaluates that date as W39 and must not determine the
workbook name or Mobile pivot window. Upload the BQ CSV, previous Master old
format and holiday workbook. Select Mobile, M2M, or both, then upload each
selected service's previous workbook and weekly PPT. Unselected templates are
not required or submitted. The Master always includes all three services.
The page preserves the
job ID when navigating away and back.

After preparation, the page lists only Master rows that need Catalyst review.
Search by quote or customer, query each quote in Catalyst, and select its
returned status. The Generate button becomes available after every listed row
has a status. Selections are kept in this browser for the current job when
navigating away and back. Download `draft-master-data.csv` if the refreshed
Master needs inspection. For the older workflow, expand **File upload and
optional corrections**, download `catalyst-review.csv`, fill its
`reviewed_status` column, and upload it instead. The reviewed `ALL DATA`
workbook remains accepted as a legacy checkpoint. The server always uses its
prepared `master_row` and quote/order identity for direct page submissions.

Jobs prepared before the focused Catalyst checkpoint may contain the entire
Master in their review CSV and have no draft CSV. Such jobs show **Refresh
Catalyst Review**. The server reruns preparation using that job's saved input
files, writes the current review package to `preparation-refreshed/`, and keeps
the original `preparation/` files untouched. Once complete, the direct review
list appears without another upload. The W38 legacy job checked during
development changed from 53,592 full-Master review rows to 66 pending rows.

The API and CLI retain optional attributed correction files for exceptions:
`manual_decisions` can carry reviewed status, completion date, lead time, or PM
approval; `slide_review` can override a confirmed `SLIDE GH MEETING` value.
An optional `pivot_review` CSV records a confirmed change to an `ADJUSTMENT`
filter in a named PivotTable. None of these files is required for normal web
generation. A note in `ADJUSTMENT` alone never approves `ACHIEVEMENT`.

Finalization creates a user-facing ZIP containing only Master Excel and each
selected service's Excel/PPT: three files for Mobile or M2M, five for both.
The API retains the legacy default `output_services=mobile` and accepts
`m2m` or `both`; `previous_m2m` and `previous_m2m_ppt` are the M2M upload fields.
PPT templates are required by the web form, but remain optional for direct API
clients that need Excel only. Audit files
(`audit.json`, `exceptions.csv`, `review-decisions.csv`, and
`bq-source-decisions.csv`) remain in the job's server-side `output/` directory.
Uploaded correction and review files remain in `input/`; they are not included
in the report ZIP. Downloading an older completed job also repackages its
existing reports without internal files, without recalculating report data.
The audit records remaining unresolved Catalyst statuses and the web shows a
warning if any remain; a generated report is not proof that all reviews were
completed.
If a Catalyst review conflicts or Excel fails, the review can be corrected and
submitted again without uploading the source files again. Completed report
artifacts are reused on retry only when input/review/code fingerprints and their
saved SHA-256 checksums match. Changing a review invalidates the saved artifacts.
Native Excel jobs run sequentially in the existing single-worker queue.
Preparation and final
exception CSVs are available from the status panel.

The endpoints are `POST /api/weekly/jobs`, `GET /api/weekly/jobs/{job_id}`,
`GET /api/weekly/jobs/{job_id}/draft`,
`GET /api/weekly/jobs/{job_id}/review`,
`GET /api/weekly/jobs/{job_id}/review-data`,
`POST /api/weekly/jobs/{job_id}/refresh-review` for legacy jobs,
`POST /api/weekly/jobs/{job_id}/finalize`, and
`GET /api/weekly/jobs/{job_id}/download`. Job files and status JSON are saved under
ignored `generated/weekly/api-jobs/`. They are not automatically deleted; server
operators need a retention policy and access control before treating them as a
long-term audit store. A server restart during an active job does not resume it.

## Business rules implemented

- Master services: `Bulk Service`, `Mobile`, and `M2M`.
- Mobile output services: `Bulk Service` and `Mobile`.
- M2M output service: `M2M`, filtered from the same finalized Master, with
  independent sequential numbering. It does not repeat Catalyst review or
  introduce a second lead-time/achievement calculation.
- A source quote already present in the previous master is not appended again.
- New source rows must have `FAB_Upload` in the explicitly selected FAB reporting
  month and no later than the report date. A first-week-of-October W39 report
  may still add September FAB rows.
  The owner's Thursday-to-Wednesday FAB periods were W37: 10-16 September,
  W38: 17-23 September, and W39: 24-30 September 2026. They describe the
  reporting window, not a strict new-row cutoff: the manual W38 Master also
  appended 597 previously missing W37-period rows, and W39 appended one
  W38-period row. Existing Master rows may have much older FAB dates.
- Blank status on a new BQ row follows the historical `In Progress` default.
  An existing unresolved row matched to BQ with both `order_status` and
  `order_completion_date` blank becomes `In Progress` before Catalyst review.
  This models the manual empty-date lookup without relying on Excel's 1900
  display. Existing active rows with no BQ match are flagged `#N/A` in the raw
  status and keep their prior pivot status until Catalyst review. An explicit
  BQ status refreshes active rows; prior `COMPLETE` or `CANCEL` is not
  downgraded by an unreviewed BQ status. Complete and cancel variants are
  normalized to `COMPLETE` and `CANCEL`.
- Week fields match Excel `WEEKNUM(date, 14)`.
- Lead time uses inclusive working days and the supplied holiday list.
- A calculated FAB-CASH lead time of zero is preserved, matching the subsequently
  approved manual-file behavior. CASH-COMPLETE uses a minimum of one day for a
  calculated zero (confirmed by 456 newly added W38 rows). This is a column-level
  calculation policy, not a global replacement of zeros. Published, integer
  FAB-CASH values are retained when both supporting dates remain unchanged.
  Missing/invalid results and rows with changed dates use the same calculation
  as new rows. Missing dates clear stale metrics. Count Cash and both revenue
  totals are always derived from source fields. An explicitly supplied lead-time
  correction takes precedence over the published baseline.
- Negative lead time remains visible as an exception.
- Achievement limits: Mobile 3 days, Bulk Service 5 days, M2M 19 days.
- PM approval must be explicitly recorded in the review CSV to override the
  calculated achievement. A reviewed workbook can also explicitly change
  `NOT ACHIEVED` to `ACHIEVED`; an `ADJUSTMENT` note alone is not approval.
  Published historical CASH-COMPLETE and achievement values are carried from the
  prior master when the supporting dates do not change. Their presence does not
  prove that historical PM approval was obtained. A new approval is never
  inferred from a note. A date change invalidates the inherited SLA calculation.
- New-row sales hierarchy and activation date stay blank to match the reviewed
  master contract.
- A reviewed Catalyst workbook can supply corrected `order_status` and a
  nonblank `order_completion_date`; derived completion fields are recalculated.
  Blank completion-date deletion still requires a reviewed status change or an
  explicit correction workflow.
- Existing rows match BQ on quote and order where available. When the first BQ
  row for a quote disagrees with an active row's matched order status, that row
  appears in `source-review.csv` for audit. New rows are compared with the first
  selected BQ row for their quote. Closed `COMPLETE`/`CANCEL` rows retain their
  lifecycle. Active rows refresh status and completion from the first quote
  match, reproducing the manual lookup workflow. This also applies to newly
  appended active rows, while newly appended closed rows retain their own
  source lifecycle. Cash dates, creation dates, order identity and financial
  fields stay with the matched order. Explicit source overrides remain optional.
  Ambiguous quote-only
  matches remain in `exceptions.csv`.
- Mobile week pivots retain a rolling four-week window and use the prior Mobile
  workbook as their template. Five M2M-only pivots are removed.
- Adjacent formula columns inherit each refreshed pivot row's fill, bold, and
  horizontal borders while retaining their own numeric format. Grand Total
  styling follows the actual total row when a pivot grows or shrinks; obsolete
  helper formulas and their formatting are cleared without removing unrelated notes.
- Performance summary pivots include all ADJUSTMENT notes in both
  Master and Mobile. This changes aggregation membership, not ACHIEVEMENT or PM
  approval. A note is not an automatic exclusion. SLIDE GH MEETING and
  detail/exception sheets retain their separate operational/investigation
  filters; an explicit pivot-review input is applied after
  the default summary policy. Manual samples with additional exclusions remain
  differences, not evidence that the default reproduces those selections.
- Every refresh includes finite, nonnegative FAB-CASH values in page filters,
  including zero, for both report roles. The policy no longer depends on a month
  rollover or cached numeric checkboxes. Negative and missing values remain in
  ALL DATA but are excluded from these average-lead-time populations.
- Summary column fitting covers numeric values, bounded/wrapped column headers,
  page-filter captions and values, and adjacent percentage helpers. It only
  widens columns as needed and does not apply sheet-wide AutoFit to ALL DATA.
- FAB timestamps preserve available time precision. Lifecycle report dates
  (`cash_dt`, `order created`, `order_completion_date`) are date-only values.
- Missing completion dates clear stale completion counts, periods and calculated
  cash-completion SLA values, even when the date was already blank last week.
  An existing published ACHIEVED value accompanied by ADJUSTMENT is preserved;
  merely filling an adjustment note does not grant a new approval.

## Create a preparation package

Run this before any Excel-writing stage:

```powershell
python -m weekly.prepare `
  --bq-csv "D:\SECM\Weekly 38\bq-results-20260924-090453-1790240701004.csv" `
  --previous-master "D:\SECM\Weekly 38\Master Raw Mobile Service - W37 2026 old format.xlsx" `
  --holiday-workbook "D:\SECM\Weekly 38\list national holiday.xlsx" `
  --report-date 2026-09-24 `
  --output-dir "generated\weekly\w38-preparation"
```

Outputs:

- `new-master-rows.csv`: selected rows in the 34-column master schema.
- `draft-master-data.csv`: refreshed prior Master plus selected new rows.
- `catalyst-review.csv`: only unresolved draft rows for Catalyst status review.
- `source-review.csv`: audit of BQ differences between exact order and first
  quote match; it is not a required review queue.
- `slide-review.csv`: template for attributable manual slide-summary values.
- `exceptions.csv`: invalid dates, negative lead time, and other review items.
- `audit.json`: counts, rejected-row reasons, and output locations.

The `generated/weekly` directory is ignored by Git.

## Complete Catalyst review

Query each listed `quote_num` in Catalyst and select `In Progress`, `Complete`,
or `Cancelled` on the web page. Direct submission requires a decision for every
listed row; the API accepts them as a `reviewed_statuses` JSON form field keyed
by `master_row`. In the CSV fallback, enter the returned status in
`reviewed_status`. This is the only column users edit in the generated
`catalyst-review.csv`; its other columns identify the Master row and provide
context. Older CSVs with additional correction or attribution columns remain
accepted. In the CSV workflow, blank statuses remain unresolved and the prior
pivot status is retained. A reviewed CSV or an existing reviewed workbook with
an `ALL DATA` sheet can be passed to `weekly.finalize`. If the workbook is the
baseline, pass a separate CSV through
`--manual-decisions` for status/date/PM/FAB-CASH corrections; this separate
correction CSV still requires `reviewed_by` and `review_reason`. A reviewed XLSX
checkpoint can restore `fab_upload` seconds for uniquely matched new rows.
The Master and Mobile `No.` columns are both regenerated as consecutive integers
starting at 1; the prior Master's fractional numbering is not retained.
The web applies the closed-order/active-quote lifecycle rule automatically, so only Catalyst status
review is required. BQ quote-first differences remain downloadable under
optional corrections. For an owner-approved exception, fill `reviewed_source`
with `quote` (or `order`) on only the affected rows in `source-review.csv` and
upload it; blank rows retain the default. The API also accepts a sparse
`reviewed_sources` JSON object keyed by `master_row`. Finalization validates
the identity of every supplied override and writes the effective choice and
method for every conflict to the server-side `output/bq-source-decisions.csv`.
Overrides can change completion, achievement, pivot totals, and PPT charts;
they must not be inferred from the FAB reporting period. Catalyst decisions
remain separate.
A reviewed non-complete status clears its completion date and dependent fields;
the raw cancellation label is `Cancelled`, while pivot statuses use `CANCEL`.
For the W38 regression only, the supplied `old format cek manual via catalist`
workbook is the reviewed source.

The W38 BQ export and W37 Master are available, but the W37 BQ export is not.
Use the manually completed W37 Master as the starting point for W38. It is not
valid to replay W37 using the W38 BQ export. The W38 manual pre-Catalyst
checkpoint marks 137 existing `#N/A` rows as `In Progress` because the quote
matches BQ while both its status and completion date are blank. Preparation
now applies that rule before producing the Catalyst review: 772 new rows are
selected and 66 existing rows remain unresolved because their quotes are absent
from BQ. In the final manual workbook, 64 of those 66 statuses become
`In Progress` and two become `Cancelled`; those remaining decisions still
require Catalyst review.

## Finalize data

```powershell
python -m weekly.finalize `
  --bq-csv "D:\SECM\Weekly 38\bq-results-20260924-090453-1790240701004.csv" `
  --previous-master "D:\SECM\Weekly 38\Master Raw Mobile Service - W37 2026 old format.xlsx" `
  --catalyst-review "D:\SECM\Weekly 38\Master Raw Mobile Service - W38 2026 old format cek manual via catalist.xlsx" `
  --source-review "<reviewed-source-review.csv>" `
  --holiday-workbook "D:\SECM\Weekly 38\list national holiday.xlsx" `
  --report-date 2026-09-24 `
  --output-dir "generated\weekly\w38-final-data"
```

The final data package contains `master-final-data.csv`, `mobile-data.csv`,
`review-decisions.csv`, `exceptions.csv`, and `audit.json`. Review the exceptions
before publishing. Relative to the supplied pre-Catalyst checkpoint, the W38
manual final includes 64 further `In Progress` decisions and seven other edits:
two cancelled statuses, one completion date, two PM achievement approvals, and
two zero FAB-CASH lead times. These require explicit, attributed review decisions;
they are not embedded as hardcoded report rules. Zero FAB-CASH corrections are
accepted only when the source dates produce zero business days.

## Write Excel workbooks

```powershell
python -m weekly.workbook `
  --data-csv "generated\weekly\w38-final-data\master-final-data.csv" `
  --template "D:\SECM\Weekly 38\Master Raw Mobile Service - W37 2026 old format.xlsx" `
  --output "generated\weekly\w38-final-data\Master Raw Mobile Service - W38 2026 old format generated.xlsx" `
  --report-date 2026-09-24 `
  --report-week 38

python -m weekly.workbook `
  --data-csv "generated\weekly\w38-final-data\mobile-data.csv" `
  --template "D:\SECM\Weekly 38\Master Raw Mobile Service - W37 2026.xlsx" `
  --output "generated\weekly\w38-final-data\Master Raw Mobile Service - W38 2026 generated.xlsx" `
  --mobile `
  --report-week 38 `
  --report-date 2026-09-24
```

Set `--report-week` to the report's Excel week number. When omitted, the Mobile
writer infers it from the latest `fab_upload`, which may be earlier than the
reporting week when no new data arrived. Use the explicit option in production.
The writer validates the Table1 shape, sample cell types/dates, PivotTable names,
and pivot cache source before publishing each workbook. It also reveals a pivot's
hidden trailing rows when the refreshed table extends into a hidden separator
inherited from the prior week's template. Existing hidden filter rows and
non-empty hidden sections are preserved; a hidden pivot end blocks publication.
`python -m weekly.compare --reference <manual.xlsx> --generated <output.xlsx>`
checks every sheet's cached cell values, chart source references, and the
visibility of rows containing data. Blank separator-row hide differences do
not affect that parity check. Add `--business-only` to compare every business
cell while ignoring only `ALL DATA` numbering and seconds within `fab_upload`;
this mode also reports visibility differences but does not treat them as data
differences. It still checks status, lead times, achievement, pivot cells, and
chart source references exactly.

The writer rebuilds `SLIDE GH MEETING` chart-source values from its refreshed
pivot. Completed months are collapsed to weighted monthly pivot subtotals;
the reporting month expands into weeks. The reporting month comes from the
Thursday-start report week, not the BQ export month (W39 remains September
even with an October export). Both Master and Mobile retain the full precision
of the refreshed pivot, including weekly values and the weighted grand total.
The template's number formats control display precision without rounding the
stored cell values.
All chart series are resized to the resulting range. Helper percentage formulas
are restored relative to each refreshed pivot's data and grand-total rows.
Both Master and Mobile include finite, nonnegative numeric FAB-CASH values in
summary page filters on every refresh, even when a prior template hid them.
Zero participates in the average; negative/blank metrics do not. The old
month-rollover-only reset and Mobile numeric checkbox inheritance are no longer
used. Separate operational/detail filters remain unchanged. Other explicitly
reviewed values belong in
`slide-review.csv` with columns
`workbook,cell,value,reviewed_by,review_reason`. These are optional, attributed
exceptions, not a required way to copy stale manual summaries. No review CSV
is needed to populate normal summaries or charts. A reference fixture copied
from the manual file is useful for regression, but is not production approval.

W39's manual Master also changes two `ADJUSTMENT` pivot filters compared with
W38, including four August orders. The default performance-summary policy now
includes all notes for both roles, covering those orders without a W39-specific
override. An additional exclusion or exceptional ordering still needs an explicit
review decision rather than inference from note text. Such optional decisions
can be supplied in an attributed `pivot-review.csv` alongside the Catalyst
review, or with `--pivot-review` to `weekly.workbook`. Its columns are
`workbook,pivot,field,item,visible,position,reviewed_by,review_reason`.
For an `ADJUSTMENT` filter decision, set `visible` to `true` or `false` and
leave `position` blank. For an approved `Week FAB` or `Week Completion` order,
set a positive `position` and leave `visible` blank. `field` and `position`
may be omitted in a filter-only CSV. For W39, the confirmed adjustment item is
visible in Master `PivotTable16` and `PivotTable20`, while the manual file puts
`W39` at position 4 in `PivotTable14`. Do not reuse these decisions for another
week without review.

## W38 and W39 reference checks

The W38 sample selects 772 new Master rows. The W39 sample uses 2 October
2026 as the BQ cutoff, September 2026 as the FAB month, and W39 as the report
week; it selects 1,270 new rows. The resulting W39 Master and Mobile datasets
have 54,862 and 54,372 rows. In reference testing, the reviewed manual Master
was used as a Catalyst checkpoint. This validates replay of the supplied
review decisions and pivot output; it does **not** establish that all status
decisions can be inferred from BQ without Catalyst access.

The W38 BQ conflict audit has 10 rows; W39 has 34. Analysis with the W40 BQ
snapshot from 8 October identified a consistent lifecycle rule across all
three samples: preserve closed rows and refresh active rows by first quote
match. W38's nine conflicting new rows are already closed; W39's 24 conflicting
new active rows refresh by quote while its six closed rows retain their order
status. Existing active rows follow quote lookup in all three samples. No
quote-specific or week-specific source overrides are needed for these patterns.
Catalyst statuses and new PM approvals are still external review inputs, not
facts inferred by this rule. The new numbering rule intentionally differs
from the manual `No.` column. BQ timestamps are only minute-precision; that is
sufficient for this report's accepted data parity.

The W39 reference also needs attributed pivot and slide review CSVs for its
manual filter, week-order, and chart-source edits. These are separate from
the generator's general business rules. The finalization audit still lists two
negative lead times and 15 quote-first matches; inspect those exceptions before
publishing. The manual PowerPoint has stale 2025 labels on two charts with
2026 data. The generated deck uses 2026 labels while preserving the same chart
numbers.

The BQ files supplied for W38/W39 only contain `FAB_Upload` to the minute,
while the manual workbooks retain seconds. This is accepted: date, hour, and
minute must match. The W39 manual Master's fractional `No.` series skips 20
positions across 16 locations, so the generator deliberately uses sequential
integers instead. `audit.json` retains the count of new rows without source
seconds for traceability; it is not a distribution warning. Use
`python -m weekly.compare --business-only --reference <manual.xlsx> --generated
<output.xlsx>` to validate these agreed fields. Historical replay using
decisions derived from a target manual workbook does **not** prove those
decisions can be inferred automatically for a future week.

### Historical fixture-based replays

The following counts describe earlier runs with target-derived review inputs,
not the current default generator. The later October validation below reports
the current no-source-override results and remaining differences explicitly.

On the earlier reviewed W39 replay, the strict comparison reported 2,539
differing Master cells: 1,270 in `No.` and 1,269 in `fab_upload`. Mobile had
1,263 differences, all in `fab_upload`; the other sheet values matched. Under
the agreed business comparison, that replay matched all manual workbook cells,
including pivot and slide-summary values; all six Mobile PPT charts matched
numerically. That W39 replay included reviewed
Catalyst, BQ source, pivot, and slide decisions derived from the reference.

On the earlier fully reviewed W38 replay, every pivot and slide-summary
cell matched the manual files. Only two `ALL DATA` cells per workbook differed:
Excel row 15316 in Master (15089 in Mobile) is `Cancelled` with no completion
date, but the manual retains stale `LEADTIME CASH-COMPLETE = 5` and
`ACHIEVEMENT = NOT ACHIEVED`. The generator clears both derived values.
Five of six PPT charts match without extra review; chart 4's historical W37
point needs the recorded presentation review to match the manual deck.
The W38 manual deck also keeps one historical lead-time point at 2 where the
current workbook's pivot yields 1. To reproduce that approved presentation
decision, upload `presentation-review.csv` with columns
`chart,series,label,value,reviewed_by,review_reason` (for this case:
`4,2,W37,2,...`). The writer updates both the chart cache and its editable
embedded workbook. No report-week-specific point is hardcoded in the generator.

After pivot refresh, the writer grows numeric summary columns to fit their
displayed values, sizes top service filters with dropdown clearance, and fits
the numeric helper column beside the status summary. It preserves wider
template widths and caps filter growth to avoid excessive horizontal spread.
Helper styling is synchronized after Excel's full recalculation, revealing
hidden pivot totals and fitting summary columns. Hidden total rows report
different display borders, so copying styles before unhiding them loses the
visible total separator. This final styling pass does not rewrite formulas or
report values, and runs before automatic calculation is restored.

## Write Mobile PowerPoint

### Earlier October snapshot validation

These runs preceded the unconditional historical FAB-CASH recalculation below.
They describe the tested engine at that time, not current whole-workbook parity.

The W40 retry with the original 8 October BQ snapshot confirmed 97 new rows,
54,959 Master rows and 54,464 Mobile rows. With only 67 Catalyst status decisions
replayed as test inputs, all numerical cells in the four pivot/monitoring sheets
match the original manual W40 Master and Mobile files, and all six Mobile PPT
chart series match numerically. This does not make the entire workbooks identical:
one valid FAB-CASH value is blank in manual, and SLIDE GH MEETING retains a
spelling difference and two Mobile static-source rounding/copy differences.
Those differences remain visible in the comparison results.

A fresh local browser run on 19 October 2026 exercised the complete W40 upload,
67 direct Catalyst selections, generation and ZIP download, without mocked API
responses. It used report date `2026-10-08`, reporting month `2026-10`, week 40,
the 8 October BQ export and W39 templates. Catalyst statuses were replayed only
as test inputs from the original manual W40: 66 InProgress and one Cancelled.
The final run after the hidden-total border fix completed in approximately
10 minutes 46 seconds on the test machine; this includes automated review
entry, not real dashboard review time. The ZIP
contained only the Master workbook, Mobile workbook and Mobile PPT. All six
worksheets in each downloaded workbook matched the previously validated W40
values, chart references and content visibility. Review selections survived a
page reload, and desktop and 390-pixel mobile views had no page overflow or
browser errors. Native Excel inspection found no helper formatting differences
across five Master and four Mobile tables, and all six PPT chart series matched
the manual W40 values while preserving template scales. The automated suite
ran 226 tests without failures, with two opt-in tests skipped.
This confirms the tested local workflow, not deployed-server
status or whole-workbook equivalence to every manual sample.

The W38/W39 data regressions no longer need BQ-source overrides. W38 still
needs two external PM approvals to reproduce the final manual achievement values;
both older samples retain stale completion SLA values on a cancelled order that
the generator clears. Native generation was also rerun for both weeks: five of
six PPT chart parts match numerically for each; the remaining historical chart
points differ from the latest pivots. W39 Mobile PIVOT PERFORMANCE matches, while
W39 Master retains explicit adjustment-filter differences and intentional ascending
week ordering instead of the manual's W39-before-W4 placement. Rebuilt static
summaries also expose manual historical values that differ from current pivots.
These are not whole-workbook parity passes. Catalyst and PM decisions replayed from an expected manual file are
test fixtures, not evidence of automatic access to those external facts.

All PPT historical lead-time points now refresh from the generated workbook.
The prior behavior that silently retained FAB-CASH points from the template is
removed. Fixed template axes and category-label spacing are preserved so weekly
and monthly chart shapes remain comparable. Values outside a fixed axis are not
clamped, replaced, or silently hidden: an editable chart overlay names the
out-of-scale point and its actual value. Old generated overlays are replaced or
removed when the next report changes. Integer chart caches and embedded workbook
cells share the same thousands-separated number format.
Explicit presentation reviews still override
approved historical exceptions. Layout, chart editability and non-chart media
remain based on the previous-week deck.

Use the previous week's deck as the template. For W38, use
`Mobile Services W37 - 2026 (Weekly).pptx`; the input W37 file is never edited.
The generator reads six chart series from the final Mobile workbook's refreshed
`PIVOT PERFORMANCE` sheet: weekly Bulk and Mobile FAB quantity/revenue,
weekly FAB-CASH/CASH-COMPLETE lead times, and monthly all-service
quantity/revenue and CASH-COMPLETE lead times. It updates both the visible
chart caches and their embedded Excel workbooks, so chart data remains editable.
In **Edit Data**, each weekly chart shows only the rolling four report weeks;
older week columns are hidden, not deleted. For example, W38 shows W35-W38
and hides W34; W39 shows W36-W39 and also hides W35.
The server needs the updated `requirements.txt` installed (`lxml` is required
for preserving the PPTX XML namespaces).
The current month's labels use the report year rather than stale labels carried
by the prior deck. Slide 4's Power BI/dashboard content is not regenerated.
Slide dates and week footers are updated; the date is made static to prevent
PowerPoint from replacing it with the computer's current date.

```powershell
python -m weekly.presentation `
  --mobile-workbook "generated\weekly\w38-final-data\Master Raw Mobile Service - W38 2026.xlsx" `
  --template "D:\SECM\Weekly 38\Mobile Services W37 - 2026 (Weekly).pptx" `
  --output "generated\weekly\w38-final-data\Mobile Services W38 - 2026 (Weekly).pptx" `
  --report-date 2026-09-24 `
  --report-week 38
```

Chart values come from the generated Mobile pivot, not copied from manually
edited PPT numbers. The supplied manual W38 PPT differs from its W38 workbook
at two historical Mobile lead-time points and has several 2025 month labels on
2026 data. Those manual discrepancies are not reproduced. Review the final
PPT against the workbook before distribution.

## Audit any weekly reference

`weekly.regression` calls `finalize_weekly_data`, the same finalizer used by the
web API. It does not contain week-specific expected counts, quote-specific
exemptions, or a second implementation of source selection. Master and Mobile
are compared independently, keyed by quote, order, and FAB minute. Duplicate
occurrences, missing and extra rows, and every business column are checked.
Only sequential `No.` and FAB seconds are excluded by the owner's agreed policy.
Numeric comparisons use an absolute tolerance of `1e-8`; text is not rewritten
to make a comparison pass.

The audit separately reports reference FAB-CASH and revenue inconsistencies
against source formulas. It does not repair the reference or excuse an actual
parity mismatch. An audit failure therefore remains a failure even when its
cause is understood. The data audit does not certify workbook layout, pivot
filters, or PPT parity; those need their own native workbook/chart comparisons.

For another sample, replace the paths and report date below. Pass an explicit
`--report-month YYYY-MM` when the FAB month differs from the BQ cutoff month.
Approved correction inputs can be passed as `--manual-decisions` and
`--source-review`; they remain explicit inputs, not inferred expected values.

```powershell
python -m weekly.regression `
  --bq-csv "D:\SECM\Weekly 38\bq-results-20260924-090453-1790240701004.csv" `
  --previous-master "D:\SECM\Weekly 38\Master Raw Mobile Service - W37 2026 old format.xlsx" `
  --expected-master "D:\SECM\Weekly 38\Master Raw Mobile Service - W38 2026 old format.xlsx" `
  --expected-mobile "D:\SECM\Weekly 38\Master Raw Mobile Service - W38 2026.xlsx" `
  --holiday-workbook "D:\SECM\Weekly 38\list national holiday.xlsx" `
  --catalyst-review "D:\SECM\Weekly 38\Master Raw Mobile Service - W38 2026 old format cek manual via catalist.xlsx" `
  --report-date 2026-09-24
```

The supplied pre-Catalyst W38 checkpoint is not a complete set of final Catalyst
and PM decisions. Its audit is expected to expose differences, not report a
golden-file pass. No known quote is exempted, and no missing decision is inferred
to make the regression pass. The older fixture-specific counts above describe
historical runs, not the current comparison result.

The same regression is available through `tests/test_weekly_reference_w38.py`:

```powershell
$env:WEEKLY_REFERENCE_DIR = "D:\SECM\Weekly 38"
python -m unittest tests.test_weekly_reference_w38 -v
```

## Cross-week calculation ownership

The October 2026 review distinguishes reproducible calculations from published
historical decisions. This policy is the same for every week:

| Field | Calculation or ownership |
| --- | --- |
| Historical FAB-CASH | Retain a published integer with unchanged FAB/CASH dates. Calculate missing/invalid results and changed-date rows; clear the value if a required date is missing. Explicit reviewed corrections take precedence. |
| New FAB-CASH | Inclusive NETWORKDAYS(FAB, CASH, uploaded holidays), keeping zero and negative results. |
| Count Cash | Always derive from the presence of a CASH date. |
| OTC+MRC | Always sum the row's raw OTC and MRC with decimal arithmetic. |
| OTC+MRC (inMio) | Always divide the total by 1,000,000, without rounding away source precision. |
| Historical CASH-COMPLETE | Retain a published integer with unchanged supporting dates; recompute missing/invalid results or after a date correction. Missing lifecycle dates clear stale SLA values. |
| New CASH-COMPLETE | Inclusive NETWORKDAYS(CASH, completion, holidays), normalizing calculated zero to one. |
| Historical achievement | Preserve the existing published result with unchanged dates. Do not treat this as newly inferred PM approval. |
| New or corrected achievement | Apply the service SLA; only an explicit reviewed approval can override it. |
| Pivot and PPT | Refresh from the resulting dataset and workbook, not the target manual report. |

Read-only inspection of the original W39-W41 references found incompatible
historical calculation bases: W39/W40 retain normalized FAB-CASH constants and
values that differ from the supplied holiday calendar, while W41 recalculates
them. In all three samples, thousands of
published CASH-COMPLETE constants instead match FAB-to-completion, while the
confirmed new-row rule is CASH-to-completion. These historical SLA values are
not silently reinterpreted or overwritten. W41 also leaves both derived revenue
columns empty on all newly appended rows even though raw OTC/MRC are populated.
The generator calculates the revenue rather than dropping it to imitate blanks.

Consequently, a single source-consistent calculation cannot be byte-for-byte or
cell-for-cell identical to every unreconciled manual reference. A consistent
historical SLA restatement needs owner approval and a corrected reference; it
must not be hidden in a week-specific rule. Original source files are never
edited during validation. Status-only Catalyst replay fixtures are test inputs,
not evidence that the application can query Catalyst automatically.

### Earlier validation before published FAB-CASH preservation

The updated production prepare/finalize/workbook/PPT functions were rerun with
the original W39, W40, and W41 BQ snapshots and prior-week templates. Only pending
Catalyst statuses were replayed from each original manual reference: 66, 67, and
67 respectively. No PM, BQ-source, pivot, slide, or presentation overrides were
supplied. Results and previews are stored outside the repository at
`D:\SECM\weekly-crossweek-consistency-20261023`.

| Sample | Master rows | Mobile rows | Remaining ALL DATA differences from manual | PPT charts matching manual |
| --- | ---: | ---: | --- | ---: |
| W39 | 54,862 | 54,372 | Historical FAB-CASH: 249 Master / 245 Mobile; one stale cancelled-order SLA and achievement in each. | 5/6 |
| W40 | 54,959 | 54,464 | Historical FAB-CASH: 250 Master / 246 Mobile. | 6/6 |
| W41 | 56,555 | 56,058 | Both revenue fields blank in manual on 1,596 new Master / 1,594 new Mobile rows. | 3/6 |

All generated source-metric checks pass. Each saved workbook's complete ALL DATA
matches its finalizer CSV, numbering is sequential, all 18 PPT charts match
their generated pivot sources and editable embedded workbooks, and all template
images are preserved. Native Excel independently checked 96 representative
NETWORKDAYS values across positive, zero, negative, weekend, and year categories.
No helper-style mismatches were found across 27 helper tables. All 12 slides
were opened and rendered in PowerPoint. Original input SHA256 hashes are unchanged.

The data consistency and native checks are not manual parity passes. Raw cell
comparisons still report historical filter/order and static summary differences.
The W41 revenue diagnostic that completes only a copy of the manual rows in
memory leaves zero ALL DATA differences, but the original reference remains
unchanged and the actual parity audit still fails. `weekly.regression` was also
executed against W41 and correctly exited nonzero for that unresolved difference.
The standard unit suite includes an API test with the real finalizer/rule engine;
only file adapters and native writers are substituted there. Both opt-in W38
reference tests were separately executed and passed, including the test that
expects the incomplete legacy checkpoint to expose differences.

## Earlier summary population and layout validation

The performance-summary policy is metric-specific. PIVOT PERFORMANCE includes
all ADJUSTMENT notes for both roles. SLIDE GH MEETING and the monitoring and
exception sheets retain their separate operational populations; resetting their
filters globally changes otherwise matching reference values. Explicit optional
pivot reviews are still applied last. No additional review is required for
normal generation.

Numeric FAB-CASH page filters in summaries are rebuilt from finite values >= 0
on every run, rather than inheriting a service-specific numeric checkbox list.
This deliberately includes zero. Layout fitting is bounded to pivot/filter/
helper ranges, preserves wider columns, wraps oversized headers, and preserves
intentionally hidden rows when AutoFit changes header heights.

The W39-W41 validation compares PIVOT PERFORMANCE by pivot name, row label and
column label as well as the strict whole-workbook coordinate comparison. This
avoids classifying reordered weeks as numeric discrepancies. With the supplied
published history preserved, count and achievement cells match all three
reference samples. Revenue matches W39/W40; W41 retains the documented derived
revenue blanks in the reference. FAB-CASH averages still differ where historical
values or reference numeric selections differ from the declared calculation
policy. These are not full manual-parity passes.

The workbook-only patch reuses the verified prepare/finalize CSV datasets and
re-runs the same native writers called by the web API. Independent aggregate
checks read saved pivot filters and recompute totals and displayed period groups
from source rows; these checks do not use manual target values. Native checks
also cover saved ALL DATA, formula/number overflow, hidden rows, helper fill/
bold/borders, editable chart data and rendered tables. Validation files and
previews are outside the repository at
`D:\SECM\weekly-pivot-validation-20261023`, not in the user's report ZIP.

## Published-history refresh policy

The current policy is `preserve_published_history_v1`. The preceding validation
tables describe earlier runs with unconditional historical FAB-CASH calculation;
they are not claimed as results for the current policy.

- The prior published master is a baseline. Valid historical lead times remain
  stable when their supporting calendar dates do not change, even if the uploaded
  holiday calendar has changed. Invalid/missing values are repaired from dates.
- A changed CASH/completion date invalidates the corresponding inherited metric.
  Missing achievement is calculated from the final CASH-COMPLETE value and the
  service SLA. An adjustment note does not grant PM approval.
- An explicit reviewed FAB-CASH correction or PM approval/rejection takes
  precedence. An inherited ACHIEVED value is not newly certified as PM-approved.
- Revenue remains an arithmetic calculation, not a manual exception. Missing
  revenue formulas are completed instead of copying blanks into pivots.
- `SLIDE GH MEETING` values use current pivot precision. Older manually typed
  numbers and premature integer rounding are not reused as calculation inputs.
- PowerPoint chart caches and editable data also retain pivot precision. Their
  `#,##0` number formats and summary text annotations provide rounded display
  values, without replacing fractional lead times or revenue in the chart data.
- `metric-decisions.csv` in the internal job directory records row identity,
  metric, previous/calculated/final values and the decision method. `audit.json`
  records policy version, the supplied holiday dates, method counts and formula
  variances. Neither file is added to the normal core-report ZIP or requires
  another web review step.

W39's original Master demonstrates the reconciliation issue directly: its
`SLIDE GH MEETING!T18` retains 4 while its own `N19` pivot cell contains
2.090909090909091. The static value and its AVG Total also equal the preceding
W38 static table. A consistent generator follows the refreshed pivot rather
than reproducing both contradictory values.

W41 restates old FAB-CASH values without changing the supporting dates. The
current preservation policy does not silently infer that restatement from the
week number. An intentional historical correction must be supplied explicitly;
matching every unreconciled manual cell is not a generation guarantee.

### Validation of the published-history policy

Fresh production prepare/finalize and native Excel/PPT runs are stored outside
the repository at `D:\SECM\weekly-history-validation-20261023`. These results
supersede the older unconditional-recalculation run above:

| Sample | Remaining ALL DATA differences from manual (Master / Mobile) | PPT charts with matching rounded label values |
| --- | --- | ---: |
| W39 | One cancelled-order CASH-COMPLETE and achievement in each; FAB-CASH matches. | 5/6 |
| W40 | One missing FAB-CASH result in each; generator calculates the missing value. | 6/6 |
| W41 | Historical FAB-CASH restatement: 249 / 245; both revenue columns blank in manual on 1,596 / 1,594 new rows. | 3/6 |

All 2,730 independent pivot aggregates match final data and saved filters, all
180 static-summary values match native GetPivotData, and 127 calculated metric
samples match native NETWORKDAYS with the declared per-column zero policy.
The six saved ALL DATA datasets match their finalizer CSVs with sequential
numbering. All 18 chart caches match generated pivot precision; editable data
matches within floating-point serialization tolerance and all template images
remain unchanged. All 12 slides were rendered in PowerPoint. Static summary
styles now move with period/total rows instead of leaving old footer formatting
on the newly inserted week, without changing the saved summary values.

W40 using the generated W39 baseline and W41 using the generated W40 baseline
produce the same final datasets as using the respective published manual
baselines. Input hashes remain unchanged. Pending Catalyst statuses were
replayed solely for testing; no new automatic Catalyst access is implied.

The standard suite runs 253 tests: 251 pass and two opt-in reference tests are
skipped when their environment variable is absent. The API test exercises the
real finalizer with file adapters/native writers substituted, checks historical
preservation and revenue completion, and asserts that the user ZIP contains
only the named core reports. This is not a full manual-parity claim: declared
filter differences, manual restatements, blank reference revenue and stale
summary/chart values remain explicitly reported.

## Next implementation stages

1. Add job retention and startup recovery after server deployment requirements are set.

## M2M Report Implementation

`weekly/profiles.py` holds workbook-specific service scope, weekly window and
summary location. `weekly/workbook.py --report-type m2m` uses the same native
Excel writer as Master/Mobile, without removing the M2M-only pivots.
`weekly/m2m_presentation.py` maps the M2M pivot sections to the existing editable
M2M deck; chart XML, embedded data, layout preservation and validation reuse
`weekly/presentation.py`. No week-specific or quote-specific target values are
used. Master processing, status resolution, holidays and the 19-day M2M SLA
remain shared with the existing pipeline.

Standalone M2M commands after finalization:

```powershell
python -m weekly.workbook --data-csv .\output\m2m-data.csv --template .\previous-m2m.xlsx --output .\output\m2m.xlsx --report-type m2m --report-week 40 --report-date 2026-10-01
python -m weekly.m2m_presentation --m2m-workbook .\output\m2m.xlsx --template .\previous-m2m.pptx --output .\output\m2m.pptx --report-week 40 --report-date 2026-10-08 --report-month 2026-10
```

The supplied M2M templates have four sheets, 20 PivotTables and four editable
PowerPoint charts. Their legacy `MONITORING ORDER - MOBILE` sheet name is
preserved. `SLIDE GH MEETING` uses Q:R instead of Mobile's S:U. The four extra
empty `Table1` columns are retained because pivot metadata references them;
populated unknown extra columns fail with an explicit mapping error instead
of being discarded. Only the canonical 34 columns receive generated data.

The weekly PPT charts show the last four report weeks; earlier embedded columns
remain hidden. The monthly lead-time chart has one M2M series, unlike Mobile's
two. Monthly labels derive their year from the selected reporting period,
correcting stale 2025 labels in the reference decks. Revenue preserves fractional
millions with `#,##0.00` display. Monthly lead-time labels sit above their points
instead of overlapping month labels, without moving chart frames or axes.
Missing M2M periods use zero for count/revenue
and a gap for an undefined lead-time mean, never a fabricated zero-day mean.
Power BI slide imagery remains unchanged.

Legacy optional slide/pivot/chart correction files still apply to Master/Mobile;
their cell/chart coordinates are not silently applied to M2M. Normal M2M output
uses refreshed native pivot values and requires no additional manual checkpoint.

The M2M ALL DATA in the supplied W39/W40/W41 files equals the corresponding
manual Master's M2M subset (490 / 495 / 497 rows). This validates the shared-data
contract, not universal manual parity: the historical Master/filter differences
documented above propagate to the M2M report too. Static manual averages may
also be rounded while the generator retains pivot precision. Do not report
these known differences as a fully identical manual-output pass.

Development samples, native outputs, comparisons and rendered previews are kept
outside the repository at `D:\SECM\m2m-development-20261025`.

### M2M Validation Results

The real multipart prepare/review/finalize API flow for W40 produced all five
core reports in the combined ZIP. Catalyst statuses were replayed from the
supplied manual sample solely for this test; no live Catalyst access is implied.
Native Excel/PowerPoint checks cover M2M W40 and W41:

| Sample | M2M rows | Independent pivot aggregates checked | Remaining ALL DATA differences from manual |
| --- | ---: | ---: | --- |
| W40 | 495 | 247, all pass | None |
| W41 | 497 | 248, all pass | Four inherited FAB-CASH values; four calculated revenue values versus manual blanks |

All generated M2M data cells match their finalized Master subset, extra
structural columns stay blank, and all four chart caches agree with editable
embedded data and current pivot precision. All eight generated slides were
rendered in native PowerPoint and inspected. Table/number-overflow checks pass.
The monthly chart's labels were moved above points after visual inspection.

This is not full pivot/manual parity. Strict coordinate comparisons still find
11 / 12 PIVOT PERFORMANCE differences and 3 / 4 SLIDE GH MEETING differences
for W40 / W41 respectively: existing historical/filter policies, rounded static
values, week-label language and refreshed empty results remain distinguishable.
For example, the manual W40 achievement Grand Total helper is 100%, while its
counts are 239 / 275; the generated formula yields 86.91%.

Browser tests cover Mobile/M2M/both selections, conditional required uploads,
actual multipart field names, zero-row review completion, reset behavior and
desktop/mobile layouts (1440px and 390px) with no page overflow or JS errors.
These browser interaction tests stub job responses; the separate W40 API test
uses the real finalizer and native report writers. Standard tests also cover
legacy Mobile requests, core-only downloads and checksum-verified retries.
A native sparse-week test uses the W41 dataset with a later empty reporting
week and verifies successful workbook/PPT output with zero counts/revenue and
blank lead-time chart points, rather than claiming a live W42 data sample.
