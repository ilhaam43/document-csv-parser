# Weekly SECM Report Development

## Current scope

The command-line workflow can prepare a Catalyst review package, finalize the
Master and Mobile datasets, and write both Excel workbooks with refreshed native
PivotTables. It uses the prior week's workbooks as templates and requires Windows,
Microsoft Excel, and pywin32 for the workbook stage. The original templates are
copied; they are never edited in place.

The Catalyst check is a human checkpoint until dashboard access is available.
The Weekly web page and API implement that checkpoint. The Mobile PowerPoint
generator uses the prior week's PPTX as an editable template; it leaves the
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
format, previous Mobile workbook, previous Mobile weekly PPT, and holiday workbook. The page preserves the
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

Finalization creates a ZIP with Master Excel, Mobile Excel, Mobile weekly PPT,
`audit.json`,
`exceptions.csv`, and `review-decisions.csv`.
Uploaded slide, pivot, and approved-correction CSVs are included when supplied,
so their decisions remain traceable with the generated files.
An uploaded presentation chart review is also included when supplied.
The audit records remaining unresolved Catalyst statuses and the web shows a
warning if any remain; a generated report is not proof that all reviews were
completed.
If a Catalyst review conflicts or Excel fails, the review can be corrected and
submitted again without uploading the source files again. Preparation and final
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
  calculation policy, not a global replacement of zeros. Historical reviewed
  values are retained when their supporting dates have not changed.
- Negative lead time remains visible as an exception.
- Achievement limits: Mobile 3 days, Bulk Service 5 days, M2M 19 days.
- PM approval must be explicitly recorded in the review CSV to override the
  calculated achievement. A reviewed workbook can also explicitly change
  `NOT ACHIEVED` to `ACHIEVED`; an `ADJUSTMENT` note alone is not approval.
  Historical approved values are carried from the prior master when the
  relevant dates do not change.
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
- The existing Mobile adjustment-inclusion pivot group includes all current
  adjustment notes, not only notes that changed since the template. This matches
  all three manual Mobile samples. It changes aggregation membership, not
  ACHIEVEMENT or PM approval. Other curated pivot filters remain unchanged.
- FAB timestamps preserve available time precision. Lifecycle report dates
  (`cash_dt`, `order created`, `order_completion_date`) are date-only values.
- Missing completion dates clear stale completion counts, periods and calculated
  cash-completion SLA values, even when the date was already blank last week.
  An existing approved ACHIEVED value accompanied by ADJUSTMENT is preserved;
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
method for every conflict to `bq-source-decisions.csv` in the output ZIP.
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
even with an October export). Mobile weekly values are rounded as in the
existing template convention; monthly averages retain their precision.
All chart series are resized to the resulting range. Helper percentage formulas
are restored relative to each refreshed pivot's data and grand-total rows.
At a detected month rollover, Master nonnegative numeric FAB-CASH filter values
are included even when a prior template hid them; negative/blank metrics remain
excluded from that filtered average. Within the same reporting month the template
selection is preserved. This rollover policy is inferred from the supplied
September/October samples, not an independently confirmed owner instruction.
Mobile templates retain their distinct, curated numeric selections;
they are not normalized to the Master filter. Other manually decided values belong in
`slide-review.csv` with columns
`workbook,cell,value,reviewed_by,review_reason`. For W38, Master cell `U21`
needs the manually confirmed total of 16; do not infer it from the pivot.
Review the chart values before distribution. A W38 reference fixture copied
from the manual file is useful for regression, but is not production approval.
The W39 reference revises W2-W4 to unrounded values, unlike W38. Use an
attributed slide review for those values; the two examples do not establish a
general rounding rule for future weeks.

W39's manual Master also changes two `ADJUSTMENT` pivot filters compared with
W38, including four August orders. This change is not inferable from BQ data.
Leave the previous week's filter state in place unless the report owner confirms
the new selection. To reproduce the reviewed W39 file, upload an attributed
`pivot-review.csv` alongside the Catalyst review, or pass `--pivot-review` to
`weekly.workbook`. Its columns are
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

## Write Mobile PowerPoint

### October snapshot validation

The W40 retry with the original 8 October BQ snapshot confirmed 97 new rows,
54,959 Master rows and 54,464 Mobile rows. With only 67 Catalyst status decisions
replayed as test inputs, all numerical cells in the four pivot/monitoring sheets
match the original manual W40 Master and Mobile files, and all six Mobile PPT
chart series match numerically. This does not make the entire workbooks identical:
one valid FAB-CASH value is blank in manual, and SLIDE GH MEETING retains a
spelling difference and two Mobile static-source rounding/copy differences.
Those differences remain visible in the comparison results.

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
removed. When new values exceed a fixed template axis, that axis expands with
headroom so the point is visible. Explicit presentation reviews still override
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

## Run the W38 golden regression

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

The audit checks the selected row count, identity multiset, every generated field,
master and mobile row counts, and mobile service scope. With only the supplied
pre-Catalyst checkpoint, it reports 64 unresolved status differences from the
manual final, while the seven other documented post-checkpoint differences are
classified separately. The generator does not silently infer missing Catalyst
decisions to make the regression pass.

The same regression is available through `tests/test_weekly_reference_w38.py`:

```powershell
$env:WEEKLY_REFERENCE_DIR = "D:\SECM\Weekly 38"
python -m unittest tests.test_weekly_reference_w38 -v
```

## Next implementation stages

1. Add the M2M output path after its business rules and manual example are reviewed.
2. Add job retention and startup recovery after server deployment requirements are set.
