"""Build final weekly master and Mobile datasets after Catalyst review."""

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter
from datetime import date
from pathlib import Path
from typing import Any

from .constants import MASTER_SERVICE_GROUPS, MOBILE_SERVICE_GROUPS, M2M_SERVICE_GROUPS
from .dates import has_explicit_seconds, parse_report_month
from .io import (
    iter_master_rows,
    load_holidays,
    write_issues_csv,
    write_master_csv,
)
from .matching import record_key
from .models import ReviewDecision, ValidationIssue
from .refresh import ExistingSourceIndex, refresh_existing_record
from .review import load_review_decisions
from .rules import build_master_record, clean_text, is_unresolved_order_status, normalize_service_group
from .rules import inclusive_networkdays
from .selection import iter_bq_rows, select_new_source_rows
from .source_review import (
    SOURCE_REVIEW_COLUMNS,
    automatic_source_choice,
    collect_source_conflicts,
    load_source_review_csv,
    source_with_reviewed_lifecycle,
)


def _issue_with_row(issue: ValidationIssue, row_number: int) -> ValidationIssue:
    if issue.row_number is not None:
        return issue
    return ValidationIssue(
        code=issue.code,
        message=issue.message,
        key=issue.key,
        row_number=row_number,
    )


def finalize_weekly_data(
    *,
    bq_csv: str | Path,
    previous_master: str | Path,
    catalyst_review: str | Path,
    source_review: str | Path | None = None,
    manual_decisions: str | Path | None = None,
    holiday_workbook: str | Path,
    report_date: date,
    output_dir: str | Path,
    report_month: date | None = None,
) -> dict[str, Any]:
    output_path = Path(output_dir)
    if report_month is not None and report_month > report_date:
        raise ValueError("report_month cannot be after report_date.")
    output_path.mkdir(parents=True, exist_ok=True)
    previous_rows = list(iter_master_rows(previous_master))
    existing_quotes = {
        clean_text(row.get("quote_num"))
        for row in previous_rows
        if clean_text(row.get("quote_num"))
    }
    source_index = ExistingSourceIndex(previous_rows)

    def indexed_source_rows():
        for source in iter_bq_rows(bq_csv):
            source_index.add(source)
            yield source

    selection = select_new_source_rows(
        indexed_source_rows(),
        existing_quote_nums=existing_quotes,
        report_date=report_date,
        report_month=report_month,
    )
    source_conflicts = collect_source_conflicts(
        previous_rows, source_index, selection.selected_rows
    )
    source_choices = (
        load_source_review_csv(source_review, source_conflicts)
        if source_review is not None else {}
    )
    first_selected: dict[str, dict[str, Any]] = {}
    for source in selection.selected_rows:
        first_selected.setdefault(clean_text(source.get("quote_num")), source)
    holidays = load_holidays(holiday_workbook)
    appended_rows: list[dict[str, Any]] = []
    effective_source_choices: dict[int, str] = {}
    for offset, source in enumerate(selection.selected_rows, start=1):
        record_number = len(previous_rows) + offset
        choice = source_choices.get(record_number + 1, automatic_source_choice(source.get("order_status")))
        effective_source_choices[record_number + 1] = choice
        if choice == "quote":
            source = source_with_reviewed_lifecycle(
                source, first_selected[clean_text(source.get("quote_num"))], "quote"
            )
        record = build_master_record(
            source,
            row_number=record_number,
            holidays=holidays,
        )
        if any(
            isinstance(record[field], int) and record[field] < 0
            for field in ("LEADTIME FAB-CASH", "LEADTIME CASH-COMPLETE")
        ):
            selection.issues.append(
                ValidationIssue(
                    code="negative_leadtime",
                    message="Generated row contains a negative lead time and needs review.",
                    key=record_key(record.get("quote_num"), record.get("order number")),
                    row_number=record_number + 1,
                )
            )
        appended_rows.append(record)

    issues = list(selection.issues)
    sources_by_row: dict[int, dict[str, Any] | None] = {}
    metric_decisions_by_row: dict[int, list[dict[str, Any]]] = {}
    draft_existing_rows: list[dict[str, Any]] = []
    for row_number, previous_row in enumerate(previous_rows, start=2):
        source, source_issue = source_index.resolve(previous_row)
        choice = source_choices.get(row_number, automatic_source_choice(previous_row.get("order_status")))
        effective_source_choices[row_number] = choice
        quote_choice = source is not None and choice == "quote"
        if quote_choice:
            source = source_with_reviewed_lifecycle(
                source, source_index.first_for_quote(previous_row.get("quote_num")), "quote"
            )
        sources_by_row[row_number] = source
        if source_issue is not None:
            issues.append(_issue_with_row(source_issue, row_number))
        metric_decisions_by_row[row_number] = []
        draft_existing_rows.append(refresh_existing_record(
            previous_row,
            source=source,
            catalyst_override=None,
            review_decision=(
                ReviewDecision(order_status=clean_text(source.get("order_status")) or "In Progress")
                if quote_choice else None
            ),
            holidays=holidays,
            metric_decisions=metric_decisions_by_row[row_number],
        ))
    draft_rows = draft_existing_rows + appended_rows
    catalyst_review_result = load_review_decisions(
        catalyst_review,
        draft_rows,
        first_new_row=len(previous_rows) + 2,
    )
    if (
        Path(catalyst_review).suffix.casefold() == ".xlsx"
        and catalyst_review_result.unmatched_new_rows
    ):
        raise ValueError(
            "Catalyst checkpoint cannot uniquely match "
            f"{catalyst_review_result.unmatched_new_rows} new rows."
        )
    catalyst_decisions = catalyst_review_result.decisions
    catalyst_count = catalyst_review_result.count
    unverified_fab_seconds = sum(
        not has_explicit_seconds(source.get("FAB_Upload"))
        and len(previous_rows) + offset + 1 not in catalyst_review_result.checkpoint_rows
        for offset, source in enumerate(selection.selected_rows, start=1)
    )
    reviewed_zero_leadtimes = 0
    for row_number, reviewed in catalyst_review_result.checkpoint_rows.items():
        row = appended_rows[row_number - len(previous_rows) - 2]
        row["fab_upload"] = reviewed.fab_upload
        if (
            reviewed.leadtime_fab_cash is not None
            and reviewed.leadtime_fab_cash != row["LEADTIME FAB-CASH"]
        ):
            raw_days = inclusive_networkdays(
                row["fab_upload"], row["cash_dt"], holidays, normalize_zero=False
            )
            if reviewed.leadtime_fab_cash != 0 or raw_days != 0 or row["LEADTIME FAB-CASH"] != 1:
                raise ValueError(f"Checkpoint FAB-CASH lead time conflicts with source at row {row_number}.")
            row["LEADTIME FAB-CASH"] = 0
            reviewed_zero_leadtimes += 1
    manual_review_decisions: dict[int, ReviewDecision] = {}
    if manual_decisions is not None:
        if Path(manual_decisions).suffix.casefold() != ".csv":
            raise ValueError("Manual decisions must be a reviewed CSV file.")
        manual_review_result = load_review_decisions(
            manual_decisions, draft_rows, require_attribution=True
        )
        manual_review_decisions = manual_review_result.decisions
    decisions = dict(catalyst_decisions)
    for row_number, manual in manual_review_decisions.items():
        base = catalyst_decisions.get(row_number, ReviewDecision())
        decisions[row_number] = ReviewDecision(
            order_status=manual.order_status or base.order_status,
            completion_date=manual.completion_date or base.completion_date,
            leadtime_fab_cash=(
                manual.leadtime_fab_cash
                if manual.leadtime_fab_cash is not None
                else base.leadtime_fab_cash
            ),
            pm_approved=manual.pm_approved if manual.pm_approved is not None else base.pm_approved,
            reviewed_by=manual.reviewed_by,
            review_reason=manual.review_reason,
        )
    refreshed_rows: list[dict[str, Any]] = []
    for row_number, previous_row in enumerate(previous_rows, start=2):
        decision = decisions.get(row_number)
        if decision is not None:
            metric_decisions_by_row[row_number] = []
        refreshed_rows.append(
            refresh_existing_record(
                previous_row,
                source=sources_by_row[row_number],
                catalyst_override=None,
                review_decision=decision,
                holidays=holidays,
                metric_decisions=metric_decisions_by_row[row_number],
            )
            if decision is not None else draft_existing_rows[row_number - 2]
        )

    for offset, record in enumerate(appended_rows, start=len(previous_rows) + 2):
        decision = decisions.get(offset)
        metric_decisions_by_row[offset] = []
        if decision is not None:
            appended_rows[offset - len(previous_rows) - 2] = refresh_existing_record(
                record,
                source=None,
                catalyst_override=None,
                review_decision=decision,
                holidays=holidays,
                metric_decisions=metric_decisions_by_row[offset],
            )
            for metric in metric_decisions_by_row[offset]:
                if metric["decision_method"] == "published_history":
                    metric["decision_method"] = "calculated_new"
        else:
            for field in ("LEADTIME FAB-CASH", "LEADTIME CASH-COMPLETE", "ACHIEVEMENT"):
                metric_decisions_by_row[offset].append({
                    "metric": field, "previous_value": None,
                    "calculated_value": record.get(field), "final_value": record.get(field),
                    "decision_method": "calculated_new",
                })

    master_rows = refreshed_rows + appended_rows
    for number, row in enumerate(master_rows, start=1):
        row["No."] = number
    mobile_rows = [
        dict(row)
        for row in master_rows
        if normalize_service_group(row.get("service_group")) in MOBILE_SERVICE_GROUPS
    ]
    for record_number, row in enumerate(mobile_rows, start=1):
        row["No."] = record_number
    m2m_rows = [dict(row) for row in master_rows
                if normalize_service_group(row.get("service_group")) in M2M_SERVICE_GROUPS]
    for record_number, row in enumerate(m2m_rows, start=1):
        row["No."] = record_number

    master_path = write_master_csv(output_path / "master-final-data.csv", master_rows)
    mobile_path = write_master_csv(output_path / "mobile-data.csv", mobile_rows)
    m2m_path = write_master_csv(output_path / "m2m-data.csv", m2m_rows)
    issues_path = write_issues_csv(output_path / "exceptions.csv", issues)
    source_decisions_path = output_path / "bq-source-decisions.csv"
    with source_decisions_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle, fieldnames=(*SOURCE_REVIEW_COLUMNS, "decision_method")
        )
        writer.writeheader()
        for conflict in source_conflicts:
            row_number = int(conflict["master_row"])
            choice = effective_source_choices[row_number]
            writer.writerow({
                **conflict,
                "reviewed_source": choice,
                "decision_method": (
                    "explicit_override" if row_number in source_choices else
                    "automatic_active_quote" if choice == "quote" else "automatic_terminal_order"
                ),
            })
    decisions_path = output_path / "review-decisions.csv"
    with decisions_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=(
                "master_row", "quote_num", "order_number", "reviewed_status",
                "reviewed_completion_date", "reviewed_fab_cash_leadtime",
                "pm_approved", "reviewed_by", "review_reason",
            ),
        )
        writer.writeheader()
        for row_number, decision in sorted(manual_review_decisions.items() if manual_decisions else decisions.items()):
            row = master_rows[row_number - 2]
            writer.writerow({
                "master_row": row_number,
                "quote_num": row.get("quote_num"),
                "order_number": row.get("order number"),
                "reviewed_status": decision.order_status or "",
                "reviewed_completion_date": (
                    decision.completion_date.date().isoformat() if decision.completion_date else ""
                ),
                "reviewed_fab_cash_leadtime": (
                    decision.leadtime_fab_cash if decision.leadtime_fab_cash is not None else ""
                ),
                "pm_approved": (
                    "yes" if decision.pm_approved is True else "no" if decision.pm_approved is False else ""
                ),
                "reviewed_by": decision.reviewed_by,
                "review_reason": decision.review_reason,
            })
    metric_decisions_path = output_path / "metric-decisions.csv"
    metric_methods: Counter[str] = Counter()
    metric_variances: Counter[str] = Counter()
    with metric_decisions_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=(
            "master_row", "quote_num", "order_number", "metric", "previous_value",
            "calculated_value", "final_value", "decision_method",
        ))
        writer.writeheader()
        for row_number, metrics in metric_decisions_by_row.items():
            row = master_rows[row_number - 2]
            for metric in metrics:
                writer.writerow({
                    "master_row": row_number, "quote_num": row.get("quote_num"),
                    "order_number": row.get("order number"), **metric,
                })
                metric_methods[metric["decision_method"]] += 1
                if metric["calculated_value"] != metric["final_value"]:
                    metric_variances[metric["metric"]] += 1
    audit_path = output_path / "audit.json"
    issue_counts = Counter(issue.code for issue in issues)
    invalid_master_services = sorted(
        {
            normalize_service_group(row.get("service_group"))
            for row in master_rows
            if normalize_service_group(row.get("service_group")) not in MASTER_SERVICE_GROUPS
        }
    )
    audit = {
        "report_date": report_date.isoformat(),
        "report_month": (report_month or report_date).strftime("%Y-%m"),
        "previous_master_rows": len(previous_rows),
        "selected_new_rows": len(selection.selected_rows),
        "new_rows_without_fab_seconds": unverified_fab_seconds,
        "master_final_rows": len(master_rows),
        "mobile_rows": len(mobile_rows),
        "m2m_rows": len(m2m_rows),
        "catalyst_decisions": catalyst_count,
        "checkpoint_matched_new_rows": len(catalyst_review_result.checkpoint_rows),
        "checkpoint_unmatched_new_rows": catalyst_review_result.unmatched_new_rows,
        "checkpoint_zero_leadtime_reviews": reviewed_zero_leadtimes,
        "manual_decisions": len(manual_review_decisions),
        "metric_policy": "preserve_published_history_v1",
        "metric_decision_methods": dict(sorted(metric_methods.items())),
        "metric_calculation_variances": dict(sorted(metric_variances.items())),
        "holiday_dates": sorted(day.isoformat() for day in holidays),
        "source_conflict_rows": len(source_conflicts),
        "source_reviewed_rows": len(source_choices),
        "source_auto_order_rows": sum(
            effective_source_choices[int(row["master_row"])] == "order"
            and int(row["master_row"]) not in source_choices for row in source_conflicts
        ),
        "source_auto_quote_rows": sum(
            effective_source_choices[int(row["master_row"])] == "quote"
            and int(row["master_row"]) not in source_choices for row in source_conflicts
        ),
        "source_quote_choices": sum(
            effective_source_choices[int(row["master_row"])] == "quote" for row in source_conflicts
        ),
        "pending_catalyst_rows": sum(
            is_unresolved_order_status(row.get("order_status")) for row in master_rows
        ),
        "invalid_master_services": invalid_master_services,
        "issue_counts": dict(sorted(issue_counts.items())),
        "rejected_source_reasons": dict(sorted(selection.rejected_reasons.items())),
        "outputs": {
            "master_final_data": str(master_path.resolve()),
            "mobile_data": str(mobile_path.resolve()),
            "m2m_data": str(m2m_path.resolve()),
            "exceptions": str(issues_path.resolve()),
            "bq_source_decisions": str(source_decisions_path.resolve()),
            "review_decisions": str(decisions_path.resolve()),
            "metric_decisions": str(metric_decisions_path.resolve()),
            "audit": str(audit_path.resolve()),
        },
    }
    audit_path.write_text(json.dumps(audit, indent=2, default=str) + "\n", encoding="utf-8")
    return audit


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bq-csv", required=True, type=Path)
    parser.add_argument("--previous-master", required=True, type=Path)
    parser.add_argument("--catalyst-review", required=True, type=Path)
    parser.add_argument("--manual-decisions", type=Path)
    parser.add_argument("--source-review", type=Path)
    parser.add_argument("--holiday-workbook", required=True, type=Path)
    parser.add_argument("--report-date", required=True, type=date.fromisoformat)
    parser.add_argument("--report-month", type=parse_report_month)
    parser.add_argument("--output-dir", required=True, type=Path)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    audit = finalize_weekly_data(
        bq_csv=args.bq_csv,
        previous_master=args.previous_master,
        catalyst_review=args.catalyst_review,
        manual_decisions=args.manual_decisions,
        source_review=args.source_review,
        holiday_workbook=args.holiday_workbook,
        report_date=args.report_date,
        report_month=args.report_month,
        output_dir=args.output_dir,
    )
    print(json.dumps(audit, indent=2, default=str))
    return 0 if not audit["invalid_master_services"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
