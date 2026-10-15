"""Create a reviewable weekly master preparation package without editing Excel."""

from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path
from typing import Any

from .catalyst import write_catalyst_review_csv
from .dates import has_explicit_seconds, parse_report_month
from .io import iter_master_rows, load_holidays, write_issues_csv, write_master_csv
from .matching import record_key
from .models import ValidationIssue
from .refresh import ExistingSourceIndex, refresh_existing_record
from .rules import build_master_record, clean_text, is_unresolved_order_status
from .selection import iter_bq_rows, select_new_source_rows
from .slide import write_slide_review_template
from .source_review import (
    automatic_source_choice, collect_source_conflicts,
    source_with_reviewed_lifecycle, write_source_review_csv,
)


def prepare_weekly_inputs(
    *,
    bq_csv: str | Path,
    previous_master: str | Path,
    holiday_workbook: str | Path,
    report_date: date,
    output_dir: str | Path,
    report_month: date | None = None,
) -> dict[str, Any]:
    """Build selected new master rows plus machine-readable audit artifacts."""
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
    holidays = load_holidays(holiday_workbook)
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
    missing_seconds = sum(
        not has_explicit_seconds(row.get("FAB_Upload")) for row in selection.selected_rows
    )

    records: list[dict[str, Any]] = []
    issues = list(selection.issues)
    first_selected: dict[str, dict[str, Any]] = {}
    for offset, source in enumerate(selection.selected_rows, start=1):
        quote_source = first_selected.setdefault(clean_text(source.get("quote_num")), source)
        source = source_with_reviewed_lifecycle(
            source, quote_source, automatic_source_choice(source.get("order_status"))
        )
        try:
            record = build_master_record(
                source,
                row_number=len(previous_rows) + offset,
                holidays=holidays,
            )
        except ValueError as exc:
            issues.append(
                ValidationIssue(
                    code="master_record_build_failed",
                    message=str(exc),
                    row_number=offset + 1,
                )
            )
            continue
        if (
            isinstance(record["LEADTIME FAB-CASH"], int)
            and record["LEADTIME FAB-CASH"] < 0
        ) or (
            isinstance(record["LEADTIME CASH-COMPLETE"], int)
            and record["LEADTIME CASH-COMPLETE"] < 0
        ):
            issues.append(
                ValidationIssue(
                    code="negative_leadtime",
                    message="Generated row contains a negative lead time and needs review.",
                    key=record_key(record.get("quote_num"), record.get("order number")),
                    row_number=len(previous_rows) + offset + 1,
                )
            )
        records.append(record)

    refreshed_rows: list[dict[str, Any]] = []
    for row_number, previous_row in enumerate(previous_rows, start=2):
        source, source_issue = source_index.resolve(previous_row)
        quote_source = source_index.first_for_quote(previous_row.get("quote_num"))
        if source is not None and quote_source is not None:
            source = source_with_reviewed_lifecycle(
                source, quote_source, automatic_source_choice(previous_row.get("order_status"))
            )
        if source_issue is not None:
            issues.append(ValidationIssue(
                code=source_issue.code,
                message=source_issue.message,
                key=source_issue.key,
                row_number=row_number,
            ))
        refreshed_rows.append(refresh_existing_record(
            previous_row,
            source=source,
            catalyst_override=None,
            holidays=holidays,
        ))
    draft_rows = refreshed_rows + records
    for number, row in enumerate(draft_rows, start=1):
        row["No."] = number
    pending_rows = {
        row_number
        for row_number, row in enumerate(draft_rows, start=2)
        if is_unresolved_order_status(row.get("order_status"))
    }

    records_path = output_path / "new-master-rows.csv"
    draft_path = output_path / "draft-master-data.csv"
    review_path = output_path / "catalyst-review.csv"
    source_review_path = output_path / "source-review.csv"
    slide_review_path = output_path / "slide-review.csv"
    issues_path = output_path / "exceptions.csv"
    audit_path = output_path / "audit.json"
    write_master_csv(records_path, records)
    write_master_csv(draft_path, draft_rows)
    write_catalyst_review_csv(review_path, draft_rows, row_numbers=pending_rows)
    write_source_review_csv(source_review_path, source_conflicts)
    write_slide_review_template(slide_review_path)
    write_issues_csv(issues_path, issues)

    audit = {
        "report_date": report_date.isoformat(),
        "report_month": (report_month or report_date).strftime("%Y-%m"),
        "previous_master_rows": len(previous_rows),
        "existing_quote_count": len(existing_quotes),
        "selected_source_rows": len(selection.selected_rows),
        "new_rows_without_fab_seconds": missing_seconds,
        "generated_master_rows": len(records),
        "draft_master_rows": len(draft_rows),
        "pending_catalyst_rows": len(pending_rows),
        "source_conflict_rows": len(source_conflicts),
        "issue_count": len(issues),
        "rejected_reasons": dict(sorted(selection.rejected_reasons.items())),
        "outputs": {
            "new_master_rows": str(records_path.resolve()),
            "draft_master_data": str(draft_path.resolve()),
            "catalyst_review": str(review_path.resolve()),
            "source_review": str(source_review_path.resolve()),
            "slide_review": str(slide_review_path.resolve()),
            "exceptions": str(issues_path.resolve()),
            "audit": str(audit_path.resolve()),
        },
    }
    audit_path.write_text(json.dumps(audit, indent=2, default=str) + "\n", encoding="utf-8")
    return audit


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bq-csv", required=True, type=Path)
    parser.add_argument("--previous-master", required=True, type=Path)
    parser.add_argument("--holiday-workbook", required=True, type=Path)
    parser.add_argument("--report-date", required=True, type=date.fromisoformat)
    parser.add_argument("--report-month", type=parse_report_month)
    parser.add_argument("--output-dir", required=True, type=Path)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    audit = prepare_weekly_inputs(
        bq_csv=args.bq_csv,
        previous_master=args.previous_master,
        holiday_workbook=args.holiday_workbook,
        report_date=args.report_date,
        report_month=args.report_month,
        output_dir=args.output_dir,
    )
    print(json.dumps(audit, indent=2, default=str))
    return 0 if audit["generated_master_rows"] == audit["selected_source_rows"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
