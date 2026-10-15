"""Read-only W38 regression audit for the weekly rule engine."""

from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict, deque
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from .constants import MASTER_COLUMNS, MOBILE_SERVICE_GROUPS
from .dates import parse_datetime
from .io import iter_master_rows, load_holidays
from .models import WeeklyAudit
from .refresh import ExistingSourceIndex, refresh_existing_record
from .review import load_review_decisions
from .rules import (
    build_master_record,
    clean_text,
    normalize_identifier,
    normalize_service_group,
)
from .selection import iter_bq_rows, select_new_source_rows


def _identity(row: dict[str, Any]) -> tuple[str, str, str]:
    fab_upload = parse_datetime(row.get("fab_upload") or row.get("FAB_Upload"))
    fab_key = fab_upload.isoformat(timespec="minutes") if fab_upload else ""
    return (
        normalize_identifier(row.get("quote_num")),
        normalize_identifier(row.get("order number", row.get("order_number"))),
        fab_key,
    )


def _normalized_value(field: str, value: Any) -> Any:
    if isinstance(value, datetime):
        if field == "fab_upload":
            return value.replace(second=0, microsecond=0)
        return value.date()
    if isinstance(value, Decimal):
        return value.quantize(Decimal("0.000001"))
    if isinstance(value, float):
        return Decimal(str(value)).quantize(Decimal("0.000001"))
    if field in {
        "quote_num",
        "order number",
        "ca_name",
        "SALES NAME",
    }:
        return clean_text(value)
    return value


def _field_mismatches(
    generated_rows: list[dict[str, Any]],
    expected_rows: list[dict[str, Any]],
) -> dict[str, int]:
    expected_by_identity: dict[tuple[str, str, str], deque[dict[str, Any]]] = defaultdict(deque)
    for row in expected_rows:
        expected_by_identity[_identity(row)].append(row)

    ignored_fields = {"No."}
    mismatches: Counter[str] = Counter()
    for generated in generated_rows:
        candidates = expected_by_identity[_identity(generated)]
        if not candidates:
            mismatches["<missing expected row>"] += 1
            continue
        expected = candidates.popleft()
        for field in MASTER_COLUMNS:
            if field in ignored_fields:
                continue
            actual_value = _normalized_value(field, generated.get(field))
            expected_value = _normalized_value(field, expected.get(field))
            if actual_value != expected_value:
                mismatches[field] += 1
    return dict(sorted(mismatches.items()))


_W38_DOCUMENTED_REFERENCE_ANOMALIES: dict[str, set[str]] = {
    # These duplicate rows were changed to Cancelled in the manual final without a
    # matching BQ row or Catalyst review value.
    "2-790232122621": {
        "order_status",
        "order_status2",
        "order_status (Confirmed by MIDO Dept.)",
    },
    # The manual workbook used the first quote match (a different order). The rule
    # engine intentionally keeps the exact quote + order match.
    "2-830925854851": {
        "order_completion_date",
        "Week Completion",
        "LEADTIME CASH-COMPLETE",
    },
    # The final manual file changed achievement after the Catalyst checkpoint.
    # A note in ADJUSTMENT is not itself evidence of PM approval.
    "2-829340096158": {"ACHIEVEMENT"},
    "2-829335407089": {"ACHIEVEMENT"},
    # The manual final clears completion on a cancelled order but leaves stale
    # lead time and achievement values from the earlier completed state.
    "2-791214170813": {"LEADTIME CASH-COMPLETE", "ACHIEVEMENT"},
}


def _classify_w38_existing_mismatches(
    generated_rows: list[dict[str, Any]],
    expected_rows: list[dict[str, Any]],
) -> tuple[dict[str, int], dict[str, int]]:
    unexpected: Counter[str] = Counter()
    documented: Counter[str] = Counter()
    for generated, expected in zip(generated_rows, expected_rows, strict=True):
        quote_num = clean_text(expected.get("quote_num"))
        allowed_fields = _W38_DOCUMENTED_REFERENCE_ANOMALIES.get(quote_num, set())
        for field in MASTER_COLUMNS:
            if field == "No.":
                continue
            if _normalized_value(field, generated.get(field)) == _normalized_value(
                field, expected.get(field)
            ):
                continue
            target = documented if field in allowed_fields else unexpected
            target[field] += 1
    return dict(sorted(unexpected.items())), dict(sorted(documented.items()))


def run_reference_audit(
    *,
    bq_csv: str | Path,
    previous_master: str | Path,
    expected_master: str | Path,
    expected_mobile: str | Path,
    holiday_workbook: str | Path,
    catalyst_review: str | Path,
    report_date: date,
) -> WeeklyAudit:
    audit = WeeklyAudit()
    previous_rows = list(iter_master_rows(previous_master))
    expected_rows = list(iter_master_rows(expected_master))
    mobile_rows = list(iter_master_rows(expected_mobile))
    existing_quotes = {
        clean_text(row.get("quote_num"))
        for row in previous_rows
        if clean_text(row.get("quote_num"))
    }
    source_index = ExistingSourceIndex(previous_rows)

    def indexed_source_rows():
        for source_row in iter_bq_rows(bq_csv):
            source_index.add(source_row)
            yield source_row

    selection = select_new_source_rows(
        indexed_source_rows(),
        existing_quote_nums=existing_quotes,
        report_date=report_date,
    )
    expected_appended = expected_rows[len(previous_rows) :]
    holidays = load_holidays(holiday_workbook)
    generated_rows = [
        build_master_record(
            row,
            row_number=len(previous_rows) + offset,
            holidays=holidays,
        )
        for offset, row in enumerate(selection.selected_rows, start=1)
    ]
    existing_refresh_issues = []
    draft_existing_rows: list[dict[str, Any]] = []
    sources: list[dict[str, Any] | None] = []
    for previous_row in previous_rows:
        source, issue = source_index.resolve(previous_row)
        sources.append(source)
        if issue is not None:
            existing_refresh_issues.append(issue)
        draft_existing_rows.append(refresh_existing_record(
            previous_row,
            source=source,
            catalyst_override=None,
            holidays=holidays,
        ))
    review_result = load_review_decisions(
        catalyst_review,
        draft_existing_rows + generated_rows,
        first_new_row=len(previous_rows) + 2,
    )
    refreshed_existing_rows: list[dict[str, Any]] = []
    for row_number, (previous_row, draft_row, source) in enumerate(
        zip(previous_rows, draft_existing_rows, sources, strict=True), start=2
    ):
        decision = review_result.decisions.get(row_number)
        refreshed_existing_rows.append(
            refresh_existing_record(
                previous_row,
                source=source,
                catalyst_override=None,
                review_decision=decision,
                holidays=holidays,
            )
            if decision is not None else draft_row
        )

    audit.check("previous_master_rows", len(previous_rows), 52_820)
    audit.check("selected_source_rows", len(selection.selected_rows), 772)
    audit.check("expected_appended_rows", len(expected_appended), 772)
    audit.check("expected_master_rows", len(expected_rows), 53_592)
    audit.check("expected_mobile_rows", len(mobile_rows), 53_108)

    source_identities = Counter(_identity(row) for row in selection.selected_rows)
    expected_identities = Counter(_identity(row) for row in expected_appended)
    audit.check(
        "selected_source_identity_multiset",
        source_identities == expected_identities,
        True,
    )
    audit.check(
        "generated_new_row_field_mismatches",
        _field_mismatches(generated_rows, expected_appended),
        {},
    )
    unexpected_existing, documented_existing = _classify_w38_existing_mismatches(
        refreshed_existing_rows,
        expected_rows[: len(previous_rows)],
    )
    audit.check("unexpected_existing_field_mismatches", unexpected_existing, {})
    audit.check(
        "documented_w38_reference_anomalies",
        documented_existing,
        {
            "ACHIEVEMENT": 3,
            "LEADTIME CASH-COMPLETE": 2,
            "Week Completion": 1,
            "order_completion_date": 1,
            "order_status": 2,
            "order_status (Confirmed by MIDO Dept.)": 2,
            "order_status2": 2,
        },
    )
    audit.check("catalyst_unmatched_new_rows", review_result.unmatched_new_rows, 0)
    audit.check(
        "source_pair_conflict_issues",
        sum(issue.code == "conflicting_source_pair" for issue in existing_refresh_issues),
        0,
    )

    invalid_mobile_services = sorted(
        {
            normalize_service_group(row.get("service_group"))
            for row in mobile_rows
            if normalize_service_group(row.get("service_group")) not in MOBILE_SERVICE_GROUPS
        }
    )
    audit.check("mobile_out_of_scope_services", invalid_mobile_services, [])
    audit.check("selection_parse_issues", len(selection.issues), 0)
    audit.check(
        "selected_rows_after_cutoff",
        sum(
            1
            for row in selection.selected_rows
            if parse_datetime(row.get("FAB_Upload")).date() > report_date
        ),
        0,
    )
    return audit


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bq-csv", required=True, type=Path)
    parser.add_argument("--previous-master", required=True, type=Path)
    parser.add_argument("--expected-master", required=True, type=Path)
    parser.add_argument("--expected-mobile", required=True, type=Path)
    parser.add_argument("--holiday-workbook", required=True, type=Path)
    parser.add_argument("--catalyst-review", required=True, type=Path)
    parser.add_argument("--report-date", required=True, type=date.fromisoformat)
    parser.add_argument("--json-output", type=Path)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    audit = run_reference_audit(
        bq_csv=args.bq_csv,
        previous_master=args.previous_master,
        expected_master=args.expected_master,
        expected_mobile=args.expected_mobile,
        holiday_workbook=args.holiday_workbook,
        catalyst_review=args.catalyst_review,
        report_date=args.report_date,
    )
    payload = json.dumps(audit.to_dict(), indent=2, default=str)
    print(payload)
    if args.json_output:
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(payload + "\n", encoding="utf-8")
    return 0 if audit.passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
