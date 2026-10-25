"""Audit any weekly sample using the production finalizer, without parity exemptions."""

from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict, deque
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

from .constants import MASTER_COLUMNS, MOBILE_SERVICE_GROUPS
from .dates import parse_datetime, parse_report_month
from .finalize import finalize_weekly_data
from .io import iter_master_csv, iter_master_rows, load_holidays
from .models import WeeklyAudit
from .rules import inclusive_networkdays, normalize_identifier, normalize_service_group, to_decimal


def _identity(row: dict[str, Any]) -> tuple[str, str, str]:
    fab = parse_datetime(row.get("fab_upload"))
    return (
        normalize_identifier(row.get("quote_num")),
        normalize_identifier(row.get("order number")),
        fab.isoformat(timespec="minutes") if fab else "",
    )


def _normalized_value(field: str, value: Any) -> Any:
    if isinstance(value, datetime):
        return value.replace(second=0, microsecond=0) if field == "fab_upload" else value.date()
    if isinstance(value, (float, Decimal)):
        return Decimal(str(value))
    return value


def _values_equal(field: str, actual: Any, expected: Any) -> bool:
    actual = _normalized_value(field, actual)
    expected = _normalized_value(field, expected)
    if isinstance(actual, bool) or isinstance(expected, bool):
        return type(actual) is type(expected) and actual == expected
    if isinstance(actual, (int, Decimal)) and isinstance(expected, (int, Decimal)):
        return abs(Decimal(actual) - Decimal(expected)) <= Decimal("0.00000001")
    return actual == expected


def compare_master_records(
    generated_rows: list[dict[str, Any]],
    expected_rows: list[dict[str, Any]],
    *,
    example_limit: int = 20,
) -> dict[str, Any]:
    """Match duplicate identities by occurrence; never ignore an unmatched row."""
    expected_by_identity: dict[tuple[str, str, str], deque] = defaultdict(deque)
    for number, row in enumerate(expected_rows, start=2):
        expected_by_identity[_identity(row)].append((number, row))
    counts: Counter[str] = Counter()
    examples: list[dict[str, Any]] = []
    for number, actual in enumerate(generated_rows, start=2):
        candidates = expected_by_identity[_identity(actual)]
        if not candidates:
            counts["<extra generated row>"] += 1
            if len(examples) < example_limit:
                examples.append({"generated_row": number, "identity": _identity(actual), "field": "<extra generated row>"})
            continue
        expected_number, expected = candidates.popleft()
        for field in MASTER_COLUMNS:
            if field == "No." or _values_equal(field, actual.get(field), expected.get(field)):
                continue
            counts[field] += 1
            if len(examples) < example_limit:
                examples.append({
                    "generated_row": number, "reference_row": expected_number,
                    "quote_num": actual.get("quote_num"), "order_number": actual.get("order number"),
                    "field": field, "generated": actual.get(field), "reference": expected.get(field),
                })
    for candidates in expected_by_identity.values():
        for number, expected in candidates:
            counts["<missing generated row>"] += 1
            if len(examples) < example_limit:
                examples.append({"reference_row": number, "identity": _identity(expected), "field": "<missing generated row>"})
    return {
        "matches": not counts,
        "generated_rows": len(generated_rows), "reference_rows": len(expected_rows),
        "differences_by_field": dict(sorted(counts.items())), "examples": examples,
    }


def reference_formula_mismatches(
    records: list[dict[str, Any]], holidays: set[date]
) -> dict[str, int]:
    """Diagnose the reference itself, without changing either comparison input."""
    counts: Counter[str] = Counter()
    for row in records:
        total = to_decimal(row.get("otc")) + to_decimal(row.get("mrc"))
        calculated = {
            "LEADTIME FAB-CASH": inclusive_networkdays(row.get("fab_upload"), row.get("cash_dt"), holidays),
            "OTC+MRC": total,
            "OTC+MRC (inMio)": total / Decimal("1000000"),
        }
        for field, value in calculated.items():
            if not _values_equal(field, value, row.get(field)):
                counts[field] += 1
    return dict(sorted(counts.items()))


def run_reference_audit(
    *,
    bq_csv: str | Path,
    previous_master: str | Path,
    expected_master: str | Path,
    expected_mobile: str | Path,
    holiday_workbook: str | Path,
    catalyst_review: str | Path,
    report_date: date,
    report_month: date | None = None,
    manual_decisions: str | Path | None = None,
    source_review: str | Path | None = None,
) -> WeeklyAudit:
    """Replay real inputs. Expected workbooks are read only after finalization."""
    with TemporaryDirectory(prefix="weekly-reference-") as directory:
        result = finalize_weekly_data(
            bq_csv=bq_csv, previous_master=previous_master,
            catalyst_review=catalyst_review, holiday_workbook=holiday_workbook,
            report_date=report_date, report_month=report_month,
            manual_decisions=manual_decisions, source_review=source_review,
            output_dir=directory,
        )
        generated_master = list(iter_master_csv(Path(directory) / "master-final-data.csv"))
        generated_mobile = list(iter_master_csv(Path(directory) / "mobile-data.csv"))
    audit = WeeklyAudit()
    holidays = load_holidays(holiday_workbook)
    for role, generated, reference in (
        ("master", generated_master, expected_master),
        ("mobile", generated_mobile, expected_mobile),
    ):
        expected = list(iter_master_rows(reference))
        audit.check(f"{role}_row_count", len(generated), len(expected))
        comparison = compare_master_records(generated, expected)
        audit.check(f"{role}_field_mismatches", comparison["differences_by_field"], {})
        audit.check(
            f"{role}_reference_formula_mismatches",
            reference_formula_mismatches(expected, holidays), {},
        )
    audit.check("pending_catalyst_rows", result["pending_catalyst_rows"], 0)
    audit.check("invalid_master_services", result["invalid_master_services"], [])
    audit.check(
        "source_pair_conflict_issues", result["issue_counts"].get("conflicting_source_pair", 0), 0
    )
    audit.check("mobile_out_of_scope_services", sorted({
        normalize_service_group(row.get("service_group")) for row in generated_mobile
        if normalize_service_group(row.get("service_group")) not in MOBILE_SERVICE_GROUPS
    }), [])
    return audit


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("bq-csv", "previous-master", "expected-master", "expected-mobile", "holiday-workbook", "catalyst-review"):
        parser.add_argument(f"--{name}", required=True, type=Path)
    parser.add_argument("--report-date", required=True, type=date.fromisoformat)
    parser.add_argument("--report-month", type=parse_report_month)
    parser.add_argument("--manual-decisions", type=Path)
    parser.add_argument("--source-review", type=Path)
    parser.add_argument("--json-output", type=Path)
    return parser


def main() -> int:
    args = vars(build_parser().parse_args())
    output = args.pop("json_output")
    audit = run_reference_audit(**args)
    payload = json.dumps(audit.to_dict(), indent=2, default=str)
    print(payload)
    if output:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(payload + "\n", encoding="utf-8")
    return 0 if audit.passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
