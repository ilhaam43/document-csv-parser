"""Validate explicit weekly corrections without inferring approval from notes."""

from __future__ import annotations

import csv
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from .catalyst import iter_catalyst_review
from .dates import parse_datetime
from .matching import record_key
from .models import CheckpointRow, ReviewDecision, ReviewLoadResult
from .rules import clean_text, normalize_identifier, normalize_order_status


def _parse_pm_approval(value: Any, row_number: int) -> bool | None:
    text = clean_text(value).casefold()
    if not text:
        return None
    if text in {"yes", "true", "1"}:
        return True
    if text in {"no", "false", "0"}:
        return False
    raise ValueError(f"Review row {row_number}: pm_approved must be yes, no, or blank.")


def _parse_completion_date(value: Any, row_number: int) -> datetime | None:
    text = clean_text(value)
    if not text:
        return None
    try:
        return datetime.strptime(text, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(
            f"Review row {row_number}: reviewed_completion_date must be YYYY-MM-DD."
        ) from exc


def _parse_leadtime(value: Any, row_number: int) -> int | None:
    text = clean_text(value)
    if not text:
        return None
    try:
        number = int(text)
    except ValueError as exc:
        raise ValueError(
            f"Review row {row_number}: reviewed_fab_cash_leadtime must be a nonnegative integer."
        ) from exc
    if number < 0:
        raise ValueError(
            f"Review row {row_number}: reviewed_fab_cash_leadtime must be a nonnegative integer."
        )
    return number


def _reviewed_approval(reviewed: dict[str, Any], expected: dict[str, Any]) -> bool | None:
    if (
        clean_text(expected.get("ACHIEVEMENT")).upper() == "NOT ACHIEVED"
        and clean_text(reviewed.get("ACHIEVEMENT")).upper() == "ACHIEVED"
    ):
        return True
    return None


def load_review_decisions(
    path: str | Path,
    master_rows: list[dict[str, Any]],
    *,
    first_new_row: int | None = None,
    require_attribution: bool = False,
) -> ReviewLoadResult:
    """Return decisions keyed by Excel row, verifying every edited row's identity."""
    review_path = Path(path)
    if review_path.suffix.casefold() == ".xlsx":
        if first_new_row is None:
            raise ValueError("Legacy Catalyst workbook requires the first new Master row.")
        target_rows: dict[tuple[object, datetime], list[int]] = defaultdict(list)
        for row_number in range(first_new_row, len(master_rows) + 2):
            row = master_rows[row_number - 2]
            fab_upload = parse_datetime(row.get("fab_upload"))
            if fab_upload is not None:
                key = record_key(row.get("quote_num"), row.get("order number"))
                target_rows[(key, fab_upload.replace(second=0, microsecond=0))].append(row_number)
        checkpoint_matches: dict[tuple[object, datetime], list[dict[str, Any]]] = defaultdict(list)
        decisions: dict[int, ReviewDecision] = {}
        last_row_number = 1
        for row_number, reviewed in enumerate(iter_catalyst_review(review_path), start=2):
            last_row_number = row_number
            fab_upload = parse_datetime(reviewed.get("fab_upload"))
            if row_number < first_new_row:
                if row_number > len(master_rows) + 1:
                    raise ValueError("Catalyst checkpoint has more historical rows than the draft.")
                expected = master_rows[row_number - 2]
                expected_fab = parse_datetime(expected.get("fab_upload"))
                expected_minute = expected_fab.replace(second=0, microsecond=0) if expected_fab else None
                reviewed_minute = fab_upload.replace(second=0, microsecond=0) if fab_upload else None
                if (
                    normalize_identifier(reviewed.get("quote_num"))
                    != normalize_identifier(expected.get("quote_num"))
                    or reviewed_minute != expected_minute
                ):
                    raise ValueError(
                        f"Catalyst checkpoint row {row_number} does not match the Master draft; "
                        "use a row-identified review CSV."
                    )
                status = clean_text(reviewed.get("order_status"))
                reviewed_completion = parse_datetime(reviewed.get("order_completion_date"))
                corrected_completion = (
                    reviewed_completion
                    if normalize_order_status(status) == "COMPLETE"
                    and reviewed_completion is not None
                    and reviewed_completion != parse_datetime(expected.get("order_completion_date"))
                    else None
                )
                corrected_status = status if status and status != clean_text(expected.get("order_status")) else None
                approved = _reviewed_approval(reviewed, expected)
                if corrected_status or corrected_completion or approved:
                    decisions[row_number] = ReviewDecision(
                        order_status=corrected_status,
                        completion_date=corrected_completion,
                        pm_approved=approved,
                    )
            elif fab_upload is not None:
                key = record_key(reviewed.get("quote_num"), reviewed.get("order number"))
                match_key = (key, fab_upload.replace(second=0, microsecond=0))
                if match_key in target_rows:
                    checkpoint_matches[match_key].append(reviewed)
        if first_new_row > 2 and last_row_number < first_new_row - 1:
            raise ValueError("Catalyst checkpoint is missing historical Master rows.")
        checkpoint_rows: dict[int, CheckpointRow] = {}
        unmatched = 0
        for key, row_numbers in target_rows.items():
            matches = checkpoint_matches.get(key, [])
            if len(row_numbers) != len(matches):
                unmatched += len(row_numbers)
                continue
            for target, candidate in zip(row_numbers, matches, strict=True):
                leadtime = candidate.get("LEADTIME FAB-CASH")
                if leadtime is not None and (
                    not isinstance(leadtime, (int, float))
                    or isinstance(leadtime, bool)
                    or int(leadtime) != leadtime
                ):
                    raise ValueError(f"Checkpoint has an invalid FAB-CASH lead time for new row {target}.")
                checkpoint_rows[target] = CheckpointRow(
                    fab_upload=parse_datetime(candidate["fab_upload"]),
                    leadtime_fab_cash=int(leadtime) if leadtime is not None else None,
                )
                status = clean_text(candidate.get("order_status"))
                reviewed_completion = parse_datetime(candidate.get("order_completion_date"))
                corrected_completion = (
                    reviewed_completion
                    if normalize_order_status(status) == "COMPLETE"
                    and reviewed_completion is not None
                    and reviewed_completion != parse_datetime(master_rows[target - 2].get("order_completion_date"))
                    else None
                )
                corrected_status = (
                    status if status and status != clean_text(master_rows[target - 2].get("order_status")) else None
                )
                approved = _reviewed_approval(candidate, master_rows[target - 2])
                if corrected_status or corrected_completion or approved:
                    decisions[target] = ReviewDecision(
                        order_status=corrected_status,
                        completion_date=corrected_completion,
                        pm_approved=approved,
                    )
        return ReviewLoadResult(decisions, len(decisions), checkpoint_rows, unmatched)
    if review_path.suffix.casefold() != ".csv":
        raise ValueError("Weekly review must be a .csv or .xlsx file.")

    decisions: dict[int, ReviewDecision] = {}
    rows_by_key: dict[object, list[int]] = defaultdict(list)
    for row_number, row in enumerate(master_rows, start=2):
        rows_by_key[record_key(row.get("quote_num"), row.get("order number"))].append(row_number)
    with review_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"quote_num", "order_number", "reviewed_status"}
        if not required.issubset(reader.fieldnames or ()):
            raise ValueError("Weekly review CSV requires quote_num, order_number, and reviewed_status.")
        for csv_row_number, row in enumerate(reader, start=2):
            status = clean_text(row.get("reviewed_status")) or None
            completion_date = _parse_completion_date(row.get("reviewed_completion_date"), csv_row_number)
            leadtime_fab_cash = _parse_leadtime(
                row.get("reviewed_fab_cash_leadtime"), csv_row_number
            )
            pm_approved = _parse_pm_approval(row.get("pm_approved"), csv_row_number)
            if (
                status is None
                and completion_date is None
                and leadtime_fab_cash is None
                and pm_approved is None
            ):
                continue
            if status is not None and status.casefold() not in {
                "cancel", "cancelled", "canceled", "complete", "completed", "in progress", "open"
            }:
                raise ValueError(f"Review row {csv_row_number}: unsupported reviewed_status {status!r}.")
            reviewer = clean_text(row.get("reviewed_by"))
            reason = clean_text(row.get("review_reason"))
            if require_attribution and (not reviewer or not reason):
                raise ValueError(
                    f"Review row {csv_row_number}: reviewed_by and review_reason are required for a correction."
                )
            key = record_key(row.get("quote_num"), row.get("order_number"))
            master_row_text = clean_text(row.get("master_row"))
            if master_row_text:
                try:
                    target_rows = [int(master_row_text)]
                except ValueError as exc:
                    raise ValueError(f"Review row {csv_row_number}: invalid master_row.") from exc
            else:
                target_rows = rows_by_key.get(key, [])
                if len(target_rows) != 1:
                    raise ValueError(
                        f"Review row {csv_row_number}: master_row is required for an ambiguous quote/order."
                    )
            target = target_rows[0]
            if target < 2 or target > len(master_rows) + 1:
                raise ValueError(f"Review row {csv_row_number}: master_row is outside the prepared data.")
            expected = master_rows[target - 2]
            if key != record_key(expected.get("quote_num"), expected.get("order number")):
                raise ValueError(
                    f"Review row {csv_row_number}: quote/order does not match master_row {target}."
                )
            if target in decisions:
                raise ValueError(f"Review row {csv_row_number}: duplicate correction for master_row {target}.")
            decisions[target] = ReviewDecision(
                order_status=status,
                completion_date=completion_date,
                leadtime_fab_cash=leadtime_fab_cash,
                pm_approved=pm_approved,
                reviewed_by=reviewer,
                review_reason=reason,
            )
    return ReviewLoadResult(decisions, len(decisions))
