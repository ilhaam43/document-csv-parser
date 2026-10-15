"""Refresh existing weekly master rows from BQ and Catalyst review data."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Mapping
from datetime import date
from typing import Any

from .catalyst import apply_catalyst_status
from .dates import excel_week_number, parse_date_only, parse_datetime
from .matching import record_key
from .models import CatalystOverride, RecordKey, ReviewDecision, ValidationIssue
from .rules import (
    achievement_for,
    clean_text,
    inclusive_networkdays,
    is_unresolved_order_status,
    normalize_identifier,
    normalize_order_status,
    to_decimal,
)


_LIFECYCLE_SOURCE_FIELDS = (
    "order_number",
    "order_status",
    "cash_date",
    "order_crea",
    "order_completion_date",
)


def _source_signature(row: Mapping[str, Any]) -> tuple[str, ...]:
    return tuple(clean_text(row.get(field)) for field in _LIFECYCLE_SOURCE_FIELDS)


class ExistingSourceIndex:
    """Memory-bounded lookup for only the source keys present in the master."""

    def __init__(self, master_rows: Iterable[Mapping[str, Any]]) -> None:
        self._target_quotes: set[str] = set()
        self._target_pairs: set[RecordKey] = set()
        for row in master_rows:
            key = record_key(row.get("quote_num"), row.get("order number"))
            if not key.quote_num:
                continue
            self._target_quotes.add(key.quote_num)
            if key.order_number:
                self._target_pairs.add(key)

        self._first_by_quote: dict[str, dict[str, Any]] = {}
        self._first_by_pair: dict[RecordKey, dict[str, Any]] = {}
        self._quote_candidate_signatures: dict[str, set[tuple[str, ...]]] = defaultdict(set)
        self._pair_conflicts: set[RecordKey] = set()

    def add(self, source: Mapping[str, Any]) -> None:
        quote = normalize_identifier(source.get("quote_num"))
        if not quote or quote not in self._target_quotes:
            return
        copied = dict(source)
        self._first_by_quote.setdefault(quote, copied)
        self._quote_candidate_signatures[quote].add(_source_signature(source))

        key = record_key(source.get("quote_num"), source.get("order_number"))
        if key not in self._target_pairs:
            return
        existing = self._first_by_pair.get(key)
        if existing is None:
            self._first_by_pair[key] = copied
        elif _source_signature(existing) != _source_signature(source):
            self._pair_conflicts.add(key)

    def resolve(self, master_row: Mapping[str, Any]) -> tuple[dict[str, Any] | None, ValidationIssue | None]:
        key = record_key(master_row.get("quote_num"), master_row.get("order number"))
        if key.order_number:
            if key in self._pair_conflicts:
                return None, ValidationIssue(
                    code="conflicting_source_pair",
                    message="BQ contains conflicting lifecycle values for the same quote and order.",
                    key=key,
                )
            return self._first_by_pair.get(key), None

        source = self._first_by_quote.get(key.quote_num)
        signatures = self._quote_candidate_signatures.get(key.quote_num, set())
        if source is not None and len(signatures) > 1:
            return source, ValidationIssue(
                code="quote_first_match_used",
                message=(
                    "Master order number is blank and BQ has multiple lifecycle candidates. "
                    "The first source row was used by the quote-lookup rule; recorded for audit."
                ),
                key=key,
            )
        return source, None

    def first_for_quote(self, quote_num: Any) -> dict[str, Any] | None:
        return self._first_by_quote.get(normalize_identifier(quote_num))

def build_existing_source_index(
    source_rows: Iterable[Mapping[str, Any]],
    master_rows: list[Mapping[str, Any]],
) -> ExistingSourceIndex:
    index = ExistingSourceIndex(master_rows)
    for source in source_rows:
        index.add(source)
    return index


def _fill_blank_from_source(current: Any, source_value: Any, *, date_value: bool = False) -> Any:
    """Fill missing lifecycle data without overwriting established history."""
    if clean_text(current) or source_value is None or clean_text(source_value) == "":
        return current
    return parse_date_only(source_value) if date_value else clean_text(source_value)


def refresh_existing_record(
    master_row: Mapping[str, Any],
    *,
    source: Mapping[str, Any] | None,
    catalyst_override: CatalystOverride | None,
    review_decision: ReviewDecision | None = None,
    holidays: set[date],
) -> dict[str, Any]:
    """Refresh source-owned fields, then recalculate all derived master fields."""
    updated = dict(master_row)
    previous_cash_date = parse_datetime(master_row.get("cash_dt"))
    previous_completion_date = parse_datetime(master_row.get("order_completion_date"))
    if catalyst_override is not None:
        updated = apply_catalyst_status(updated, catalyst_override)
    if review_decision is not None and review_decision.order_status is not None:
        updated["order_status"] = (
            "Cancelled" if normalize_order_status(review_decision.order_status) == "CANCEL"
            else review_decision.order_status
        )

    if source is not None:
        updated["cash_dt"] = _fill_blank_from_source(
            updated.get("cash_dt"), source.get("cash_date"), date_value=True
        )
        updated["order number"] = _fill_blank_from_source(
            updated.get("order number"), source.get("order_number")
        )
        updated["order created"] = _fill_blank_from_source(
            updated.get("order created"), source.get("order_crea"), date_value=True
        )
    has_reviewed_status = catalyst_override is not None or (
        review_decision is not None and review_decision.order_status is not None
    )
    previous_status = clean_text(master_row.get("order_status"))
    if not has_reviewed_status:
        source_status = clean_text(source.get("order_status")) if source is not None else ""
        previous_normalized = normalize_order_status(previous_status)
        if source is None and previous_normalized == "In Progress":
            updated["order_status"] = "#N/A"
        elif source_status and previous_normalized not in {"COMPLETE", "CANCEL"}:
            updated["order_status"] = source_status
        elif (
            source is not None
            and is_unresolved_order_status(previous_status)
            and not source_status
            and not clean_text(source.get("order_completion_date"))
        ):
            # A matched blank BQ lifecycle is active, unlike a missing quote.
            updated["order_status"] = "In Progress"

    raw_status = clean_text(updated.get("order_status"))
    updated["order_status"] = raw_status
    if is_unresolved_order_status(raw_status):
        normalized_status = normalize_order_status(master_row.get("order_status2"))
        updated["order_status2"] = master_row.get("order_status2") or normalized_status
        updated["order_status (Confirmed by MIDO Dept.)"] = (
            master_row.get("order_status (Confirmed by MIDO Dept.)") or normalized_status
        )
    else:
        normalized_status = normalize_order_status(raw_status)
        updated["order_status2"] = normalized_status
        updated["order_status (Confirmed by MIDO Dept.)"] = normalized_status
    if not is_unresolved_order_status(raw_status) and normalized_status == "COMPLETE" and source is not None:
        updated["order_completion_date"] = _fill_blank_from_source(
            updated.get("order_completion_date"),
            source.get("order_completion_date"),
            date_value=True,
        )
    if review_decision is not None and review_decision.completion_date is not None:
        if normalized_status != "COMPLETE":
            raise ValueError("Completion-date correction requires a complete order status.")
        updated["order_completion_date"] = parse_date_only(review_decision.completion_date)
    if has_reviewed_status and normalized_status != "COMPLETE":
        updated["order_completion_date"] = None
    if (
        not is_unresolved_order_status(raw_status)
        and normalized_status == "CANCEL"
        and source is not None
        and not clean_text(source.get("order_completion_date"))
    ):
        updated["order_completion_date"] = None

    fab_upload = parse_datetime(updated.get("fab_upload"))
    cash_date = parse_datetime(updated.get("cash_dt"))
    completion_date = parse_datetime(updated.get("order_completion_date"))
    cash_changed = cash_date != previous_cash_date
    completion_changed = completion_date != previous_completion_date
    if cash_changed or cash_date is None or fab_upload is None:
        updated["Count Cash"] = 1 if cash_date else None
        updated["LEADTIME FAB-CASH"] = inclusive_networkdays(
            fab_upload, cash_date, holidays
        )
    if completion_changed or completion_date is None:
        updated["Count Completion"] = 1 if completion_date else None
        if completion_date:
            updated["Week Completion"] = f"W{excel_week_number(completion_date)}"
            updated["Month Completion"] = completion_date.strftime("%b-%y")
            updated["Year Completion"] = str(completion_date.year)
        else:
            updated["Week Completion"] = None
            updated["Month Completion"] = None
            updated["Year Completion"] = None
            updated["LEADTIME CASH-COMPLETE"] = None
    if (cash_changed or completion_changed) and completion_date:
        updated["LEADTIME CASH-COMPLETE"] = inclusive_networkdays(
            cash_date, completion_date, holidays, normalize_zero=True
        )
        updated["ACHIEVEMENT"] = achievement_for(
            service_group=updated.get("service_group"),
            completion_date=completion_date,
            leadtime_cash_complete=updated["LEADTIME CASH-COMPLETE"],
        )
    if cash_date is None or completion_date is None:
        updated["LEADTIME CASH-COMPLETE"] = None
        updated["ACHIEVEMENT"] = (
            "ACHIEVED"
            if clean_text(master_row.get("ADJUSTMENT"))
            and master_row.get("ACHIEVEMENT") == "ACHIEVED"
            and normalized_status != "CANCEL"
            else None
        )
    if review_decision is not None and review_decision.leadtime_fab_cash is not None:
        if (
            review_decision.leadtime_fab_cash == 0
            and inclusive_networkdays(fab_upload, cash_date, holidays, normalize_zero=False) != 0
        ):
            raise ValueError("Reviewed zero FAB-CASH lead time does not match the source dates.")
        updated["LEADTIME FAB-CASH"] = review_decision.leadtime_fab_cash
    if review_decision is not None and review_decision.pm_approved is True:
        if completion_date is None:
            raise ValueError("PM approval requires an order completion date.")
        updated["ACHIEVEMENT"] = "ACHIEVED"
    elif review_decision is not None and review_decision.pm_approved is False:
        updated["ACHIEVEMENT"] = achievement_for(
            service_group=updated.get("service_group"),
            completion_date=completion_date,
            leadtime_cash_complete=updated.get("LEADTIME CASH-COMPLETE"),
        )

    otc = to_decimal(updated.get("otc"))
    mrc = to_decimal(updated.get("mrc"))
    total = otc + mrc
    updated["OTC+MRC"] = total
    updated["OTC+MRC (inMio)"] = total / to_decimal(1_000_000)
    return updated
