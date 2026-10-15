"""Pure business rules for the weekly master and mobile reports."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any, Collection

from .constants import ACHIEVEMENT_MAX_DAYS, ORDER_STATUS_ALIASES, SERVICE_ALIASES
from .dates import as_date, excel_week_number, parse_date_only, parse_datetime


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def normalize_identifier(value: Any) -> str:
    return clean_text(value).casefold()


def normalize_service_group(value: Any) -> str:
    text = clean_text(value)
    return SERVICE_ALIASES.get(text.casefold(), text)


def normalize_order_status(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return "In Progress"
    return ORDER_STATUS_ALIASES.get(text.casefold(), text)


def is_unresolved_order_status(value: Any) -> bool:
    return clean_text(value).casefold() in {"", "#n/a", "n/a", "na"}


def to_decimal(value: Any) -> Decimal:
    if value is None:
        return Decimal("0")
    text = str(value).strip().replace(",", "")
    if not text or text == "-":
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"Invalid numeric value: {value!r}") from exc


def inclusive_networkdays(
    start: Any,
    end: Any,
    holidays: Collection[date] = (),
    *,
    normalize_zero: bool = False,
) -> int | None:
    """Return inclusive NETWORKDAYS; preserve zero unless normalization is requested."""
    start_date = as_date(start)
    end_date = as_date(end)
    if start_date is None or end_date is None:
        return None

    direction = 1 if start_date <= end_date else -1
    low, high = (start_date, end_date) if direction == 1 else (end_date, start_date)
    holiday_set = set(holidays)
    business_days = 0
    current = low
    while current <= high:
        if current.weekday() < 5 and current not in holiday_set:
            business_days += 1
        current += timedelta(days=1)
    result = direction * business_days
    if normalize_zero and result == 0:
        return 1
    return result


def achievement_for(
    *,
    service_group: Any,
    completion_date: Any,
    leadtime_cash_complete: int | None,
    pm_approved: bool = False,
) -> str | None:
    """Calculate achievement only where the reviewed rule is known."""
    if parse_datetime(completion_date) is None:
        return None
    if pm_approved:
        return "ACHIEVED"

    service = normalize_service_group(service_group)
    max_days = ACHIEVEMENT_MAX_DAYS.get(service)
    if max_days is None or leadtime_cash_complete is None or leadtime_cash_complete < 0:
        return None
    return "ACHIEVED" if leadtime_cash_complete <= max_days else "NOT ACHIEVED"


def build_master_record(
    source: dict[str, Any],
    *,
    row_number: int,
    holidays: Collection[date] = (),
) -> dict[str, Any]:
    """Convert one selected BQ row into the 34-column weekly master schema."""
    fab_upload = parse_datetime(source.get("FAB_Upload"))
    if fab_upload is None:
        raise ValueError("FAB_Upload is required for a selected weekly source row.")
    cash_date = parse_date_only(source.get("cash_date"))
    order_created = parse_date_only(source.get("order_crea"))
    completion_date = parse_date_only(source.get("order_completion_date"))
    service_group = normalize_service_group(source.get("service_group"))
    raw_status = clean_text(source.get("order_status")) or "In Progress"
    normalized_status = normalize_order_status(raw_status)
    otc = to_decimal(source.get("otc"))
    mrc = to_decimal(source.get("mrc"))
    total = otc + mrc
    fab_cash = inclusive_networkdays(fab_upload, cash_date, holidays)
    cash_complete = inclusive_networkdays(cash_date, completion_date, holidays, normalize_zero=True)

    return {
        "No.": row_number,
        "Month of FAB Upload": fab_upload.strftime("%B-%y"),
        "Week FAB": f"W{excel_week_number(fab_upload)}",
        "YEAR FAB Upload": str(fab_upload.year),
        "quote_num": clean_text(source.get("quote_num")),
        "ca_name": str(source.get("ca_name") or "").strip(),
        "service_group": service_group,
        "order_status": raw_status,
        "order_status2": normalized_status,
        "order_status (Confirmed by MIDO Dept.)": normalized_status,
        "fab_upload": fab_upload,
        "Count FAB": 1,
        "cash_dt": cash_date,
        "Count Cash": 1 if cash_date else None,
        "order number": clean_text(source.get("order_number")) or None,
        "order created": order_created,
        "order_completion_date": completion_date,
        "Count Completion": 1 if completion_date else None,
        "Week Completion": f"W{excel_week_number(completion_date)}" if completion_date else None,
        "Month Completion": completion_date.strftime("%b-%y") if completion_date else None,
        "Year Completion": str(completion_date.year) if completion_date else None,
        "otc": otc,
        "mrc": mrc,
        "OTC+MRC": total,
        "OTC+MRC (inMio)": total / Decimal("1000000"),
        "LEADTIME FAB-CASH": fab_cash,
        "LEADTIME CASH-COMPLETE": cash_complete,
        "ACHIEVEMENT": achievement_for(
            service_group=service_group,
            completion_date=completion_date,
            leadtime_cash_complete=cash_complete,
        ),
        "ADJUSTMENT": None,
        # This column is maintained manually downstream; mld_dt is not equivalent.
        "ACTIVATION DATE": None,
        "SALES NAME": str(source.get("sales_name") or "").strip() or None,
        # Manual W38 leaves hierarchy blank for newly appended rows.
        "DEPT SALES": None,
        "DIVISION SALES": None,
        "GROUP SALES": None,
    }
