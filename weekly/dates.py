"""Date parsing and Excel-compatible week helpers."""

from __future__ import annotations

from datetime import date, datetime, timedelta
import re
from typing import Any


_DATE_FORMATS = (
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
)


def parse_datetime(value: Any) -> datetime | None:
    """Parse source/Excel date values without guessing non-date identifiers."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value <= 0:
            return None
        return datetime(1899, 12, 30) + timedelta(days=float(value))

    text = str(value).strip()
    if not text:
        return None
    iso_text = text.removesuffix(" UTC") + "+00:00" if text.endswith(" UTC") else text
    try:
        return datetime.fromisoformat(iso_text.replace("Z", "+00:00")).replace(tzinfo=None)
    except ValueError:
        pass
    for date_format in _DATE_FORMATS:
        try:
            return datetime.strptime(text, date_format)
        except ValueError:
            continue
    raise ValueError(f"Unsupported date value: {value!r}")


def has_explicit_seconds(value: Any) -> bool:
    """Whether a source timestamp preserves seconds, including an explicit :00."""
    if isinstance(value, datetime):
        return True
    if isinstance(value, str):
        return bool(re.search(r"\d{1,2}:\d{2}:\d{2}(?:\D|$)", value.strip()))
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def as_date(value: Any) -> date | None:
    parsed = parse_datetime(value)
    return parsed.date() if parsed else None


def parse_date_only(value: Any) -> datetime | None:
    """Lifecycle report columns contain calendar dates, unlike the FAB timestamp."""
    parsed = parse_datetime(value)
    return datetime.combine(parsed.date(), datetime.min.time()) if parsed else None


def excel_week_number(value: date | datetime, return_type: int = 14) -> int:
    """Match Excel WEEKNUM(date, 14), where weeks start on Thursday."""
    if return_type != 14:
        raise ValueError("The weekly report only supports Excel WEEKNUM return_type 14.")
    day = value.date() if isinstance(value, datetime) else value
    year_start = date(day.year, 1, 1)
    thursday_index = 3  # Python weekday: Monday=0.
    offset = (year_start.weekday() - thursday_index) % 7
    return ((day - year_start).days + offset) // 7 + 1


def same_reporting_month(value: date | datetime, report_date: date) -> bool:
    day = value.date() if isinstance(value, datetime) else value
    return (day.year, day.month) == (report_date.year, report_date.month)


def parse_report_month(value: str) -> date:
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", value):
        raise ValueError("report_month must be YYYY-MM.")
    return date.fromisoformat(f"{value}-01")
