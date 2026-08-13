#!/usr/bin/env python3
"""Generate a Daily Tracking IDE workbook from the current IDE dashboard."""

from __future__ import annotations

import argparse
import calendar
import numbers
import os
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Iterable

import pandas as pd
from dateutil import parser as date_parser
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter


DEFAULT_INPUT_DIR = "input-ide"
DEFAULT_REFERENCE_DIR = "vlookup-ide"
DEFAULT_OUTPUT_DIR = "output-ide"

RAW_SHEET_NAME = "ICT ORDER 2026"
ALL_ORDER_SHEET_NAME = "ALL ORDER"
PIVOT_SHEET_NAME = "PIVOT"
TABLE_NAME = "Table47"
TABLE_STYLE_NAME = "TableStyleMedium2"

QUOTE_ID_HEADER = "QUOTE ID"
ORDER_TYPE_HEADER = "ORDER TYPE"
OTC_HEADER = "OTC"
MRC_HEADER = "MRC"
FAB_UPLOAD_HEADER = "FAB UPLOAD DATE"
YEAR_FAB_UPLOAD_HEADER = "YEAR FAB UPLOAD"
NEW_RFS_INITIAL_HEADER = "NEW RFS INITIAL"
ACTUAL_RFS_DATE_HEADER = "ACTUAL RFS DATE"
AGING_OF_RFS_HEADER = "Aging Of RFS"
STATUS_ORDER_HEADER = "STATUS ORDER"
AOS_SENT_DATE_HEADER = "AOS SENT DATE"
ON_TIME_DELIVERY_HEADER = "ON TIME DELIVERY"
STATUS_DELIVERY_HEADER = "STATUS DELIVERY"
PRE_INSTALLATION_START_HEADER = "Pre-Installation Start Date"
PRE_INSTALLATION_AGING_HEADER = "Pre-Installation Aging"
PRE_INSTALLATION_RANGE_HEADER = "Range Aging On Pre-Installation"
DEPT_HEADER = "DEPT."
PM_HEADER = "PM"
APM_HEADER = "APM"

DATE_HEADERS = (
    "QUOTE DATE",
    "FAB SIGN DATE",
    FAB_UPLOAD_HEADER,
    NEW_RFS_INITIAL_HEADER,
    "CASH DATE",
    "SO DATE",
    "SERVICE READY",
    ACTUAL_RFS_DATE_HEADER,
    AOS_SENT_DATE_HEADER,
    "AOS SIGN DATE",
    "SO COMPLETION DATE",
    PRE_INSTALLATION_START_HEADER,
)

DATE_ONLY_FORMAT = "mm/dd/yy"
DATE_TIME_FORMAT = "mm/dd/yy h:mm AM/PM"
YEAR_TEXT_FORMAT = "@"
FONT_NAME = "Indosat Regular"
FONT_SIZE = 11
MAX_COLUMN_WIDTH = 45
MIN_COLUMN_WIDTH = 10
PIVOT_PERCENTAGE_MIN_WIDTH = 24
UNKNOWN_YEAR = "1900"
EXCEL_ZERO_DATE_SERIAL = 0
EXCEL_NA_ERROR = "#N/A"
MIN_VALID_DATE_YEAR = 1900
MAX_VALID_DATE_YEAR = 2100
MISSING_TEXT = {"", "nan", "none", "null", "n/a", "#n/a", "nat"}
MAX_OPERATIONAL_FUTURE_YEARS = 5
PREVIOUS_FALLBACK_HEADERS = (
    DEPT_HEADER,
    PM_HEADER,
    APM_HEADER,
    STATUS_DELIVERY_HEADER,
)

STATUS_NORMALIZATION = {
    "pre-installation": "04-Pre-Installation",
    "04-pre-installation": "04-Pre-Installation",
    "installation": "06-Installation",
    "06-installation": "06-Installation",
    "waiting aos": "07-Waiting AOS/UAT On Hold",
    "07-waiting aos/uat on hold": "07-Waiting AOS/UAT On Hold",
    "so complete": "SO Complete",
    "cancel": "Cancel",
}


@dataclass(frozen=True)
class InputFiles:
    raw: Path
    previous: Path


@dataclass(frozen=True)
class BuildResult:
    dataframe: pd.DataFrame
    invalid_dates: dict[str, list[str]]
    zero_fallback_count: int
    duplicate_quote_ids: tuple[str, ...]
    duplicate_rows_preserved: int
    date_fallback_count: int
    field_fallback_count: int


@dataclass(frozen=True)
class PivotPercentageBlock:
    pivot_name: str
    header_row_offset: int
    original_header_row: int
    original_end_row: int
    scratch_start_row: int
    formula: str
    formula_row: int
    pivot_header_columns: dict[int, str]


def is_missing(value: object) -> bool:
    if value is None or value is pd.NA:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(value, str) and value.strip().lower() in MISSING_TEXT


def is_excel_zero_date(value: object) -> bool:
    """Return whether a value represents Excel's serial-zero date sentinel."""
    if isinstance(value, time):
        return value == time.min
    if isinstance(value, numbers.Real) and not isinstance(value, bool):
        return float(value) == EXCEL_ZERO_DATE_SERIAL
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return (value.year, value.month, value.day) == (1899, 12, 30)
    return isinstance(value, str) and value.strip() in {"0", "0.0", "00:00:00"}


def normalize_quote_id(value: object) -> str:
    if is_missing(value):
        return ""
    text = str(value).strip()
    return text[:-2] if text.endswith(".0") and text[:-2].isdigit() else text


def normalize_header(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).replace("\ufeff", "").strip())


def normalize_status(value: object) -> object:
    if is_missing(value):
        return pd.NA
    text = str(value).strip()
    return STATUS_NORMALIZATION.get(text.lower(), text)


def parse_report_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Report date must use YYYY-MM-DD.") from exc


def date_from_filename(path: Path) -> date | None:
    numeric_match = re.search(r"(?<!\d)(20\d{6})(?!\d)", path.stem)
    if numeric_match:
        try:
            return datetime.strptime(numeric_match.group(1), "%Y%m%d").date()
        except ValueError:
            pass

    named_match = re.search(
        r"(?<!\d)(\d{1,2})\s+([A-Za-z]+)\s+(20\d{2})(?!\d)",
        path.stem,
        flags=re.IGNORECASE,
    )
    if named_match:
        try:
            return datetime.strptime(" ".join(named_match.groups()), "%d %B %Y").date()
        except ValueError:
            pass

    separated_numeric_match = re.search(
        r"(?<!\d)(\d{1,2})[\s._-]+(\d{1,2})[\s._-]+((?:20)?\d{2})(?!\d)",
        path.stem,
    )
    if separated_numeric_match:
        day_text, month_text, year_text = separated_numeric_match.groups()
        if len(year_text) == 2:
            year_text = f"20{year_text}"
        try:
            return date(int(year_text), int(month_text), int(day_text))
        except ValueError:
            pass

    return None


def determine_report_date(explicit_date: date | None, raw: Path) -> date:
    if explicit_date is not None:
        return explicit_date

    raw_date = date_from_filename(raw)
    if raw_date is not None and raw.stem.casefold().startswith("ide dashboard"):
        return raw_date

    raise ValueError(
        "Raw workbook filename must include the report date, for example: "
        "'IDE DASHBOARD 15 August 2026.xlsx'."
    )


def resolve_single_file(path: Path, suffix: str, description: str) -> Path:
    path = path.resolve()
    if path.is_file():
        if path.suffix.lower() != suffix:
            raise ValueError(f"{description} must be a {suffix} file: {path}")
        return path
    if not path.is_dir():
        raise FileNotFoundError(f"{description} path does not exist: {path}")

    candidates = sorted(
        item
        for item in path.glob(f"*{suffix}")
        if item.is_file() and not item.name.startswith("~$")
    )
    if not candidates:
        raise FileNotFoundError(f"No {suffix} file found for {description} in: {path}")
    if len(candidates) > 1:
        names = ", ".join(item.name for item in candidates)
        raise ValueError(
            f"Multiple files found for {description}. Keep exactly one file in {path}: {names}"
        )
    return candidates[0]


def resolve_inputs(args: argparse.Namespace) -> InputFiles:
    return InputFiles(
        raw=resolve_single_file(args.input, ".xlsx", "IDE raw workbook"),
        previous=resolve_single_file(args.reference, ".xlsx", "previous IDE workbook"),
    )


def output_filename(report_date: date) -> str:
    month_name = calendar.month_name[report_date.month]
    return f"Daily Tracking IDE {report_date.day} {month_name} {report_date.year}.xlsx"


def resolve_output_path(output: Path, report_date: date) -> Path:
    output = output.resolve()
    if output.suffix.lower() == ".xlsx":
        return output
    return output / output_filename(report_date)


def target_header(report_date: date) -> str:
    return f"TARGET  Detemined as 1 {calendar.month_abbr[report_date.month]} {report_date:%y}"


def previous_status_header(previous: Path, report_date: date) -> str:
    previous_date = date_from_filename(previous) or (report_date - timedelta(days=1))
    return (
        f"STATUS DELIVERY {previous_date.day} "
        f"{calendar.month_name[previous_date.month]} {previous_date.year}"
    )


def read_raw_dataframe(raw_path: Path) -> pd.DataFrame:
    raw = pd.read_excel(
        raw_path,
        sheet_name=RAW_SHEET_NAME,
        header=0,
        skiprows=[1],
        dtype=object,
    )
    raw_workbook = load_workbook(raw_path, read_only=True, data_only=True, keep_links=False)
    try:
        raw_worksheet = raw_workbook[RAW_SHEET_NAME]
        for dataframe_row, worksheet_row in enumerate(
            raw_worksheet.iter_rows(min_row=3, max_col=len(raw.columns))
        ):
            if dataframe_row >= len(raw):
                break
            for column_index, cell in enumerate(worksheet_row):
                if cell.data_type == "e":
                    raw.iat[dataframe_row, column_index] = str(cell.value)
    finally:
        raw_workbook.close()

    raw.columns = [normalize_header(column) for column in raw.columns]
    if QUOTE_ID_HEADER not in raw.columns:
        raise ValueError(f"Raw sheet is missing required column: {QUOTE_ID_HEADER}")

    quote_ids = raw[QUOTE_ID_HEADER].map(normalize_quote_id)
    raw = raw.loc[quote_ids.ne("")].copy()
    raw[QUOTE_ID_HEADER] = quote_ids.loc[raw.index]
    return raw.reset_index(drop=True)


def read_previous_dataframe(previous_path: Path) -> pd.DataFrame:
    previous = pd.read_excel(
        previous_path,
        sheet_name=ALL_ORDER_SHEET_NAME,
        dtype=object,
    )
    previous.columns = [normalize_header(column) for column in previous.columns]
    if QUOTE_ID_HEADER not in previous.columns:
        raise ValueError(f"Previous workbook is missing required column: {QUOTE_ID_HEADER}")
    previous[QUOTE_ID_HEADER] = previous[QUOTE_ID_HEADER].map(normalize_quote_id)
    return previous


def occurrence_keys(dataframe: pd.DataFrame) -> list[tuple[str, int]]:
    quote_ids = dataframe[QUOTE_ID_HEADER].map(normalize_quote_id)
    occurrences = quote_ids.groupby(quote_ids, sort=False).cumcount()
    return list(zip(quote_ids, occurrences, strict=False))


def occurrence_lookup(
    previous: pd.DataFrame,
    value_header: str,
) -> dict[tuple[str, int], object]:
    if value_header not in previous.columns:
        return {}
    return dict(
        zip(occurrence_keys(previous), previous[value_header], strict=False)
    )


def first_quote_lookup(previous: pd.DataFrame, value_header: str) -> dict[str, object]:
    if value_header not in previous.columns:
        return {}
    lookup: dict[str, object] = {}
    for quote_id, value in zip(
        previous[QUOTE_ID_HEADER].map(normalize_quote_id),
        previous[value_header],
        strict=False,
    ):
        if quote_id and quote_id not in lookup:
            lookup[quote_id] = value
    return lookup


def apply_previous_field_fallbacks(
    current: pd.DataFrame,
    previous: pd.DataFrame,
    headers: Iterable[str],
) -> int:
    fallback_count = 0
    current_keys = occurrence_keys(current)
    for header in headers:
        if header not in current.columns or header not in previous.columns:
            continue
        previous_lookup = occurrence_lookup(previous, header)
        values: list[object] = []
        for current_value, key in zip(current[header], current_keys, strict=False):
            previous_value = previous_lookup.get(key)
            if is_missing(current_value) and not is_missing(previous_value):
                values.append(previous_value)
                fallback_count += 1
            else:
                values.append(current_value)
        current[header] = pd.Series(values, index=current.index, dtype=object)
    return fallback_count


def numeric_value(value: object) -> int | float | None:
    if is_missing(value):
        return None
    if isinstance(value, str):
        value = value.strip().replace(",", "").replace(" ", "")
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return None
    number = float(number)
    return int(number) if number.is_integer() else number


def parse_excel_datetime(value: object) -> datetime | None:
    if is_missing(value) or isinstance(value, time):
        return None
    if isinstance(value, pd.Timestamp):
        parsed = value.to_pydatetime()
        return parsed if MIN_VALID_DATE_YEAR <= parsed.year <= MAX_VALID_DATE_YEAR else None
    if isinstance(value, datetime):
        return value if MIN_VALID_DATE_YEAR <= value.year <= MAX_VALID_DATE_YEAR else None
    if isinstance(value, date):
        parsed = datetime.combine(value, time.min)
        return parsed if MIN_VALID_DATE_YEAR <= parsed.year <= MAX_VALID_DATE_YEAR else None

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 1 <= float(value) <= 2_958_465:
            try:
                parsed = datetime(1899, 12, 30) + timedelta(days=float(value))
                return parsed if MIN_VALID_DATE_YEAR <= parsed.year <= MAX_VALID_DATE_YEAR else None
            except (OverflowError, ValueError):
                return None

    text = str(value).strip()
    if text.lower() in MISSING_TEXT or text in {"-", "00:00:00"}:
        return None
    text = re.sub(r"/{2,}", "/", text)
    text = re.sub(r"\s+UTC$", "", text, flags=re.IGNORECASE)

    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
    )
    for fmt in formats:
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed if MIN_VALID_DATE_YEAR <= parsed.year <= MAX_VALID_DATE_YEAR else None
        except ValueError:
            continue

    try:
        parsed = date_parser.parse(text, dayfirst=False, fuzzy=False)
        return parsed if MIN_VALID_DATE_YEAR <= parsed.year <= MAX_VALID_DATE_YEAR else None
    except (OverflowError, TypeError, ValueError):
        return None


def normalize_date_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, list[str]]]:
    result = df.copy()
    invalid_dates: dict[str, list[str]] = {}
    for header in DATE_HEADERS:
        if header not in result.columns:
            continue
        normalized_values: list[object] = []
        invalid_values: list[str] = []
        for value in result[header]:
            parsed = parse_excel_datetime(value)
            if is_excel_zero_date(value):
                normalized_values.append(EXCEL_ZERO_DATE_SERIAL)
            elif parsed is not None:
                normalized_values.append(parsed)
            elif is_missing(value) or str(value).strip() == "-":
                normalized_values.append(None)
            else:
                normalized_values.append(str(value).strip())
                invalid_values.append(str(value).strip())
        result[header] = pd.Series(normalized_values, index=result.index, dtype=object)
        if invalid_values:
            invalid_dates[header] = sorted(set(invalid_values))
    return result, invalid_dates


def target_value(value: object, report_date: date) -> str:
    parsed = parse_excel_datetime(value)
    month_name = calendar.month_name[report_date.month]
    if parsed is None:
        return "Target Not Yet Inputted"
    source_month = (parsed.year, parsed.month)
    report_month = (report_date.year, report_date.month)
    if source_month < report_month:
        return f"Before {month_name}"
    if source_month > report_month:
        return f"After {month_name}"
    return f"Target {month_name}"


def is_plausible_operational_date(value: datetime | None, report_date: date) -> bool:
    if value is None:
        return False
    if value.year == MIN_VALID_DATE_YEAR:
        return True
    return MIN_VALID_DATE_YEAR < value.year <= report_date.year + MAX_OPERATIONAL_FUTURE_YEARS


def differs_only_by_year(current_value: datetime | None, previous_value: datetime | None) -> bool:
    if current_value is None or previous_value is None:
        return False
    return (
        current_value.year != previous_value.year
        and (current_value.month, current_value.day)
        == (previous_value.month, previous_value.day)
    )


def reconciled_date_value(
    current_value: object,
    previous_value: object,
    report_date: date,
) -> tuple[object, bool]:
    if is_excel_zero_date(current_value):
        return EXCEL_ZERO_DATE_SERIAL, False

    current_date = parse_excel_datetime(current_value)
    previous_date = parse_excel_datetime(previous_value)

    if differs_only_by_year(current_date, previous_date) and is_plausible_operational_date(
        previous_date,
        report_date,
    ):
        return previous_date, True
    if is_plausible_operational_date(current_date, report_date):
        return current_value, False
    if is_excel_zero_date(previous_value):
        return EXCEL_ZERO_DATE_SERIAL, True
    if is_plausible_operational_date(previous_date, report_date):
        return previous_date, True
    return current_value, False


def status_order_value(aging: object) -> object:
    if is_missing(aging):
        return pd.NA
    aging_number = numeric_value(aging)
    if aging_number is None:
        return pd.NA
    if aging_number < -5:
        return "On Track"
    if aging_number < 0:
        return "Potential Delay"
    return "Delay"


def pre_installation_range_value(aging: object) -> object:
    aging_number = numeric_value(aging)
    if aging_number is None:
        return pd.NA
    if aging_number <= 14:
        return "1-7 Days"
    if aging_number <= 30:
        return "8-14 Days"
    if aging_number <= 60:
        return "15-30 Days"
    return "> 30 Days"


def financial_value_or_zero(value: object) -> int | float:
    parsed = numeric_value(value)
    return 0 if parsed is None else parsed


def insert_after(df: pd.DataFrame, after_header: str, header: str, values: Iterable[object]) -> None:
    if after_header not in df.columns:
        raise ValueError(f"Could not insert {header}; missing anchor column: {after_header}")
    position = df.columns.get_loc(after_header) + 1
    df.insert(position, header, list(values))


def build_all_order(
    raw: pd.DataFrame,
    previous: pd.DataFrame,
    report_date: date,
    previous_path: Path,
) -> BuildResult:
    result = raw.copy().reset_index(drop=True)
    previous = previous.copy().reset_index(drop=True)
    quote_ids = result[QUOTE_ID_HEADER].map(normalize_quote_id)
    duplicate_mask = quote_ids.duplicated(keep=False)
    duplicate_quote_ids = tuple(sorted(quote_ids.loc[duplicate_mask].unique()))
    duplicate_rows_preserved = int(quote_ids.duplicated(keep="first").sum())
    current_keys = occurrence_keys(result)
    previous_keys = set(occurrence_keys(previous))
    field_fallback_count = apply_previous_field_fallbacks(
        result,
        previous,
        PREVIOUS_FALLBACK_HEADERS,
    )
    date_fallback_count = 0

    for header in DATE_HEADERS:
        if header not in result.columns:
            continue
        previous_date_lookup = occurrence_lookup(previous, header)
        reconciled_values: list[object] = []
        for current_value, key in zip(result[header], current_keys, strict=False):
            reconciled_value, used_fallback = reconciled_date_value(
                current_value,
                previous_date_lookup.get(key),
                report_date,
            )
            reconciled_values.append(reconciled_value)
            date_fallback_count += int(used_fallback)
        result[header] = pd.Series(reconciled_values, index=result.index, dtype=object)

    order_type_lookup = occurrence_lookup(previous, ORDER_TYPE_HEADER)
    otc_lookup = occurrence_lookup(previous, OTC_HEADER)
    mrc_lookup = occurrence_lookup(previous, MRC_HEADER)
    previous_status_lookup = first_quote_lookup(previous, STATUS_DELIVERY_HEADER)
    previous_quote_ids = set(previous[QUOTE_ID_HEADER].map(normalize_quote_id))
    pre_installation_start_lookup = first_quote_lookup(previous, PRE_INSTALLATION_START_HEADER)

    order_types: list[object] = []
    otc_values: list[object] = []
    mrc_values: list[object] = []
    previous_status_values: list[object] = []
    pre_installation_start_values: list[object] = []
    zero_fallback_count = 0

    current_status_values = result[STATUS_DELIVERY_HEADER].map(normalize_status)
    for quote_id, key, current_status in zip(
        quote_ids,
        current_keys,
        current_status_values,
        strict=False,
    ):
        previous_order_type = order_type_lookup.get(key)
        order_types.append(
            "New Registration" if is_missing(previous_order_type) else str(previous_order_type).strip()
        )

        previous_otc = numeric_value(otc_lookup.get(key))
        previous_mrc = numeric_value(mrc_lookup.get(key))
        zero_fallback_count += int(previous_otc is None or previous_mrc is None)
        otc_values.append(financial_value_or_zero(previous_otc))
        mrc_values.append(financial_value_or_zero(previous_mrc))

        if quote_id not in previous_quote_ids:
            previous_status_values.append(EXCEL_NA_ERROR)
        else:
            previous_status = previous_status_lookup.get(quote_id)
            previous_status_values.append(0 if is_missing(previous_status) else previous_status)

        previous_start = pre_installation_start_lookup.get(quote_id)
        is_active_order = (
            not is_missing(current_status)
            and str(current_status).strip().casefold() not in {"so complete", "cancel"}
        )
        if quote_id not in previous_quote_ids or (is_missing(previous_start) and is_active_order):
            pre_installation_start_values.append(
                datetime.combine(report_date - timedelta(days=1), time.min)
            )
        elif is_excel_zero_date(previous_start):
            pre_installation_start_values.append(EXCEL_ZERO_DATE_SERIAL)
        else:
            pre_installation_start_values.append(parse_excel_datetime(previous_start))

    insert_after(result, "QUOTE NAME", ORDER_TYPE_HEADER, order_types)
    insert_after(result, "PARTNERS", OTC_HEADER, otc_values)
    insert_after(result, OTC_HEADER, MRC_HEADER, mrc_values)

    fab_dates = result[FAB_UPLOAD_HEADER].map(parse_excel_datetime)
    year_values = [str(value.year) if not is_missing(value) else UNKNOWN_YEAR for value in fab_dates]
    insert_after(result, FAB_UPLOAD_HEADER, YEAR_FAB_UPLOAD_HEADER, year_values)

    dynamic_target_header = target_header(report_date)
    target_values = result[NEW_RFS_INITIAL_HEADER].map(
        lambda value: target_value(value, report_date)
    )
    insert_after(result, NEW_RFS_INITIAL_HEADER, dynamic_target_header, target_values)

    actual_rfs_dates = result[ACTUAL_RFS_DATE_HEADER].map(parse_excel_datetime)
    aging_values = [
        (report_date - value.date()).days if not is_missing(value) else pd.NA
        for value in actual_rfs_dates
    ]
    insert_after(result, ACTUAL_RFS_DATE_HEADER, AGING_OF_RFS_HEADER, aging_values)
    insert_after(
        result,
        AGING_OF_RFS_HEADER,
        STATUS_ORDER_HEADER,
        [status_order_value(value) for value in aging_values],
    )

    insert_after(result, AOS_SENT_DATE_HEADER, ON_TIME_DELIVERY_HEADER, [pd.NA] * len(result))

    dynamic_previous_status_header = previous_status_header(previous_path, report_date)
    insert_after(
        result,
        STATUS_DELIVERY_HEADER,
        dynamic_previous_status_header,
        previous_status_values,
    )
    insert_after(
        result,
        dynamic_previous_status_header,
        PRE_INSTALLATION_START_HEADER,
        pre_installation_start_values,
    )

    pre_installation_aging_values = [
        (report_date - value.date()).days if isinstance(value, datetime) else pd.NA
        for value in pre_installation_start_values
    ]
    insert_after(
        result,
        PRE_INSTALLATION_START_HEADER,
        PRE_INSTALLATION_AGING_HEADER,
        pre_installation_aging_values,
    )
    insert_after(
        result,
        PRE_INSTALLATION_AGING_HEADER,
        PRE_INSTALLATION_RANGE_HEADER,
        [pre_installation_range_value(value) for value in pre_installation_aging_values],
    )

    result[STATUS_DELIVERY_HEADER] = result[STATUS_DELIVERY_HEADER].map(normalize_status)
    result, invalid_dates = normalize_date_columns(result)
    result = result.reset_index(drop=True)

    if len(result.columns) != 61:
        raise ValueError(
            f"ALL ORDER must contain 61 columns, found {len(result.columns)}. "
            "The raw workbook structure may have changed."
        )

    return BuildResult(
        dataframe=result,
        invalid_dates=invalid_dates,
        zero_fallback_count=zero_fallback_count,
        duplicate_quote_ids=duplicate_quote_ids,
        duplicate_rows_preserved=duplicate_rows_preserved,
        date_fallback_count=date_fallback_count,
        field_fallback_count=field_fallback_count,
    )


def com_value(value: object) -> object:
    if value is None or value is pd.NA:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        delta = value - datetime(1899, 12, 30)
        return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86_400
    if isinstance(value, date) and not isinstance(value, datetime):
        delta = datetime.combine(value, time.min) - datetime(1899, 12, 30)
        return delta.days
    if isinstance(value, bool):
        return value
    if isinstance(value, numbers.Integral):
        return int(value)
    if isinstance(value, numbers.Real):
        return float(value)
    return value


def find_sheet(workbook, sheet_name: str):
    for sheet in workbook.Worksheets:
        if str(sheet.Name) == sheet_name:
            return sheet
    return None


def apply_date_formats(worksheet, df: pd.DataFrame) -> None:
    def contiguous_ranges(rows: list[int]) -> Iterable[tuple[int, int]]:
        if not rows:
            return
        start = previous = rows[0]
        for row_number in rows[1:]:
            if row_number != previous + 1:
                yield start, previous
                start = row_number
            previous = row_number
        yield start, previous

    for header in DATE_HEADERS:
        if header not in df.columns:
            continue
        column_number = df.columns.get_loc(header) + 1
        column_letter = get_column_letter(column_number)
        date_only_rows: list[int] = []
        date_time_rows: list[int] = []
        for row_number, value in enumerate(df[header], start=2):
            if is_excel_zero_date(value):
                date_time_rows.append(row_number)
                continue
            if not isinstance(value, datetime):
                continue
            if value.time() == time.min:
                date_only_rows.append(row_number)
            else:
                date_time_rows.append(row_number)

        for rows, number_format in (
            (date_only_rows, DATE_ONLY_FORMAT),
            (date_time_rows, DATE_TIME_FORMAT),
        ):
            for start_row, end_row in contiguous_ranges(rows):
                worksheet.Range(
                    f"{column_letter}{start_row}:{column_letter}{end_row}"
                ).NumberFormat = number_format


def apply_previous_status_na_errors(excel, worksheet, df: pd.DataFrame) -> None:
    status_headers = [
        header
        for header in df.columns
        if str(header).startswith(f"{STATUS_DELIVERY_HEADER} ")
    ]
    for header in status_headers:
        column_number = df.columns.get_loc(header) + 1
        worksheet.Range(
            worksheet.Cells(2, column_number),
            worksheet.Cells(len(df) + 1, column_number),
        ).NumberFormat = "General"
        error_rows = [
            row_number
            for row_number, value in enumerate(df[header], start=2)
            if str(value).strip().upper() == EXCEL_NA_ERROR
        ]
        if not error_rows:
            continue

        start = previous = error_rows[0]
        ranges: list[tuple[int, int]] = []
        for row_number in error_rows[1:]:
            if row_number != previous + 1:
                ranges.append((start, previous))
                start = row_number
            previous = row_number
        ranges.append((start, previous))

        auto_fill_setting = None
        try:
            auto_fill_setting = excel.AutoCorrect.AutoFillFormulasInLists
            excel.AutoCorrect.AutoFillFormulasInLists = False
        except Exception:
            auto_fill_setting = None
        try:
            for start_row, end_row in ranges:
                target = worksheet.Range(
                    worksheet.Cells(start_row, column_number),
                    worksheet.Cells(end_row, column_number),
                )
                target.Formula = "=NA()"
                target.Calculate()
                target.Copy()
                target.PasteSpecial(Paste=-4163)  # xlPasteValues
                excel.CutCopyMode = False
        finally:
            if auto_fill_setting is not None:
                try:
                    excel.AutoCorrect.AutoFillFormulasInLists = auto_fill_setting
                except Exception:
                    pass


def autofit_with_limits(worksheet, column_count: int) -> None:
    for column_number in range(1, column_count + 1):
        column = worksheet.Columns(column_number)
        column.AutoFit()
        width = float(column.ColumnWidth)
        if width > MAX_COLUMN_WIDTH:
            column.ColumnWidth = MAX_COLUMN_WIDTH
        elif width < MIN_COLUMN_WIDTH:
            column.ColumnWidth = MIN_COLUMN_WIDTH


def pivot_header_columns(worksheet, table_range, header_row: int) -> dict[int, str]:
    headers: dict[int, str] = {}
    start_column = int(table_range.Column)
    end_column = start_column + int(table_range.Columns.Count) - 1
    for column_number in range(start_column, end_column + 1):
        value = worksheet.Cells(header_row, column_number).Value
        if value is None:
            continue
        normalized = " ".join(str(value).strip().casefold().split())
        if normalized:
            headers[column_number] = normalized
    return headers


def capture_pivot_percentage_blocks(workbook, worksheet):
    """Move percentage blocks aside so dynamic PivotTable columns can expand."""
    scratch = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    scratch.Name = "__PERCENTAGE_LAYOUT__"

    used_range = worksheet.UsedRange
    used_values = used_range.Value2
    percentage_headers: list[tuple[int, int]] = []
    if used_values is not None:
        for row_offset, row in enumerate(used_values):
            for column_offset, value in enumerate(row):
                if isinstance(value, str) and "percentage" in value.casefold():
                    percentage_headers.append(
                        (used_range.Row + row_offset, used_range.Column + column_offset)
                    )

    blocks: list[PivotPercentageBlock] = []
    scratch_row = 1
    pivot_tables = worksheet.PivotTables()
    for header_row, percentage_column in percentage_headers:
        candidates = []
        for pivot_index in range(1, pivot_tables.Count + 1):
            pivot_table = pivot_tables(pivot_index)
            table_range = pivot_table.TableRange2
            table_bottom = table_range.Row + table_range.Rows.Count - 1
            table_right = table_range.Column + table_range.Columns.Count - 1
            if table_range.Row <= header_row <= table_bottom and table_right < percentage_column:
                candidates.append((table_right, pivot_table))
        if not candidates:
            continue

        _, pivot_table = max(candidates, key=lambda item: item[0])
        table_range = pivot_table.TableRange2
        formula_row = header_row + 1
        formula = str(worksheet.Cells(formula_row, percentage_column).Formula or "")
        if not formula.startswith("="):
            continue

        end_row = formula_row
        while end_row + 1 <= used_range.Row + used_range.Rows.Count - 1:
            next_formula = str(
                worksheet.Cells(end_row + 1, percentage_column).Formula or ""
            )
            if not next_formula.startswith("="):
                break
            end_row += 1

        source_range = worksheet.Range(
            worksheet.Cells(header_row, percentage_column),
            worksheet.Cells(end_row, percentage_column),
        )
        source_range.Copy(Destination=scratch.Cells(scratch_row, 1))
        blocks.append(
            PivotPercentageBlock(
                pivot_name=str(pivot_table.Name),
                header_row_offset=header_row - int(table_range.Row),
                original_header_row=header_row,
                original_end_row=end_row,
                scratch_start_row=scratch_row,
                formula=formula,
                formula_row=formula_row,
                pivot_header_columns=pivot_header_columns(
                    worksheet,
                    table_range,
                    header_row,
                ),
            )
        )
        source_range.Clear()
        scratch_row += end_row - header_row + 2

    scratch.Visible = 2  # xlSheetVeryHidden
    return scratch, blocks


def remap_percentage_formula(
    formula: str,
    source_row: int,
    target_row: int,
    column_mapping: dict[int, int],
) -> str:
    cell_reference = re.compile(r"(?<![A-Z0-9_])(\$?)([A-Z]{1,3})(\$?)(\d+)")
    row_delta = target_row - source_row

    def replace(match: re.Match[str]) -> str:
        column_absolute, column_text, row_absolute, row_text = match.groups()
        source_column = column_index_from_string(column_text)
        target_column = column_mapping.get(source_column, source_column)
        row_number = int(row_text)
        if not row_absolute:
            row_number += row_delta
        return (
            f"{column_absolute}{get_column_letter(target_column)}"
            f"{row_absolute}{row_number}"
        )

    return cell_reference.sub(replace, formula)


def restore_pivot_percentage_blocks(worksheet, scratch, blocks: list[PivotPercentageBlock]) -> None:
    for block in blocks:
        pivot_table = worksheet.PivotTables(block.pivot_name)
        table_range = pivot_table.TableRange2
        if int(table_range.Rows.Count) <= 1 or int(table_range.Columns.Count) <= 1:
            raise RuntimeError(
                f"PivotTable '{block.pivot_name}' did not render after refresh."
            )

        header_row = int(table_range.Row) + block.header_row_offset
        percentage_column = int(table_range.Column) + int(table_range.Columns.Count)
        end_row = int(table_range.Row) + int(table_range.Rows.Count) - 1
        source_height = block.original_end_row - block.original_header_row + 1
        source_range = scratch.Range(
            scratch.Cells(block.scratch_start_row, 1),
            scratch.Cells(block.scratch_start_row + source_height - 1, 1),
        )
        destination_range = worksheet.Range(
            worksheet.Cells(header_row, percentage_column),
            worksheet.Cells(header_row + source_height - 1, percentage_column),
        )
        source_range.Copy(Destination=destination_range)

        formula_start_row = header_row + 1
        if end_row < formula_start_row:
            continue

        if end_row != header_row + source_height - 1:
            body_template = scratch.Cells(block.scratch_start_row + 1, 1)
            grand_total_template = scratch.Cells(
                block.scratch_start_row + source_height - 1,
                1,
            )
            if end_row > formula_start_row:
                body_template.Copy(
                    Destination=worksheet.Range(
                        worksheet.Cells(formula_start_row, percentage_column),
                        worksheet.Cells(end_row - 1, percentage_column),
                    )
                )
            grand_total_template.Copy(
                Destination=worksheet.Cells(end_row, percentage_column)
            )
            if end_row < header_row + source_height - 1:
                worksheet.Range(
                    worksheet.Cells(end_row + 1, percentage_column),
                    worksheet.Cells(header_row + source_height - 1, percentage_column),
                ).Clear()

        current_headers = pivot_header_columns(worksheet, table_range, header_row)
        current_columns_by_header = {
            header: column for column, header in current_headers.items()
        }
        column_mapping = {
            old_column: current_columns_by_header[header]
            for old_column, header in block.pivot_header_columns.items()
            if header in current_columns_by_header
        }
        formulas = tuple(
            (
                remap_percentage_formula(
                    block.formula,
                    block.formula_row,
                    row_number,
                    column_mapping,
                ),
            )
            for row_number in range(formula_start_row, end_row + 1)
        )
        worksheet.Range(
            worksheet.Cells(formula_start_row, percentage_column),
            worksheet.Cells(end_row, percentage_column),
        ).Formula = formulas


def adjust_pivot_column_widths(worksheet) -> None:
    pivot_columns: set[int] = set()
    pivot_tables = worksheet.PivotTables()
    for pivot_index in range(1, pivot_tables.Count + 1):
        table_range = pivot_tables(pivot_index).TableRange2
        pivot_columns.update(
            range(table_range.Column, table_range.Column + table_range.Columns.Count)
        )

    percentage_columns: set[int] = set()
    used_range = worksheet.UsedRange
    used_values = used_range.Value2
    if used_values is not None:
        for row in used_values:
            for offset, value in enumerate(row):
                if isinstance(value, str) and "percentage" in value.casefold():
                    percentage_columns.add(used_range.Column + offset)

    for column_number in sorted(pivot_columns | percentage_columns):
        column = worksheet.Columns(column_number)
        column.AutoFit()
        width = float(column.ColumnWidth)
        if width > MAX_COLUMN_WIDTH:
            column.ColumnWidth = MAX_COLUMN_WIDTH
        elif width < MIN_COLUMN_WIDTH:
            column.ColumnWidth = MIN_COLUMN_WIDTH

    for column_number in percentage_columns:
        column = worksheet.Columns(column_number)
        if float(column.ColumnWidth) < PIVOT_PERCENTAGE_MIN_WIDTH:
            column.ColumnWidth = PIVOT_PERCENTAGE_MIN_WIDTH


def repair_pivot_percentage_ranges(worksheet) -> None:
    used_range = worksheet.UsedRange
    used_values = used_range.Value2
    percentage_headers: list[tuple[int, int]] = []
    if used_values is not None:
        for row_offset, row in enumerate(used_values):
            for column_offset, value in enumerate(row):
                if isinstance(value, str) and "percentage" in value.casefold():
                    percentage_headers.append(
                        (used_range.Row + row_offset, used_range.Column + column_offset)
                    )

    pivot_tables = worksheet.PivotTables()
    for header_row, percentage_column in percentage_headers:
        related_bottom_rows: list[int] = []
        for pivot_index in range(1, pivot_tables.Count + 1):
            table_range = pivot_tables(pivot_index).TableRange2
            table_bottom = table_range.Row + table_range.Rows.Count - 1
            table_right = table_range.Column + table_range.Columns.Count - 1
            if table_range.Row <= header_row <= table_bottom and table_right < percentage_column:
                related_bottom_rows.append(table_bottom)
        if not related_bottom_rows:
            continue

        target_end_row = max(related_bottom_rows)
        formula_start_row = header_row + 1
        first_formula_cell = worksheet.Cells(formula_start_row, percentage_column)
        formula_r1c1 = str(first_formula_cell.FormulaR1C1 or "")
        if not formula_r1c1.startswith("="):
            continue

        old_end_row = formula_start_row
        while old_end_row + 1 <= used_range.Row + used_range.Rows.Count - 1:
            next_formula = str(
                worksheet.Cells(old_end_row + 1, percentage_column).FormulaR1C1 or ""
            )
            if not next_formula.startswith("="):
                break
            old_end_row += 1

        old_grand_total_cell = worksheet.Cells(old_end_row, percentage_column)
        target_grand_total_cell = worksheet.Cells(target_end_row, percentage_column)
        if target_end_row != old_end_row:
            old_grand_total_cell.Copy(Destination=target_grand_total_cell)

        if target_end_row > old_end_row:
            body_style_source = worksheet.Cells(max(formula_start_row, old_end_row - 1), percentage_column)
            body_style_source.Copy(
                Destination=worksheet.Range(
                    worksheet.Cells(old_end_row, percentage_column),
                    worksheet.Cells(target_end_row - 1, percentage_column),
                )
            )
        elif target_end_row < old_end_row:
            worksheet.Range(
                worksheet.Cells(target_end_row + 1, percentage_column),
                worksheet.Cells(old_end_row, percentage_column),
            ).Clear()

        worksheet.Range(
            worksheet.Cells(formula_start_row, percentage_column),
            worksheet.Cells(target_end_row, percentage_column),
        ).FormulaR1C1 = formula_r1c1


def validate_target_complete_pivots(worksheet, dataframe: pd.DataFrame) -> None:
    pivot_tables = worksheet.PivotTables()
    top_pivots = [
        pivot_tables(index)
        for index in range(1, pivot_tables.Count + 1)
        if pivot_tables(index).TableRange2.Row <= 6
    ]
    summary_pivot = next(
        (pivot for pivot in top_pivots if pivot.TableRange2.Column == 1),
        None,
    )
    status_pivot = next(
        (pivot for pivot in top_pivots if pivot.TableRange2.Column > 1),
        None,
    )
    if summary_pivot is None or status_pivot is None:
        raise RuntimeError("Could not identify both TARGET COMPLETE PivotTables.")

    mask = pd.Series(True, index=dataframe.index)
    for header in (YEAR_FAB_UPLOAD_HEADER, next(
        (column for column in dataframe.columns if str(column).startswith("TARGET")),
        "",
    )):
        if not header:
            continue
        field = summary_pivot.PivotFields(header)
        visible_items = {
            str(item.Name).strip().casefold()
            for item in field.PivotItems()
            if item.Visible
        }
        if visible_items:
            normalized_values = dataframe[header].map(
                lambda value: "(blank)"
                if is_missing(value)
                else str(value).strip().casefold()
            )
            mask &= normalized_values.isin(visible_items)

    expected_count = int(
        dataframe.loc[mask, QUOTE_ID_HEADER].map(normalize_quote_id).ne("").sum()
    )
    summary_count = int(summary_pivot.GetPivotData(f"Count of {QUOTE_ID_HEADER}").Value)
    status_grand_total_row = status_pivot.TableRange2.Rows.Count
    status_values: list[int] = []
    for column in range(2, status_pivot.TableRange2.Columns.Count + 1):
        value = status_pivot.TableRange2.Cells(status_grand_total_row, column).Value
        if isinstance(value, numbers.Number):
            status_values.append(int(value))
    status_count = expected_count if expected_count in status_values else sum(status_values)

    if expected_count != summary_count or expected_count != status_count:
        raise RuntimeError(
            "TARGET COMPLETE count mismatch: "
            f"source={expected_count}, summary_pivot={summary_count}, "
            f"status_pivot={status_count}."
        )


def update_workbook_via_com(
    workbook_path: Path,
    df: pd.DataFrame,
) -> int:
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 and Microsoft Excel are required to preserve and refresh IDE PivotTables."
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    previous_calculation = None
    percentage_scratch = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        try:
            previous_calculation = excel.Calculation
            excel.Calculation = -4135  # xlCalculationManual
        except Exception:
            previous_calculation = None

        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False)
        all_order_sheet = find_sheet(workbook, ALL_ORDER_SHEET_NAME)
        pivot_sheet = find_sheet(workbook, PIVOT_SHEET_NAME)
        if all_order_sheet is None or pivot_sheet is None:
            raise RuntimeError("Previous workbook must contain ALL ORDER and PIVOT sheets.")

        row_count = len(df) + 1
        column_count = len(df.columns)
        table = all_order_sheet.ListObjects(TABLE_NAME)
        old_last_row = max(
            int(all_order_sheet.UsedRange.Rows.Count),
            int(table.Range.Rows.Count),
        )
        clear_last_row = max(old_last_row, row_count)
        all_order_sheet.Range(
            all_order_sheet.Cells(2, 1),
            all_order_sheet.Cells(clear_last_row, column_count),
        ).ClearContents()

        target_range = all_order_sheet.Range(
            all_order_sheet.Cells(1, 1),
            all_order_sheet.Cells(row_count, column_count),
        )
        table.Resize(target_range)
        table.TableStyle = TABLE_STYLE_NAME
        year_fab_upload_column = df.columns.get_loc(YEAR_FAB_UPLOAD_HEADER) + 1
        all_order_sheet.Range(
            all_order_sheet.Cells(2, year_fab_upload_column),
            all_order_sheet.Cells(row_count, year_fab_upload_column),
        ).NumberFormat = YEAR_TEXT_FORMAT
        body_values = tuple(
            tuple(com_value(value) for value in row)
            for row in df.itertuples(index=False, name=None)
        )
        all_order_sheet.Range(
            all_order_sheet.Cells(2, 1),
            all_order_sheet.Cells(row_count, column_count),
        ).Value = body_values
        for column_number, header in enumerate(df.columns, start=1):
            all_order_sheet.Cells(1, column_number).Value = str(header)
        apply_previous_status_na_errors(excel, all_order_sheet, df)
        if old_last_row > row_count:
            all_order_sheet.Rows(f"{row_count + 1}:{old_last_row}").Delete()

        target_range.Font.Name = FONT_NAME
        target_range.Font.Size = FONT_SIZE
        target_range.HorizontalAlignment = -4108  # xlCenter
        target_range.VerticalAlignment = -4108  # xlCenter

        otc_column = df.columns.get_loc(OTC_HEADER) + 1
        mrc_column = df.columns.get_loc(MRC_HEADER) + 1
        all_order_sheet.Range(
            all_order_sheet.Cells(2, otc_column),
            all_order_sheet.Cells(row_count, mrc_column),
        ).NumberFormat = "#,##0"

        all_order_sheet.Range(
            all_order_sheet.Cells(2, year_fab_upload_column),
            all_order_sheet.Cells(row_count, year_fab_upload_column),
        ).NumberFormat = YEAR_TEXT_FORMAT

        apply_date_formats(all_order_sheet, df)

        formula_column = df.columns.get_loc(ON_TIME_DELIVERY_HEADER) + 1
        actual_rfs_column = df.columns.get_loc(ACTUAL_RFS_DATE_HEADER) + 1
        initial_rfs_column = df.columns.get_loc(NEW_RFS_INITIAL_HEADER) + 1
        actual_offset = actual_rfs_column - formula_column
        initial_offset = initial_rfs_column - formula_column
        formula_values = []
        formula_rows: list[bool] = []
        for actual_rfs, initial_rfs in zip(
            df[ACTUAL_RFS_DATE_HEADER],
            df[NEW_RFS_INITIAL_HEADER],
            strict=False,
        ):
            has_required_dates = isinstance(actual_rfs, datetime) and isinstance(
                initial_rfs, datetime
            )
            formula_rows.append(has_required_dates)
            if has_required_dates:
                formula_values.append(
                    (
                        f'=IF(RC[{actual_offset}]>RC[{initial_offset}],'
                        '"DELAY","ONTIME")',
                    )
                )
            else:
                formula_values.append(("",))
        formula_range = all_order_sheet.Range(
            all_order_sheet.Cells(2, formula_column),
            all_order_sheet.Cells(row_count, formula_column),
        )
        formula_range.FormulaR1C1 = tuple(formula_values)
        blank_formula_rows = [
            row_number
            for row_number, has_formula in enumerate(formula_rows, start=2)
            if not has_formula
        ]
        if blank_formula_rows:
            start_row = previous_row = blank_formula_rows[0]
            for row_number in [*blank_formula_rows[1:], None]:
                if row_number is None or row_number != previous_row + 1:
                    all_order_sheet.Range(
                        all_order_sheet.Cells(start_row, formula_column),
                        all_order_sheet.Cells(previous_row, formula_column),
                    ).ClearContents()
                    if row_number is None:
                        break
                    start_row = row_number
                previous_row = row_number

        autofit_with_limits(all_order_sheet, column_count)
        all_order_sheet.Activate()
        excel.ActiveWindow.FreezePanes = False
        excel.ActiveWindow.SplitColumn = 2
        excel.ActiveWindow.SplitRow = 0
        excel.ActiveWindow.FreezePanes = True

        percentage_scratch, percentage_blocks = capture_pivot_percentage_blocks(
            workbook,
            pivot_sheet,
        )

        refreshed_count = 0
        pivot_tables = pivot_sheet.PivotTables()
        for pivot_index in range(1, pivot_tables.Count + 1):
            pivot_table = pivot_tables(pivot_index)
            try:
                pivot_cache = pivot_table.PivotCache()
                source_data = str(pivot_cache.SourceData)
                if TABLE_NAME.casefold() not in source_data.casefold():
                    raise RuntimeError(
                        f"PivotTable source is {source_data!r}, expected {TABLE_NAME}."
                    )
                pivot_cache.Refresh()
                pivot_table.RefreshTable()
                refreshed_count += 1
            except Exception as exc:
                raise RuntimeError(
                    f"Could not refresh PivotTable '{pivot_table.Name}': {exc}"
                ) from exc

        restore_pivot_percentage_blocks(
            pivot_sheet,
            percentage_scratch,
            percentage_blocks,
        )
        percentage_scratch.Visible = -1  # xlSheetVisible
        percentage_scratch.Delete()
        percentage_scratch = None
        repair_pivot_percentage_ranges(pivot_sheet)
        adjust_pivot_column_widths(pivot_sheet)
        validate_target_complete_pivots(pivot_sheet, df)

        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception:
            pass
        excel.Calculate()
        try:
            formula_errors = pivot_sheet.UsedRange.SpecialCells(-4123, 16)  # formulas, errors
            error_addresses = [
                area.Address for area in formula_errors.Areas
            ]
        except Exception:
            error_addresses = []
        if error_addresses:
            raise RuntimeError(
                "PIVOT contains formula errors after refresh: " + ", ".join(error_addresses)
            )
        workbook.Save()
        return refreshed_count
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                if previous_calculation is not None:
                    excel.Calculation = previous_calculation
            except Exception:
                pass
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def validate_output(
    output_path: Path,
    expected_quote_ids: list[str],
    expected_pivot_count: int,
) -> dict[str, object]:
    workbook = load_workbook(output_path, read_only=False, data_only=False, keep_links=False)
    try:
        missing_sheets = {
            ALL_ORDER_SHEET_NAME,
            PIVOT_SHEET_NAME,
        } - set(workbook.sheetnames)
        if missing_sheets:
            raise ValueError(f"Output is missing sheets: {', '.join(sorted(missing_sheets))}")

        worksheet = workbook[ALL_ORDER_SHEET_NAME]
        if worksheet.max_column != 61:
            raise ValueError(f"Output ALL ORDER has {worksheet.max_column} columns, expected 61.")
        actual_rows = worksheet.max_row - 1
        expected_rows = len(expected_quote_ids)
        if actual_rows != expected_rows:
            raise ValueError(f"Output has {actual_rows} data rows, expected {expected_rows}.")
        if worksheet.freeze_panes != "C1":
            raise ValueError(f"Output freeze pane is {worksheet.freeze_panes}, expected C1.")
        if TABLE_NAME not in worksheet.tables:
            raise ValueError(f"Output is missing Excel table: {TABLE_NAME}")
        expected_ref = f"A1:{get_column_letter(worksheet.max_column)}{worksheet.max_row}"
        if worksheet.tables[TABLE_NAME].ref != expected_ref:
            raise ValueError(
                f"{TABLE_NAME} range is {worksheet.tables[TABLE_NAME].ref}, expected {expected_ref}."
            )

        headers = {worksheet.cell(1, column).value: column for column in range(1, 62)}
        quote_column = headers[QUOTE_ID_HEADER]
        otc_column = headers[OTC_HEADER]
        mrc_column = headers[MRC_HEADER]
        year_fab_upload_column = headers[YEAR_FAB_UPLOAD_HEADER]
        actual_quote_ids: list[str] = []
        invalid_financial_values = 0
        invalid_year_values = 0
        for row_number in range(2, worksheet.max_row + 1):
            actual_quote_ids.append(
                normalize_quote_id(worksheet.cell(row_number, quote_column).value)
            )
            for column in (otc_column, mrc_column):
                if numeric_value(worksheet.cell(row_number, column).value) is None:
                    invalid_financial_values += 1
            year_cell = worksheet.cell(row_number, year_fab_upload_column)
            year_text = str(year_cell.value).strip()
            if (
                not re.fullmatch(r"\d{4}", year_text)
                or not MIN_VALID_DATE_YEAR <= int(year_text) <= MAX_VALID_DATE_YEAR
                or year_cell.number_format != YEAR_TEXT_FORMAT
            ):
                invalid_year_values += 1
        if actual_quote_ids != expected_quote_ids:
            mismatch_index = next(
                (
                    index
                    for index, (actual, expected) in enumerate(
                        zip(actual_quote_ids, expected_quote_ids, strict=False),
                        start=2,
                    )
                    if actual != expected
                ),
                None,
            )
            raise ValueError(
                "Output Quote ID sequence does not match the filtered raw data"
                + (f" at Excel row {mismatch_index}." if mismatch_index else ".")
            )
        if invalid_financial_values:
            raise ValueError(
                f"Output contains {invalid_financial_values} invalid OTC/MRC values."
            )
        if invalid_year_values:
            raise ValueError(
                f"Output contains {invalid_year_values} invalid or date-formatted "
                f"{YEAR_FAB_UPLOAD_HEADER} values."
            )

        pivot_count = len(workbook[PIVOT_SHEET_NAME]._pivots)
        if pivot_count != expected_pivot_count:
            raise ValueError(
                f"Output contains {pivot_count} PivotTables, expected {expected_pivot_count}."
            )
        return {
            "rows": actual_rows,
            "columns": worksheet.max_column,
            "pivot_tables": pivot_count,
            "table_ref": expected_ref,
        }
    finally:
        workbook.close()


def generate_ide_tracking(
    files: InputFiles,
    output_path: Path,
    report_date: date,
) -> dict[str, object]:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    raw = read_raw_dataframe(files.raw)
    previous = read_previous_dataframe(files.previous)
    build = build_all_order(raw, previous, report_date, files.previous)

    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix="ide_tracking_",
            suffix=".xlsx",
            dir=output_path.parent,
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)
        shutil.copy2(files.previous, temp_path)
        refreshed_count = update_workbook_via_com(temp_path, build.dataframe)
        validation = validate_output(
            temp_path,
            build.dataframe[QUOTE_ID_HEADER].map(normalize_quote_id).tolist(),
            refreshed_count,
        )
        try:
            os.replace(temp_path, output_path)
        except PermissionError as exc:
            raise PermissionError(
                f"Cannot replace output workbook because it is open or locked: {output_path}. "
                "Close the workbook in Excel and run the command again."
            ) from exc
        temp_path = None
    finally:
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass

    return {
        **validation,
        "refreshed_pivots": refreshed_count,
        "collabs_fallback_rows": 0,
        "zero_fallback_rows": build.zero_fallback_count,
        "duplicate_quote_ids": build.duplicate_quote_ids,
        "duplicate_rows_preserved": build.duplicate_rows_preserved,
        "date_fallback_rows": build.date_fallback_count,
        "field_fallback_rows": build.field_fallback_count,
        "invalid_dates": build.invalid_dates,
    }


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate Daily Tracking IDE from the IDE dashboard and previous workbook.",
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        default=Path(DEFAULT_INPUT_DIR),
        help=f"Raw IDE workbook or directory. Defaults to .\\{DEFAULT_INPUT_DIR}.",
    )
    parser.add_argument(
        "-r",
        "--reference",
        type=Path,
        default=Path(DEFAULT_REFERENCE_DIR),
        help=f"Previous Daily Tracking IDE workbook or directory. Defaults to .\\{DEFAULT_REFERENCE_DIR}.",
    )
    parser.add_argument(
        "-c",
        "--collabs",
        type=Path,
        help="Deprecated and ignored. OTC/MRC now use the previous workbook, then zero.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path(DEFAULT_OUTPUT_DIR),
        help=f"Output workbook or directory. Defaults to .\\{DEFAULT_OUTPUT_DIR}.",
    )
    parser.add_argument(
        "--report-date",
        type=parse_report_date,
        help="Explicit report date using YYYY-MM-DD. Auto-detected when omitted.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        files = resolve_inputs(args)
        report_date = determine_report_date(args.report_date, files.raw)
        output_path = resolve_output_path(args.output, report_date)
        print(f"[1/4] Reading raw data: {files.raw.name}")
        print(f"[2/4] Loading previous mapping and OTC/MRC: {files.previous.name}")
        if args.collabs is not None:
            print("[warn] --collabs is deprecated and ignored.", file=sys.stderr)
        print("[3/4] Updating ALL ORDER and refreshing PivotTables in Excel")
        result = generate_ide_tracking(files, output_path, report_date)
        print("[4/4] Validating output")
        print(
            f"Wrote {output_path} | rows={result['rows']} | columns={result['columns']} | "
            f"pivots={result['pivot_tables']}"
        )
        print(
            f"Fallbacks: OTC/MRC zero={result['zero_fallback_rows']}, "
            f"dates={result['date_fallback_rows']}, "
            f"fields={result['field_fallback_rows']}"
        )
        duplicate_quote_ids = result["duplicate_quote_ids"]
        if duplicate_quote_ids:
            print(
                f"Preserved {result['duplicate_rows_preserved']} repeated rows across "
                f"{len(duplicate_quote_ids)} Quote IDs; lookups are matched by occurrence."
            )
        invalid_dates = result["invalid_dates"]
        if invalid_dates:
            print("[warn] Invalid source date values were preserved as text:", file=sys.stderr)
            for header, values in invalid_dates.items():
                print(f"  - {header}: {', '.join(values)}", file=sys.stderr)
        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
