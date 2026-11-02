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

from completion_threshold_rules import (
    CompletionThresholds,
    STANDARD_PROFILE,
    WEEKLY_MODE,
    completion_thresholds_for,
)
from excel_pivot_layout import (
    EXCEL_MAX_ROW,
    PivotAlignmentModel,
    apply_completion_icon_thresholds,
    capture_pivot_alignments,
    completion_mode_for_row,
    completion_section_modes,
    normalize_layout_text,
    pivot_data_column_key,
    restore_pivot_alignments,
    resize_dynamic_pivot_section_banners,
)
from excel_pivot_filters import configure_pivot_cache_for_current_source


DEFAULT_INPUT_DIR = "input-ide"
DEFAULT_REFERENCE_DIR = "vlookup-ide"
DEFAULT_COLLABS_DIR = "collabs-ide"
DEFAULT_OUTPUT_DIR = "output-ide"

RAW_SHEET_NAME = "ICT ORDER 2026"
RAW_SOURCE_SHEET_NAME = "source"
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
SO_COMPLETION_DATE_HEADER = "SO COMPLETION DATE"
MONTH_COMPLETE_HEADER = "MONTH COMPLETE"
PRE_INSTALLATION_START_HEADER = "Pre-Installation Start Date"
PRE_INSTALLATION_AGING_HEADER = "Pre-Installation Aging"
PRE_INSTALLATION_RANGE_HEADER = "Range Aging On Pre-Installation"
DEPT_HEADER = "DEPT."
PM_HEADER = "PM"
APM_HEADER = "APM"

REQUIRED_RAW_HEADERS = (
    QUOTE_ID_HEADER,
    "QUOTE NAME",
    "PARTNERS",
    FAB_UPLOAD_HEADER,
    NEW_RFS_INITIAL_HEADER,
    ACTUAL_RFS_DATE_HEADER,
    AOS_SENT_DATE_HEADER,
    STATUS_DELIVERY_HEADER,
    SO_COMPLETION_DATE_HEADER,
)

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

# Legacy reference parity follows Excel's month-first text conversion. The
# display format is separate; typed dates are never regionally reparsed.
DEFAULT_RFS_TEXT_DATE_POLICY = "legacy-excel"
MONTH_FIRST_TEXT_DATE_HEADERS = (
    FAB_UPLOAD_HEADER,
)
DASHBOARD_RFS_DATE_HEADERS = (
    NEW_RFS_INITIAL_HEADER,
    ACTUAL_RFS_DATE_HEADER,
)
RAW_SOURCE_QUOTE_ID_COLUMN = 4  # D
RAW_SOURCE_NEW_RFS_COLUMN = 38  # AL

DATE_ONLY_FORMAT = "mm/dd/yy"
DATE_TIME_FORMAT = "mm/dd/yy h:mm AM/PM"
RFS_DATE_ONLY_FORMAT = "dd/mm/yyyy"
RFS_DATE_TIME_FORMAT = "dd/mm/yyyy h:mm AM/PM"
YEAR_TEXT_FORMAT = "@"
FONT_NAME = "Indosat Regular"
FONT_SIZE = 11
MAX_COLUMN_WIDTH = 45
MIN_COLUMN_WIDTH = 10
PIVOT_PERCENTAGE_MIN_WIDTH = 24
PIVOT_SECTION_TITLE_MARKERS = (
    "target complete",
    "order aging start from pre-installation status",
    "ontime delivery",
    "status order with aging from rfs commit date",
    "delay completion",
    "target after",
    "target not inputted",
)
PIVOT_SECTION_SIDE_HEADER_MARKERS = ("percentage of completion",)
UNKNOWN_YEAR = "1900"
EXCEL_ZERO_DATE_SERIAL = 0
EXCEL_NA_ERROR = "#N/A"
MIN_VALID_DATE_YEAR = 1900
MAX_VALID_DATE_YEAR = 2100
MISSING_TEXT = {"", "nan", "none", "null", "n/a", "#n/a", "nat"}
MAX_OPERATIONAL_FUTURE_YEARS = 5
REPORT_MONTH_NUMBERS = {
    **{month_name.casefold(): month_number for month_number, month_name in enumerate(calendar.month_name) if month_name},
    **{month_name.casefold(): month_number for month_number, month_name in enumerate(calendar.month_abbr) if month_name},
    "sept": 9,
    "januari": 1,
    "februari": 2,
    "maret": 3,
    "april": 4,
    "mei": 5,
    "juni": 6,
    "juli": 7,
    "agustus": 8,
    "september": 9,
    "oktober": 10,
    "november": 11,
    "desember": 12,
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "mei": 5,
    "jun": 6,
    "jul": 7,
    "agu": 8,
    "agt": 8,
    "sep": 9,
    "okt": 10,
    "nov": 11,
    "des": 12,
}
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
    collabs: Path


@dataclass(frozen=True)
class BuildResult:
    dataframe: pd.DataFrame
    invalid_dates: dict[str, list[str]]
    collabs_fallback_count: int
    zero_fallback_count: int
    duplicate_quote_ids: tuple[str, ...]
    duplicate_rows_preserved: int
    date_fallback_count: int
    field_fallback_count: int
    source_date_corrections: tuple[dict[str, object], ...] = ()
    financial_source_policy: str = "reference-snapshot"
    month_complete_format: str = "MMMM"
    reference_workflow: str = "continuation"
    historical_status_source: str = STATUS_DELIVERY_HEADER
    planned_rfs_year_rollovers: tuple[dict[str, object], ...] = ()


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


@dataclass(frozen=True)
class TargetPivotPageFilter:
    pivot_name: str
    position: int
    selection_kinds: tuple[str, ...]
    show_all: bool


@dataclass(frozen=True)
class DeliveryPivotPageFilter:
    pivot_name: str
    field_name: str
    selected_items: tuple[str, ...]
    show_all: bool
    is_page_field: bool = True


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


def validate_raw_headers(headers: Iterable[object]) -> None:
    names = pd.Index([normalize_header(header) for header in headers])
    duplicates = sorted(set(names[names.duplicated()]))
    if duplicates:
        raise ValueError("Raw sheet contains duplicate columns: " + ", ".join(duplicates))
    missing = sorted(set(REQUIRED_RAW_HEADERS).difference(names))
    if missing:
        raise ValueError("Raw sheet is missing required columns: " + ", ".join(missing))


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
        day_text, month_text, year_text = named_match.groups()
        month_number = REPORT_MONTH_NUMBERS.get(month_text.casefold())
        if month_number is not None:
            try:
                return date(int(year_text), month_number, int(day_text))
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
        collabs=resolve_single_file(args.collabs, ".csv", "collabs fallback"),
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


def recover_source_quote_date_years(raw: pd.DataFrame, worksheet) -> tuple[dict[str, object], ...]:
    """Correct an invalid year only with agreeing quote-created source evidence."""
    header = "QUOTE DATE"
    if header not in raw.columns:
        return ()
    invalid: dict[int, datetime] = {}
    for index, value in raw[header].items():
        if isinstance(value, (pd.Timestamp, datetime)):
            parsed = value.to_pydatetime() if isinstance(value, pd.Timestamp) else value
        elif isinstance(value, str):
            try:
                parsed = datetime.fromisoformat(value.strip())
            except ValueError:
                continue
        else:
            continue
        if not MIN_VALID_DATE_YEAR <= parsed.year <= MAX_VALID_DATE_YEAR:
            invalid[index] = parsed
    if not invalid:
        return ()
    quote_ids = raw[QUOTE_ID_HEADER].map(normalize_quote_id)
    needed = {quote_ids.at[index] for index in invalid}
    rows = worksheet.iter_rows(values_only=True)
    headers = {str(value).strip().casefold(): i for i, value in enumerate(next(rows, ())) if value is not None}
    if "quote_num" not in headers or "quote_created" not in headers:
        return ()
    candidates: dict[str, set[date]] = {}
    for row in rows:
        quote = normalize_quote_id(row[headers["quote_num"]])
        if quote not in needed:
            continue
        parsed = parse_excel_datetime(row[headers["quote_created"]])
        if parsed is not None:
            candidates.setdefault(quote, set()).add(parsed.date())
    corrections = []
    for index, original in invalid.items():
        quote = quote_ids.at[index]
        dates = candidates.get(quote, set())
        if len(dates) != 1:
            continue
        confirmed = next(iter(dates))
        if (original.month, original.day) != (confirmed.month, confirmed.day):
            continue
        corrected = original.replace(year=confirmed.year)
        raw.at[index, header] = corrected
        corrections.append({
            "quote_id": quote, "field": header,
            "original": original.isoformat(), "corrected": corrected.isoformat(),
            "evidence": f"{RAW_SOURCE_SHEET_NAME}.quote_created",
        })
    return tuple(corrections)


def read_raw_dataframe(
    raw_path: Path, *, previous: pd.DataFrame | None = None, reference_continuation: bool = True,
) -> pd.DataFrame:
    raw = pd.read_excel(
        raw_path,
        sheet_name=RAW_SHEET_NAME,
        header=0,
        skiprows=[1],
        dtype=object,
    )
    raw.columns = [normalize_header(column) for column in raw.columns]
    validate_raw_headers(raw.columns)

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

        raw.attrs["source_date_corrections"] = (
            recover_source_quote_date_years(raw, raw_workbook[RAW_SOURCE_SHEET_NAME])
            if RAW_SOURCE_SHEET_NAME in raw_workbook.sheetnames else ()
        )

        if NEW_RFS_INITIAL_HEADER in raw.columns:
            raw.attrs["rfs_original_text_rows"] = {
                header: tuple(
                    (key[0].casefold(), key[1])
                    for key, value in zip(occurrence_keys(raw), raw[header], strict=True)
                    if isinstance(value, str) and not is_missing(value)
                )
                for header in DASHBOARD_RFS_DATE_HEADERS if header in raw
            }
            quote_ids = raw[QUOTE_ID_HEADER].map(normalize_quote_id)
            unresolved_new_rfs = raw[NEW_RFS_INITIAL_HEADER].map(
                lambda value: (
                    not is_missing(value)
                    and not is_excel_zero_date(value)
                    and str(value).strip() != "-"
                    and (parse_dashboard_field_datetime(value, NEW_RFS_INITIAL_HEADER) is None
                         or (isinstance(value, str) and re.search(r"/{2,}", value) is not None))
                )
            )
            if previous is not None and reference_continuation:
                history = occurrence_lookup(previous, NEW_RFS_INITIAL_HEADER)
                history_values = pd.Series(
                    [history.get(key) for key in occurrence_keys(raw)], index=raw.index, dtype=object,
                )
                recovered_history = unresolved_new_rfs & history_values.map(
                    lambda value: is_excel_zero_date(value)
                    or parse_dashboard_field_datetime(value, NEW_RFS_INITIAL_HEADER) is not None
                )
                raw.loc[recovered_history, NEW_RFS_INITIAL_HEADER] = history_values.loc[recovered_history]
                raw.attrs["reference_date_recoveries"] = int(recovered_history.sum())
                unresolved_new_rfs &= ~recovered_history
            # A broken quote lookup can be recovered from current rows only
            # when every valid date for that quote agrees. Valid rows stay intact.
            quote_dates: dict[str, set[datetime]] = {}
            for quote_id, value in zip(quote_ids, raw[NEW_RFS_INITIAL_HEADER], strict=False):
                parsed = parse_dashboard_field_datetime(value, NEW_RFS_INITIAL_HEADER)
                if quote_id and parsed is not None:
                    quote_dates.setdefault(quote_id, set()).add(parsed)
            consistent_dates = {
                quote_id: next(iter(values))
                for quote_id, values in quote_dates.items() if len(values) == 1
            }
            recovered_current = unresolved_new_rfs & quote_ids.isin(consistent_dates)
            raw.loc[recovered_current, NEW_RFS_INITIAL_HEADER] = quote_ids.loc[
                recovered_current
            ].map(consistent_dates)
            unresolved_new_rfs &= ~recovered_current
            required_quote_ids = set(quote_ids.loc[unresolved_new_rfs])
            source_lookup = (
                load_source_new_rfs_lookup(raw_workbook[RAW_SOURCE_SHEET_NAME], required_quote_ids)
                if required_quote_ids and RAW_SOURCE_SHEET_NAME in raw_workbook.sheetnames
                else {}
            )
            source_values = quote_ids.map(source_lookup)
            has_source_value = unresolved_new_rfs & source_values.map(
                lambda value: is_excel_zero_date(value)
                or parse_dashboard_field_datetime(value, NEW_RFS_INITIAL_HEADER) is not None
            )
            raw.loc[has_source_value, NEW_RFS_INITIAL_HEADER] = quote_ids.loc[
                has_source_value
            ].map(source_lookup)
    finally:
        raw_workbook.close()

    quote_ids = raw[QUOTE_ID_HEADER].map(normalize_quote_id)
    raw = raw.loc[quote_ids.ne("")].copy()
    declared = [i for i, header in enumerate(raw.columns) if not str(header).startswith("Unnamed:")]
    if declared:
        raw = raw.iloc[:, :max(declared) + 1].copy()
    empty_export_columns = [
        header for header in raw.columns
        if str(header).startswith("Unnamed:") and raw[header].map(is_missing).all()
    ]
    raw = raw.drop(columns=empty_export_columns)
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
    workbook = load_workbook(previous_path, read_only=True, data_only=True, keep_links=False)
    try:
        for index, row in enumerate(workbook[ALL_ORDER_SHEET_NAME].iter_rows(min_row=2, max_col=len(previous.columns))):
            if index >= len(previous):
                break
            for column, cell in enumerate(row):
                if cell.data_type == "e":
                    previous.iat[index, column] = str(cell.value)
    finally:
        workbook.close()
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


def exact_quote_key(value: object) -> str:
    """Excel's exact text lookup is case-insensitive but preserves whitespace."""
    return "" if is_missing(value) else str(value).casefold()


def first_quote_lookup(
    previous: pd.DataFrame, value_header: str, *, normalize_keys: bool = True,
) -> dict[str, object]:
    if value_header not in previous.columns:
        return {}
    lookup: dict[str, object] = {}
    for quote_id, value in zip(
        previous[QUOTE_ID_HEADER].map(normalize_quote_id if normalize_keys else exact_quote_key),
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
    for header in headers:
        if header not in current.columns or header not in previous.columns:
            continue
        previous_lookup = first_quote_lookup(previous, header)
        values: list[object] = []
        for quote_id, current_value in zip(
            current[QUOTE_ID_HEADER].map(normalize_quote_id),
            current[header],
            strict=False,
        ):
            previous_value = previous_lookup.get(quote_id)
            if is_missing(current_value) and not is_missing(previous_value):
                values.append(previous_value)
                fallback_count += 1
            elif is_missing(current_value):
                values.append(current_value)
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


def load_collabs_lookup(
    collabs_path: Path,
    required_quote_ids: set[str],
) -> dict[str, tuple[int | float | None, int | float | None]]:
    """Load only OTC/MRC rows needed as fallback from the large collabs CSV."""
    if not required_quote_ids:
        return {}

    lookup: dict[str, tuple[int | float | None, int | float | None]] = {}
    required_headers = {"quote_num", "otc", "mrc"}
    reader = pd.read_csv(
        collabs_path,
        usecols=lambda column: str(column).strip().lower() in required_headers,
        dtype="string",
        chunksize=200_000,
        low_memory=False,
    )
    with reader:
        for chunk in reader:
            chunk.columns = [str(column).strip().lower() for column in chunk.columns]
            missing_headers = required_headers.difference(chunk.columns)
            if missing_headers:
                raise ValueError(
                    "Collabs CSV is missing required columns: "
                    + ", ".join(sorted(missing_headers))
                )

            chunk["quote_num"] = chunk["quote_num"].map(normalize_quote_id)
            matches = chunk.loc[chunk["quote_num"].isin(required_quote_ids)]
            for row in matches.itertuples(index=False):
                quote_id = normalize_quote_id(row.quote_num)
                if quote_id and quote_id not in lookup:
                    lookup[quote_id] = (numeric_value(row.otc), numeric_value(row.mrc))
            if required_quote_ids.issubset(lookup):
                break

    return lookup


def resolve_financial_values(
    previous_otc: object,
    previous_mrc: object,
    collabs_values: tuple[object, object] = (None, None),
) -> tuple[int | float, int | float, bool, bool]:
    """Resolve OTC/MRC using previous workbook, collabs CSV, then zero."""
    otc_value = numeric_value(previous_otc)
    mrc_value = numeric_value(previous_mrc)
    collabs_otc = numeric_value(collabs_values[0])
    collabs_mrc = numeric_value(collabs_values[1])
    used_collabs = False
    used_zero = False

    if otc_value is None and collabs_otc is not None:
        otc_value = collabs_otc
        used_collabs = True
    if mrc_value is None and collabs_mrc is not None:
        mrc_value = collabs_mrc
        used_collabs = True
    if otc_value is None:
        otc_value = 0
        used_zero = True
    if mrc_value is None:
        mrc_value = 0
        used_zero = True

    return otc_value, mrc_value, used_collabs, used_zero


def parse_excel_datetime(
    value: object, *, dayfirst: bool = True, allow_alternate_order: bool = True,
) -> datetime | None:
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
    fallback_dayfirst = dayfirst

    if not allow_alternate_order:
        year_first_date = re.fullmatch(
            r"(\d{4})([/.-])(\d{1,2})\2(\d{1,2})((?:[ T].*)?)", text,
        )
        numeric_date = re.fullmatch(
            r"(\d{1,2})([/.-])(\d{1,2})\2(\d{2,4})((?:[ T].*)?)", text,
        )
        if year_first_date is not None:
            year, _, month, day, suffix = year_first_date.groups()
        elif numeric_date is not None:
            first, _, second, year, suffix = numeric_date.groups()
            day, month = (first, second) if dayfirst else (second, first)
        if year_first_date is not None or numeric_date is not None:
            try:
                if len(year) != 4:
                    return None
                datetime(int(year), int(month), int(day))
            except ValueError:
                return None
            # Validate the declared order before dateutil can reinterpret an
            # invalid month as a day. ISO also preserves any time suffix.
            text = f"{int(year):04d}-{int(month):02d}-{int(day):02d}{suffix}"
            fallback_dayfirst = False

    day_first_formats = (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
    )
    month_first_formats = (
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
    )
    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        *(day_first_formats if dayfirst else month_first_formats),
        *((month_first_formats if dayfirst else day_first_formats) if allow_alternate_order else ()),
    )
    for fmt in formats:
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed if MIN_VALID_DATE_YEAR <= parsed.year <= MAX_VALID_DATE_YEAR else None
        except ValueError:
            continue

    try:
        parsed = date_parser.parse(text, dayfirst=fallback_dayfirst, fuzzy=False)
        return parsed if MIN_VALID_DATE_YEAR <= parsed.year <= MAX_VALID_DATE_YEAR else None
    except (OverflowError, TypeError, ValueError):
        return None


def parse_dashboard_field_datetime(
    value: object, header: str, *, rfs_text_policy: str = DEFAULT_RFS_TEXT_DATE_POLICY,
) -> datetime | None:
    if rfs_text_policy not in {"legacy-excel", "day-first"}:
        raise ValueError(f"Unknown RFS text date policy: {rfs_text_policy}")
    is_rfs = header in DASHBOARD_RFS_DATE_HEADERS
    month_first = header in MONTH_FIRST_TEXT_DATE_HEADERS or (is_rfs and rfs_text_policy == "legacy-excel")
    return parse_excel_datetime(
        value,
        dayfirst=not month_first,
        allow_alternate_order=not is_rfs or rfs_text_policy == "legacy-excel",
    )


def load_source_new_rfs_lookup(worksheet, required_quote_ids: set[str]) -> dict[str, object]:
    """Reproduce the raw workbook's VLOOKUP(D:AL, 35, FALSE) first-match rule."""
    if not required_quote_ids:
        return {}

    lookup: dict[str, object] = {}
    for row in worksheet.iter_rows(
        min_col=RAW_SOURCE_QUOTE_ID_COLUMN,
        max_col=RAW_SOURCE_NEW_RFS_COLUMN,
        values_only=True,
    ):
        quote_id = normalize_quote_id(row[0])
        if not quote_id or quote_id not in required_quote_ids or quote_id in lookup:
            continue

        source_value = row[RAW_SOURCE_NEW_RFS_COLUMN - RAW_SOURCE_QUOTE_ID_COLUMN]
        parsed = parse_dashboard_field_datetime(source_value, NEW_RFS_INITIAL_HEADER)
        lookup[quote_id] = (
            datetime.combine(parsed.date(), time.min)
            if parsed is not None
            else source_value
        )
        if len(lookup) == len(required_quote_ids):
            break

    return lookup


def normalize_dashboard_text_dates(df: pd.DataFrame) -> None:
    """Apply each source field's date contract before lookups and metrics."""
    for header in (*MONTH_FIRST_TEXT_DATE_HEADERS, *DASHBOARD_RFS_DATE_HEADERS):
        if header not in df.columns:
            continue

        values: list[object] = []
        for value in df[header]:
            if not isinstance(value, str) or is_missing(value) or value.strip() == "-":
                values.append(value)
                continue

            parsed = parse_dashboard_field_datetime(value, header)
            values.append(parsed if parsed is not None else value)
        df[header] = pd.Series(values, index=df.index, dtype=object)


def normalize_date_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, list[str]]]:
    result = df.copy()
    invalid_dates: dict[str, list[str]] = {}
    for header in DATE_HEADERS:
        if header not in result.columns:
            continue
        normalized_values: list[object] = []
        invalid_values: list[str] = []
        for value in result[header]:
            parsed = parse_dashboard_field_datetime(value, header)
            if is_excel_zero_date(value):
                normalized_values.append(EXCEL_ZERO_DATE_SERIAL)
            elif parsed is not None:
                normalized_values.append(parsed)
            elif is_missing(value) or str(value).strip() == "-":
                normalized_values.append(None)
            else:
                normalized_values.append(
                    None if header in DASHBOARD_RFS_DATE_HEADERS else
                    value if isinstance(value, (datetime, date, pd.Timestamp)) else str(value).strip()
                )
                invalid_values.append(str(value).strip())
        result[header] = pd.Series(normalized_values, index=result.index, dtype=object)
        if invalid_values:
            invalid_dates[header] = sorted(set(invalid_values))
    return result, invalid_dates


def target_period(report_date: date) -> date:
    """Return the business month used by the target buckets.

    The first calendar-day snapshot still closes the preceding month. From the
    second day onward, the report uses the current month.
    """
    if report_date.day == 1:
        return report_date - timedelta(days=1)
    return report_date


def target_value(value: object, period: date) -> str:
    parsed = parse_dashboard_field_datetime(value, NEW_RFS_INITIAL_HEADER)
    month_name = calendar.month_name[period.month]
    if parsed is None:
        return "Target Not Yet Inputted"
    source_month = (parsed.year, parsed.month)
    target_month = (period.year, period.month)
    if source_month < target_month:
        return f"Before {month_name}"
    if source_month > target_month:
        return f"After {month_name}"
    return f"Target {month_name}"


def target_item_kind(value: object) -> str | None:
    normalized = " ".join(str(value).strip().casefold().split())
    if normalized in {"(all)", "all"}:
        return "all"
    if "not yet inputted" in normalized:
        return "missing"
    if normalized.startswith("before "):
        return "before"
    if normalized.startswith("after "):
        return "after"
    if normalized.startswith("target "):
        return "target"
    return None


def target_item_label(kind: str, period: date) -> str:
    month_name = calendar.month_name[period.month]
    labels = {
        "before": f"Before {month_name}",
        "target": f"Target {month_name}",
        "after": f"After {month_name}",
        "missing": "Target Not Yet Inputted",
    }
    try:
        return labels[kind]
    except KeyError as exc:
        raise ValueError(f"Unknown target item kind: {kind}") from exc


def target_value_with_previous(
    current_value: object,
    previous_target_value: object,
    period: date,
) -> str:
    """Match the manual carry-forward behavior within one target period."""
    previous_kind = target_item_kind(previous_target_value)
    period_month = calendar.month_name[period.month].casefold()
    previous_target_text = str(previous_target_value).strip().casefold()
    same_period = previous_kind == "missing" or previous_target_text.endswith(period_month)

    if same_period and previous_kind not in {None, "missing", "all"}:
        return target_item_label(previous_kind, period)

    return target_value(current_value, period)


def is_plausible_operational_date(value: datetime | None, report_date: date) -> bool:
    if value is None:
        return False
    if value.year == MIN_VALID_DATE_YEAR:
        return True
    return MIN_VALID_DATE_YEAR < value.year <= report_date.year + MAX_OPERATIONAL_FUTURE_YEARS


def reconciled_date_value(
    current_value: object,
    previous_value: object,
    report_date: date,
    *,
    header: str = "",
    preserve_reference_year: bool = True,
) -> tuple[object, bool]:
    if is_excel_zero_date(current_value):
        return EXCEL_ZERO_DATE_SERIAL, False

    current_date = parse_dashboard_field_datetime(current_value, header)
    previous_date = parse_dashboard_field_datetime(previous_value, header)

    if is_missing(current_value) or str(current_value).strip() == "-":
        return current_value, False

    if (
        preserve_reference_year and current_date is not None
        and is_plausible_operational_date(previous_date, report_date)
        and current_date.year != previous_date.year
        and (current_date.month, current_date.day) == (previous_date.month, previous_date.day)
        and current_date.replace(year=previous_date.year) == previous_date
    ):
        return previous_date, True
    if is_plausible_operational_date(current_date, report_date):
        return current_date, False
    if is_excel_zero_date(previous_value):
        return EXCEL_ZERO_DATE_SERIAL, True
    if is_plausible_operational_date(previous_date, report_date):
        return previous_date, True
    return current_value, False


def reconcile_date_column(
    current: pd.DataFrame,
    previous: pd.DataFrame,
    header: str,
    report_date: date,
    *,
    reference_year_keys: set[tuple[str, int]] | None = None,
) -> tuple[pd.Series, int]:
    """Reconcile dates by occurrence under the declared reference-year policy."""
    previous_lookup = occurrence_lookup(previous, header)
    reconciled_values: list[object] = []
    fallback_count = 0

    for current_value, occurrence_key in zip(
        current[header],
        occurrence_keys(current),
        strict=False,
    ):
        reconciled_value, used_fallback = reconciled_date_value(
            current_value,
            previous_lookup.get(occurrence_key),
            report_date,
            header=header,
            preserve_reference_year=(
                reference_year_keys is None
                or (occurrence_key[0].casefold(), occurrence_key[1]) in reference_year_keys
            ),
        )
        reconciled_values.append(reconciled_value)
        fallback_count += int(used_fallback)

    return pd.Series(reconciled_values, index=current.index, dtype=object), fallback_count


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


def planned_rfs_with_completion_year(
    initial_value: object, completion_value: object, report_date: date, *, is_text_source: bool,
) -> tuple[object, bool]:
    """Recover legacy text years without turning historical RFS into future plans."""
    initial = parse_dashboard_field_datetime(initial_value, NEW_RFS_INITIAL_HEADER)
    completion = parse_dashboard_field_datetime(completion_value, SO_COMPLETION_DATE_HEADER)
    if not is_text_source or initial is None or completion is None or initial.year >= completion.year:
        return initial_value, False
    try:
        corrected = initial.replace(year=completion.year)
    except ValueError:
        # A leap-day commit cannot be moved into a non-leap year unambiguously.
        return initial_value, False
    if corrected.date() > report_date:
        return initial_value, False
    return corrected, True


def insert_after(df: pd.DataFrame, after_header: str, header: str, values: Iterable[object]) -> None:
    if after_header not in df.columns:
        raise ValueError(f"Could not insert {header}; missing anchor column: {after_header}")
    position = df.columns.get_loc(after_header) + 1
    df.insert(position, header, list(values))


def build_all_order(
    raw: pd.DataFrame,
    previous: pd.DataFrame,
    collabs_path: Path,
    report_date: date,
    previous_path: Path,
) -> BuildResult:
    validate_raw_headers(raw.columns)
    result = raw.copy().reset_index(drop=True)
    previous = previous.copy().reset_index(drop=True)
    prior_report_date = date_from_filename(previous_path)
    continuation = delivery_pivot_policy(prior_report_date, report_date) == "existing-period"
    # Daily carry-forward and a new-period rebuild have different legacy lookup contracts.
    raw_ancillary_dates = {
        header: result[header].copy() for header in DATE_HEADERS
        if header in result and header not in (*DASHBOARD_RFS_DATE_HEADERS, FAB_UPLOAD_HEADER, SO_COMPLETION_DATE_HEADER, "QUOTE DATE", "SO DATE")
    }
    original_text_rows = raw.attrs.get("rfs_original_text_rows", {})
    rfs_text_rows = {
        header: set(original_text_rows.get(header, ())) | {
            (key[0].casefold(), key[1])
            for key, value in zip(occurrence_keys(result), result[header], strict=True)
            if isinstance(value, str) and not is_missing(value)
        }
        for header in DASHBOARD_RFS_DATE_HEADERS
    }
    source_month = result.pop(MONTH_COMPLETE_HEADER) if MONTH_COMPLETE_HEADER in result else None
    previous_months = previous.get(MONTH_COMPLETE_HEADER, pd.Series(dtype=object))
    abbreviated_months = previous_months.map(
        lambda value: isinstance(value, str) and re.fullmatch(r"[A-Za-z]{3}-\d{2}", value.strip()) is not None
    ).any()
    month_format = "MMM-YY" if not continuation or abbreviated_months else "MMMM"
    if continuation:
        historical_ids: dict[str, object] = {}
        for value in previous[QUOTE_ID_HEADER]:
            historical_ids.setdefault(normalize_quote_id(value).casefold(), value)
        result[QUOTE_ID_HEADER] = pd.Series([
            historical_ids.get(normalize_quote_id(value).casefold(), value) for value in result[QUOTE_ID_HEADER]
        ], index=result.index, dtype=object)

        if "NOMOR PO" in result:
            result["NOMOR PO"] = result["NOMOR PO"].map(
                lambda value: int(value.strip())
                if isinstance(value, str) and re.fullmatch(r"(?:0|[1-9]\d{0,14})", value.strip())
                else value
            )
    normalize_dashboard_text_dates(result)
    quote_ids = result[QUOTE_ID_HEADER].map(normalize_quote_id)
    duplicate_mask = quote_ids.duplicated(keep=False)
    duplicate_quote_ids = tuple(sorted(quote_ids.loc[duplicate_mask].unique()))
    duplicate_rows_preserved = int(quote_ids.duplicated(keep="first").sum())
    field_fallback_count = apply_previous_field_fallbacks(
        result,
        previous,
        PREVIOUS_FALLBACK_HEADERS,
    )
    date_fallback_count = int(raw.attrs.get("reference_date_recoveries", 0))

    for header in DATE_HEADERS:
        if header not in result.columns:
            continue
        result[header], fallback_count = reconcile_date_column(
            result,
            previous,
            header,
            report_date,
            reference_year_keys=(
                None if continuation or header == FAB_UPLOAD_HEADER else rfs_text_rows.get(header, set())
            ),
        )
        date_fallback_count += fallback_count

    planned_year_rollovers = []
    if not continuation:
        for index, key in enumerate(occurrence_keys(result)):
            original = result.at[index, NEW_RFS_INITIAL_HEADER]
            value, changed = planned_rfs_with_completion_year(
                original, result.at[index, SO_COMPLETION_DATE_HEADER], report_date,
                is_text_source=(key[0].casefold(), key[1]) in rfs_text_rows[NEW_RFS_INITIAL_HEADER],
            )
            if changed:
                result.at[index, NEW_RFS_INITIAL_HEADER] = value
                planned_year_rollovers.append({
                    "quote_id": key[0], "occurrence": key[1],
                    "field": NEW_RFS_INITIAL_HEADER, "original": str(original),
                    "corrected": str(value), "evidence": SO_COMPLETION_DATE_HEADER,
                })

    if continuation:
        first_actual_dates = first_quote_lookup(result, ACTUAL_RFS_DATE_HEADER)
        result[ACTUAL_RFS_DATE_HEADER] = pd.Series([
            value if is_missing(value) else first_actual_dates.get(quote_id, value)
            for quote_id, value in zip(quote_ids, result[ACTUAL_RFS_DATE_HEADER], strict=True)
        ], index=result.index, dtype=object)

    order_type_lookup = first_quote_lookup(previous, ORDER_TYPE_HEADER)
    reference_financials = all(header in previous.columns for header in (OTC_HEADER, MRC_HEADER))
    otc_lookup = first_quote_lookup(previous, OTC_HEADER, normalize_keys=not reference_financials)
    mrc_lookup = first_quote_lookup(previous, MRC_HEADER, normalize_keys=not reference_financials)
    baseline_headers = [header for header in previous if header.startswith(f"{STATUS_DELIVERY_HEADER} ")]
    if len(baseline_headers) > 1:
        raise ValueError("Previous IDE workbook has multiple historical status baselines.")
    baseline_header = baseline_headers[0] if baseline_headers else STATUS_DELIVERY_HEADER
    previous_status_lookup = first_quote_lookup(previous, baseline_header, normalize_keys=False)
    status_quote_keys = {exact_quote_key(value) for value in previous[QUOTE_ID_HEADER]}
    previous_quote_ids = set(previous[QUOTE_ID_HEADER].map(normalize_quote_id))
    pre_installation_start_lookup = first_quote_lookup(
        previous, PRE_INSTALLATION_START_HEADER, normalize_keys=False,
    )
    previous_current_status = first_quote_lookup(previous, STATUS_DELIVERY_HEADER)

    reference_financial_keys = {exact_quote_key(value) for value in previous[QUOTE_ID_HEADER]}
    if not reference_financials:
        missing_financial_ids = {
            quote_id for quote_id in quote_ids
            if numeric_value(otc_lookup.get(quote_id)) is None
            or numeric_value(mrc_lookup.get(quote_id)) is None
        }
    elif continuation:
        missing_financial_ids = {
            quote_id for quote_id, original in zip(quote_ids, result[QUOTE_ID_HEADER], strict=True)
            if exact_quote_key(original) not in reference_financial_keys
        }
    else:
        missing_financial_ids = set()
    collabs_lookup = load_collabs_lookup(collabs_path, missing_financial_ids) if missing_financial_ids else {}

    order_types: list[object] = []
    otc_values: list[object] = []
    mrc_values: list[object] = []
    previous_status_values: list[object] = []
    pre_installation_start_values: list[object] = []
    collabs_fallback_count = 0
    zero_fallback_count = 0

    current_status_values = result[STATUS_DELIVERY_HEADER].map(normalize_status)
    for quote_id, raw_quote_id, current_status in zip(
        quote_ids,
        result[QUOTE_ID_HEADER],
        current_status_values,
        strict=False,
    ):
        previous_order_type = order_type_lookup.get(quote_id)
        order_types.append(
            "New Registration" if is_missing(previous_order_type) else str(previous_order_type).strip()
        )

        financial_key = exact_quote_key(raw_quote_id) if reference_financials else quote_id
        previous_otc, previous_mrc, used_collabs, used_zero = resolve_financial_values(
            otc_lookup.get(financial_key),
            mrc_lookup.get(financial_key),
            collabs_lookup.get(quote_id, (None, None))
            if not reference_financials or (continuation and financial_key not in reference_financial_keys)
            else (None, None),
        )

        otc_values.append(previous_otc)
        mrc_values.append(previous_mrc)
        collabs_fallback_count += int(used_collabs)
        zero_fallback_count += int(used_zero)

        status_key = exact_quote_key(raw_quote_id)
        if status_key not in status_quote_keys:
            previous_status_values.append(EXCEL_NA_ERROR)
        else:
            previous_status = previous_status_lookup.get(status_key)
            previous_status_values.append(
                previous_status if str(previous_status).strip() == EXCEL_NA_ERROR
                else (0 if is_missing(previous_status) else previous_status)
            )

        previous_start = pre_installation_start_lookup.get(exact_quote_key(raw_quote_id))
        previous_start_date = parse_excel_datetime(previous_start)
        placeholder = (
            continuation and is_missing(current_status)
            and is_missing(previous_current_status.get(quote_id))
            and previous_start_date is not None and prior_report_date is not None
            and previous_start_date.date() == prior_report_date - timedelta(days=1)
        )
        has_pre_installation_event = str(current_status).casefold() in {
            "04-pre-installation", "06-installation",
        }
        active_order = not is_missing(current_status) and str(current_status).casefold() not in {"so complete", "cancel"}
        initialize_start = (
            quote_id not in previous_quote_ids or (is_missing(previous_start) and active_order) or placeholder
        ) if continuation else (is_missing(previous_start) and has_pre_installation_event)
        if initialize_start:
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
    target_bucket_period = target_period(report_date)
    previous_target_header = next(
        (
            column
            for column in previous.columns
            if normalize_header(column).casefold().startswith("target ")
        ),
        None,
    )
    previous_target_lookup = (
        occurrence_lookup(previous, previous_target_header)
        if previous_target_header is not None
        else {}
    )
    target_values = [
        target_value_with_previous(
            current_value,
            previous_target_lookup.get(occurrence_key),
            target_bucket_period,
        )
        for current_value, occurrence_key in zip(
            result[NEW_RFS_INITIAL_HEADER],
            occurrence_keys(result),
            strict=False,
        )
    ]
    insert_after(result, NEW_RFS_INITIAL_HEADER, dynamic_target_header, target_values)

    actual_rfs_dates = result[ACTUAL_RFS_DATE_HEADER].map(
        lambda value: parse_dashboard_field_datetime(value, ACTUAL_RFS_DATE_HEADER)
    )
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

    dynamic_previous_status_header = (
        baseline_header if continuation and baseline_headers else previous_status_header(previous_path, report_date)
    )
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
    if not continuation:
        for header, values in raw_ancillary_dates.items():
            for index, original in values.items():
                if isinstance(original, (str, datetime, date, pd.Timestamp)):
                    result.at[index, header] = original
    result = result.reset_index(drop=True)
    if MONTH_COMPLETE_HEADER in result.columns:
        result.pop(MONTH_COMPLETE_HEADER)
    insert_after(result, SO_COMPLETION_DATE_HEADER, MONTH_COMPLETE_HEADER, [pd.NA] * len(result))
    if source_month is not None and month_format == "MMM-YY":
        suffix = 2
        while f"{MONTH_COMPLETE_HEADER}{suffix}" in result.columns:
            suffix += 1
        result[f"{MONTH_COMPLETE_HEADER}{suffix}"] = source_month.reset_index(drop=True)

    return BuildResult(
        dataframe=result,
        invalid_dates=invalid_dates,
        collabs_fallback_count=collabs_fallback_count,
        zero_fallback_count=zero_fallback_count,
        duplicate_quote_ids=duplicate_quote_ids,
        duplicate_rows_preserved=duplicate_rows_preserved,
        date_fallback_count=date_fallback_count,
        field_fallback_count=field_fallback_count,
        source_date_corrections=tuple(raw.attrs.get("source_date_corrections", ())),
        financial_source_policy=("reference-and-new-row-collabs" if continuation else "reference-snapshot") if reference_financials else "collabs-fallback",
        month_complete_format=month_format,
        reference_workflow="continuation" if continuation else "period-refresh",
        historical_status_source=baseline_header,
        planned_rfs_year_rollovers=tuple(planned_year_rollovers),
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


def contiguous_row_ranges(rows: list[int]) -> Iterable[tuple[int, int]]:
    if not rows:
        return
    start = previous = rows[0]
    for row_number in rows[1:]:
        if row_number != previous + 1:
            yield start, previous
            start = row_number
        previous = row_number
    yield start, previous


def apply_text_formats(worksheet, df: pd.DataFrame) -> None:
    for column_number, header in enumerate(df.columns, start=1):
        text_rows = [
            row_number for row_number, value in enumerate(df[header], start=2)
            if isinstance(value, str)
        ]
        column_letter = get_column_letter(column_number)
        for start_row, end_row in contiguous_row_ranges(text_rows):
            worksheet.Range(f"{column_letter}{start_row}:{column_letter}{end_row}").NumberFormat = "@"


def apply_month_complete_formula(worksheet, df: pd.DataFrame, *, date_format: str = "MMMM") -> None:
    month_column = df.columns.get_loc(MONTH_COMPLETE_HEADER) + 1
    completion_column = df.columns.get_loc(SO_COMPLETION_DATE_HEADER) + 1
    month_range = worksheet.Range(
        worksheet.Cells(2, month_column), worksheet.Cells(len(df) + 1, month_column)
    )
    month_range.NumberFormat = "General"
    # Full-month templates retain their legacy January-for-empty TEXT behavior.
    offset = completion_column - month_column
    if date_format == "MMM-YY":
        month_range.FormulaR1C1 = f'=IF(ISNUMBER(RC[{offset}]),TEXT(RC[{offset}],"[$-409]mmm-yy"),"")'
    elif date_format == "MMMM":
        month_range.FormulaR1C1 = f'=TEXT(RC[{offset}],"[$-409]mmmm")'
    else:
        raise ValueError(f"Unsupported completion month format: {date_format}")
    month_range.Calculate()


def set_pivot_page_field_position(pivot_table, field, position: int) -> None:
    page_count = int(pivot_table.PageFields.Count)
    if page_count < 1:
        raise RuntimeError(f"PivotTable '{pivot_table.Name}' has no page filters.")
    # A source-header rename can remove a page field during cache refresh.
    valid_position = min(max(1, position), page_count)
    try:
        if int(field.Position) != valid_position:
            field.Position = valid_position
    except Exception as exc:
        raise RuntimeError(
            f"Could not position page filter '{field.Name}' in PivotTable "
            f"'{pivot_table.Name}' at {valid_position} of {page_count}: {exc}"
        ) from exc


def delivery_pivot_policy(previous_report_date: date | None, report_date: date) -> str:
    """Keep an existing period's selections; start a new period cumulatively."""
    if previous_report_date is not None:
        previous_period = target_period(previous_report_date)
        current_period = target_period(report_date)
        if (previous_period.year, previous_period.month) == (current_period.year, current_period.month):
            return "existing-period"
    return "cumulative-active"


def delivery_pivot_fields(pivot) -> tuple[str, ...]:
    columns = {normalize_header(field.Name) for field in pivot.ColumnFields}
    if ON_TIME_DELIVERY_HEADER in columns:
        return MONTH_COMPLETE_HEADER, STATUS_DELIVERY_HEADER
    if STATUS_ORDER_HEADER in columns:
        return (STATUS_DELIVERY_HEADER,)
    return ()


def capture_delivery_pivot_filters(worksheet) -> tuple[DeliveryPivotPageFilter, ...]:
    states = []
    for pivot in worksheet.PivotTables():
        names = delivery_pivot_fields(pivot)
        if not names:
            continue
        fields = {normalize_header(field.Name): field for field in pivot.PivotFields()}
        for name in names:
            field = fields.get(name)
            if field is None:
                raise RuntimeError(f"PivotTable '{pivot.Name}' is missing {name}.")
            items = tuple(field.PivotItems())
            selected = tuple(str(item.Name) for item in items if item.Visible)
            if int(field.Orientation) == 3 and not field.EnableMultiplePageItems:
                current_page = str(field.CurrentPage)
                selected = (
                    tuple(str(item.Name) for item in items)
                    if current_page == "(All)" else (current_page,)
                )
            states.append(DeliveryPivotPageFilter(
                str(pivot.Name), name, selected, len(selected) == len(items),
                int(field.Orientation) == 3,
            ))
    return tuple(states)


def set_delivery_pivot_selection(field, selected: tuple[str, ...] | None) -> None:
    field.ClearAllFilters()
    # Excel can report every item hidden in single-select (All) mode. Use
    # explicit checkboxes for verifiable selections, including all items.
    field.EnableMultiplePageItems = True
    items = tuple(field.PivotItems())
    for item in items:
        try:
            item.Visible = True
        except Exception as exc:
            raise RuntimeError(
                f"Could not show item '{item.Name}' in pivot field '{field.Name}': {exc}"
            ) from exc
    if selected is not None:
        allowed = {name.casefold() for name in selected}
        present = {str(item.Name).casefold() for item in items}
        if not allowed.intersection(present):
            raise RuntimeError(f"Pivot field '{field.Name}' has none of its selected items: {selected}.")
        for item in items:
            if str(item.Name).casefold() not in allowed:
                item.Visible = False
    actual = {str(item.Name).casefold() for item in items if item.Visible}
    expected = {str(item.Name).casefold() for item in items} if selected is None else allowed & present
    if actual != expected:
        raise RuntimeError(f"Pivot field '{field.Name}' did not retain its item selection.")


def configure_delivery_pivot_filters(
    worksheet,
    states: tuple[DeliveryPivotPageFilter, ...],
    policy: str,
    report_date: date,
) -> None:
    if policy not in {"existing-period", "cumulative-active"}:
        raise ValueError(f"Unknown delivery pivot policy: {policy}")
    saved = {(state.pivot_name, state.field_name): state for state in states}
    pivots = worksheet.PivotTables()
    for index in range(1, pivots.Count + 1):
        pivot = pivots(index)
        names = delivery_pivot_fields(pivot)
        for name in names:
            field = next(
                (field for field in pivot.PivotFields() if normalize_header(field.Name) == name),
                None,
            )
            if field is None:
                raise RuntimeError(f"PivotTable '{pivot.Name}' is missing {name}.")
            if name == MONTH_COMPLETE_HEADER:
                field.Orientation = 3  # xlPageField
                set_pivot_page_field_position(pivot, field, int(pivot.PageFields.Count) - 1)
            if policy == "existing-period":
                state = saved[(str(pivot.Name), name)]
                if name == MONTH_COMPLETE_HEADER and not state.is_page_field:
                    # Legacy reports had no month page field. Retain the old
                    # generator's monthly scope when that month exists.
                    selected = tuple(
                        str(item.Name) for item in field.PivotItems()
                        if REPORT_MONTH_NUMBERS.get(str(item.Name).strip().casefold()) == report_date.month
                    ) or None
                else:
                    selected = None if state.show_all else state.selected_items
            elif name == STATUS_DELIVERY_HEADER and MONTH_COMPLETE_HEADER not in names:
                # RFS aging includes every open status, including Waiting AOS.
                selected = tuple(
                    str(item.Name) for item in field.PivotItems()
                    if str(item.Name).casefold() not in {"cancel", "so complete"}
                )
            else:
                selected = None
            set_delivery_pivot_selection(field, selected)
        if names:
            pivot.RefreshTable()


def apply_date_formats(worksheet, df: pd.DataFrame) -> None:
    for header in df.columns:
        column_number = df.columns.get_loc(header) + 1
        column_letter = get_column_letter(column_number)
        date_only_rows: list[int] = []
        date_time_rows: list[int] = []
        for row_number, value in enumerate(df[header], start=2):
            if header in DATE_HEADERS and is_excel_zero_date(value):
                date_time_rows.append(row_number)
                continue
            if not isinstance(value, datetime):
                continue
            if value.time() == time.min:
                date_only_rows.append(row_number)
            else:
                date_time_rows.append(row_number)

        for rows, number_format in (
            (date_only_rows, RFS_DATE_ONLY_FORMAT if header in DASHBOARD_RFS_DATE_HEADERS else DATE_ONLY_FORMAT),
            (date_time_rows, RFS_DATE_TIME_FORMAT if header in DASHBOARD_RFS_DATE_HEADERS else DATE_TIME_FORMAT),
        ):
            for start_row, end_row in contiguous_row_ranges(rows):
                worksheet.Range(
                    f"{column_letter}{start_row}:{column_letter}{end_row}"
                ).NumberFormat = number_format


def apply_na_errors(excel, worksheet, df: pd.DataFrame) -> None:
    """Restore #N/A sentinels as real Excel errors before PivotTable refresh."""
    for header in df.columns:
        error_rows = [
            row_number
            for row_number, value in enumerate(df[header], start=2)
            if str(value).strip().upper() == EXCEL_NA_ERROR
        ]
        if not error_rows:
            continue

        column_number = df.columns.get_loc(header) + 1
        worksheet.Range(
            worksheet.Cells(2, column_number),
            worksheet.Cells(len(df) + 1, column_number),
        ).NumberFormat = "General"

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
                # Keep the native error formula; freezing its value through
                # Copy/PasteSpecial depends on the shared Windows clipboard.
                target.Formula = "=NA()"
                target.Calculate()
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


def find_target_pivot_field(pivot_table, expected_header: str | None = None):
    fields = pivot_table.PivotFields()
    expected = normalize_header(expected_header).casefold() if expected_header else None
    fallback = None
    for field_index in range(1, fields.Count + 1):
        field = fields(field_index)
        name = normalize_header(field.Name)
        if expected is not None and name.casefold() == expected:
            return field
        if fallback is None and name.casefold().startswith("target "):
            fallback = field
    return fallback


def capture_target_pivot_filters(worksheet) -> list[TargetPivotPageFilter]:
    """Capture target page filters before the monthly source header changes."""
    states: list[TargetPivotPageFilter] = []
    pivot_tables = worksheet.PivotTables()
    for pivot_index in range(1, pivot_tables.Count + 1):
        pivot_table = pivot_tables(pivot_index)
        field = find_target_pivot_field(pivot_table)
        if field is None or int(field.Orientation) != 3:  # xlPageField
            continue

        current_page = str(field.CurrentPage)
        show_all = target_item_kind(current_page) == "all"
        selection_kinds: list[str] = []
        try:
            multiple_items = bool(field.EnableMultiplePageItems)
        except Exception:
            multiple_items = False

        if multiple_items and not show_all:
            items = field.PivotItems()
            for item_index in range(1, items.Count + 1):
                item = items(item_index)
                if not bool(item.Visible):
                    continue
                kind = target_item_kind(item.Name)
                if kind is not None and kind != "all" and kind not in selection_kinds:
                    selection_kinds.append(kind)
        elif not show_all:
            kind = target_item_kind(current_page)
            if kind is None:
                raise RuntimeError(
                    f"Unsupported target filter item {current_page!r} in "
                    f"PivotTable '{pivot_table.Name}'."
                )
            selection_kinds.append(kind)

        states.append(
            TargetPivotPageFilter(
                pivot_name=str(pivot_table.Name),
                position=int(field.Position),
                selection_kinds=tuple(selection_kinds),
                show_all=show_all,
            )
        )
    return states


def restore_target_pivot_filters(
    worksheet,
    states: list[TargetPivotPageFilter],
    report_date: date,
) -> None:
    """Rebind monthly target filters after the PivotCache schema refresh."""
    expected_header = target_header(report_date)
    period = target_period(report_date)
    for state in states:
        pivot_table = worksheet.PivotTables(state.pivot_name)
        field = find_target_pivot_field(pivot_table, expected_header)
        if field is None:
            raise RuntimeError(
                f"PivotTable '{state.pivot_name}' is missing target field "
                f"{expected_header!r} after refresh."
            )

        field.Orientation = 3  # xlPageField
        set_pivot_page_field_position(pivot_table, field, state.position)
        field.ClearAllFilters()
        if state.show_all:
            field.CurrentPage = "(All)"
            pivot_table.RefreshTable()
            continue

        desired_labels = {
            target_item_label(kind, period).casefold(): kind
            for kind in state.selection_kinds
        }
        available_items: dict[str, object] = {}
        items = field.PivotItems()
        for item_index in range(1, items.Count + 1):
            item = items(item_index)
            available_items[str(item.Name).strip().casefold()] = item

        missing_labels = sorted(set(desired_labels) - set(available_items))
        if missing_labels:
            raise RuntimeError(
                f"PivotTable '{state.pivot_name}' cannot apply target filter; "
                f"missing item(s): {', '.join(missing_labels)}."
            )

        if len(desired_labels) == 1:
            selected_label = next(iter(desired_labels))
            field.CurrentPage = str(available_items[selected_label].Name)
        else:
            field.EnableMultiplePageItems = True
            for label in desired_labels:
                available_items[label].Visible = True
            for label, item in available_items.items():
                item.Visible = label in desired_labels
        pivot_table.RefreshTable()


def update_pivot_period_titles(worksheet, report_date: date) -> None:
    month_name = calendar.month_name[report_date.month].upper()
    year = report_date.year
    replacements = {
        "target complete": (
            f"TARGET COMPLETE {month_name} {year} "
            f"(Based on Dashboard 1 {month_name} {year})"
        ),
        "order aging start from pre-installation status": (
            f"ORDER AGING START FROM PRE-INSTALLATION STATUS {month_name} {year % 100:02d}"
        ),
        "status order with aging from rfs commit date": (
            f"STATUS ORDER WITH AGING FROM RFS COMMIT DATE {month_name} {year % 100:02d}"
        ),
        "delay completion": (
            "DELAY COMPLETION (SHOULD BE COMPLETE BEFORE "
            f"{month_name} {year})"
        ),
        "target after": f"TARGET AFTER {month_name} {year}",
    }

    used_range = worksheet.UsedRange
    used_values = used_range.Value2
    if used_values is None:
        return
    for row_offset, row in enumerate(used_values):
        for column_offset, value in enumerate(row):
            if not isinstance(value, str):
                continue
            normalized = " ".join(value.strip().casefold().split())
            for prefix, replacement in replacements.items():
                if normalized.startswith(prefix):
                    worksheet.Cells(
                        used_range.Row + row_offset,
                        used_range.Column + column_offset,
                    ).Value = replacement
                    break


def hidden_row_ranges_between_sections(
    section_bottom_row: int,
    next_title_row: int,
    content_rows: Iterable[int] = (),
) -> list[tuple[int, int]]:
    """Return blank row runs to hide while preserving useful section spacing."""
    gap_start = section_bottom_row + 1
    gap_end = next_title_row - 1
    if gap_start > gap_end:
        return []

    protected_rows = {gap_start, gap_end, gap_end - 1}
    for row_number in content_rows:
        if gap_start <= row_number <= gap_end:
            protected_rows.add(row_number)
            protected_rows.add(max(gap_start, row_number - 1))

    hidden_rows = [
        row_number
        for row_number in range(gap_start, gap_end + 1)
        if row_number not in protected_rows
    ]
    if not hidden_rows:
        return []

    ranges: list[tuple[int, int]] = []
    range_start = previous_row = hidden_rows[0]
    for row_number in hidden_rows[1:]:
        if row_number != previous_row + 1:
            ranges.append((range_start, previous_row))
            range_start = row_number
        previous_row = row_number
    ranges.append((range_start, previous_row))
    return ranges


def normalize_pivot_section_row_visibility(worksheet) -> None:
    """Hide vacated Pivot rows using the refreshed section boundaries."""
    used_range = worksheet.UsedRange
    used_values = used_range.Value2
    if used_values is None:
        return
    if not isinstance(used_values, tuple):
        value_rows = ((used_values,),)
    elif used_values and not isinstance(used_values[0], tuple):
        value_rows = (used_values,)
    else:
        value_rows = used_values

    title_markers = tuple(
        normalize_layout_text(marker) for marker in PIVOT_SECTION_TITLE_MARKERS
    )
    title_rows: set[int] = set()
    content_rows: set[int] = set()
    for row_offset, row_values in enumerate(value_rows, start=int(used_range.Row)):
        if not isinstance(row_values, tuple):
            row_values = (row_values,)
        nonblank_values = [value for value in row_values if value not in (None, "")]
        if nonblank_values:
            content_rows.add(row_offset)
        if any(
            marker in normalize_layout_text(value)
            for value in nonblank_values
            for marker in title_markers
        ):
            title_rows.add(row_offset)

    sorted_title_rows = sorted(title_rows)
    if len(sorted_title_rows) < 2:
        return

    pivot_ranges: list[tuple[int, int]] = []
    pivot_row_ranges: list[tuple[int, int]] = []
    pivot_tables = worksheet.PivotTables()
    for pivot_index in range(1, int(pivot_tables.Count) + 1):
        pivot_table = pivot_tables(pivot_index)
        table_range = pivot_table.TableRange2
        pivot_start_row = int(table_range.Row)
        pivot_ranges.append(
            (
                pivot_start_row,
                pivot_start_row + int(table_range.Rows.Count) - 1,
            )
        )
        row_range = pivot_table.RowRange
        if row_range is not None:
            row_start = int(row_range.Row)
            pivot_row_ranges.append(
                (row_start, row_start + int(row_range.Rows.Count) - 1)
            )

    for pivot_start_row, pivot_bottom_row in pivot_row_ranges:
        worksheet.Range(
            worksheet.Cells(pivot_start_row, 1),
            worksheet.Cells(pivot_bottom_row, 1),
        ).EntireRow.Hidden = False
    for title_row in sorted_title_rows:
        worksheet.Rows(title_row).Hidden = False

    for title_row, next_title_row in zip(
        sorted_title_rows,
        sorted_title_rows[1:],
        strict=False,
    ):
        section_bottoms = [
            bottom_row
            for pivot_start_row, bottom_row in pivot_ranges
            if title_row < pivot_start_row < next_title_row
        ]
        if not section_bottoms:
            continue

        section_bottom_row = max(section_bottoms)
        gap_start = section_bottom_row + 1
        gap_end = next_title_row - 1
        if gap_start > gap_end:
            continue

        worksheet.Range(
            worksheet.Cells(gap_start, 1),
            worksheet.Cells(gap_end, 1),
        ).EntireRow.Hidden = False
        gap_content_rows = {
            row_number
            for row_number in content_rows
            if gap_start <= row_number <= gap_end
        }
        for hidden_start, hidden_end in hidden_row_ranges_between_sections(
            section_bottom_row,
            next_title_row,
            gap_content_rows,
        ):
            worksheet.Range(
                worksheet.Cells(hidden_start, 1),
                worksheet.Cells(hidden_end, 1),
            ).EntireRow.Hidden = True


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


def preserve_pivot_refresh_layout(worksheet) -> None:
    # Shared-cache refresh can reformat columns through any connected pivot.
    pivots = worksheet.PivotTables()
    for index in range(1, pivots.Count + 1):
        pivot = pivots(index)
        pivot.PreserveFormatting = True
        pivot.HasAutoFormat = False


def separate_delivery_pivot_sections(worksheet, *, minimum_gap: int = 2) -> int:
    """Keep independent Aging and OTD sections apart without moving paired pivots."""
    used = worksheet.UsedRange
    titles = []
    for row, values in enumerate(used.Value2, int(used.Row)):
        for column, value in enumerate(values, int(used.Column)):
            key = normalize_layout_text(value)
            if any(key.startswith(normalize_layout_text(marker)) for marker in PIVOT_SECTION_TITLE_MARKERS):
                titles.append((row, column, key))
    aging = next((t for t in titles if t[2].startswith('orderagingstartfrompreinstallationstatus')), None)
    delivery = next((t for t in titles if t[2].startswith('ontimedelivery')), None)
    if aging is None or delivery is None or aging[0] != delivery[0]:
        return 0
    next_row = min((row for row, _, _ in titles if row > aging[0]), default=EXCEL_MAX_ROW)
    regions = [pivot.TableRange2 for pivot in worksheet.PivotTables()
               if aging[0] < int(pivot.TableRange2.Row) < next_row]
    left = [region for region in regions if int(region.Column) < delivery[1]]
    right = [region for region in regions if int(region.Column) >= delivery[1]]
    if not left or not right:
        return 0
    left_end = max(int(region.Column + region.Columns.Count - 1) for region in left)
    right_start = min(int(region.Column) for region in right)
    shift = max(0, left_end + minimum_gap + 1 - right_start)
    if shift == 0:
        return 0
    bottom = max(int(region.Row + region.Rows.Count - 1) for region in right)
    end_column = max(int(region.Column + region.Columns.Count - 1) for region in right)
    extension = worksheet.Range(worksheet.Cells(delivery[0], end_column + 1),
                                worksheet.Cells(bottom, end_column + shift))
    if worksheet.Application.WorksheetFunction.CountA(extension):
        raise RuntimeError('Cannot separate Aging and OTD pivots: destination contains other content.')
    block = worksheet.Range(worksheet.Cells(delivery[0], delivery[1]), worksheet.Cells(bottom, end_column))
    block.Cut(Destination=worksheet.Cells(delivery[0], delivery[1] + shift))
    worksheet.Application.CutCopyMode = False
    return shift


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


def normalize_pivot_percentage_borders(
    worksheet,
    blocks: list[PivotPercentageBlock],
) -> None:
    """Keep red outlines on percentage blocks without boxing every data row."""
    xl_edge_left = 7
    xl_edge_top = 8
    xl_edge_bottom = 9
    xl_edge_right = 10
    xl_inside_horizontal = 12
    xl_continuous = 1
    xl_line_style_none = -4142
    xl_medium = -4138
    xl_red = 255

    def apply_red_outline(cell) -> None:
        for border_index in (
            xl_edge_left,
            xl_edge_top,
            xl_edge_bottom,
            xl_edge_right,
        ):
            border = cell.Borders(border_index)
            border.LineStyle = xl_continuous
            border.Color = xl_red
            border.Weight = xl_medium

    for block in blocks:
        pivot_table = worksheet.PivotTables(block.pivot_name)
        table_range = pivot_table.TableRange2
        header_row = int(table_range.Row) + block.header_row_offset
        end_row = int(table_range.Row) + int(table_range.Rows.Count) - 1
        percentage_column = int(table_range.Column) + int(table_range.Columns.Count)

        first_body_row = header_row + 1
        last_body_row = end_row - 1
        if end_row <= header_row:
            continue

        if last_body_row >= first_body_row:
            body_range = worksheet.Range(
                worksheet.Cells(first_body_row, percentage_column),
                worksheet.Cells(last_body_row, percentage_column),
            )
            for border_index in (xl_edge_left, xl_edge_right):
                border = body_range.Borders(border_index)
                border.LineStyle = xl_continuous
                border.Color = xl_red
                border.Weight = xl_medium
            body_range.Borders(xl_inside_horizontal).LineStyle = xl_line_style_none
            body_range.Borders(xl_edge_bottom).LineStyle = xl_line_style_none

            first_body_cell = worksheet.Cells(first_body_row, percentage_column)
            first_body_border = first_body_cell.Borders(xl_edge_top)
            first_body_border.LineStyle = xl_continuous
            first_body_border.Color = xl_red
            first_body_border.Weight = xl_medium

        # Apply section boundaries last because Excel treats adjacent cell
        # borders as one shared edge when ranges are reformatted.
        apply_red_outline(worksheet.Cells(header_row, percentage_column))
        apply_red_outline(worksheet.Cells(end_row, percentage_column))


def completion_thresholds(
    report_date: date,
    mode: str = WEEKLY_MODE,
) -> CompletionThresholds:
    """Return the shared Report 4 completion thresholds."""
    return completion_thresholds_for(
        report_date,
        profile=STANDARD_PROFILE,
        mode=mode,
    )


def apply_weekly_completion_thresholds(
    worksheet,
    blocks: list[PivotPercentageBlock],
    report_date: date,
) -> None:
    """Update percentage legends and traffic-light rules for the report week."""
    used_range = worksheet.UsedRange
    used_values = used_range.Value2
    value_rows = ()
    if used_values is not None:
        if not isinstance(used_values, tuple):
            value_rows = ((used_values,),)
        elif used_values and not isinstance(used_values[0], tuple):
            value_rows = (used_values,)
        else:
            value_rows = used_values

    section_modes = completion_section_modes(value_rows, int(used_range.Row))
    if value_rows:
        for row_offset, row in enumerate(value_rows):
            for column_offset, value in enumerate(row):
                if not isinstance(value, str):
                    continue
                normalized = value.strip().casefold()
                mode = completion_mode_for_row(
                    int(used_range.Row) + row_offset,
                    section_modes,
                )
                thresholds = completion_thresholds(report_date, mode)
                legend_values = {
                    "green": f"Green : >{thresholds.green_percent}%",
                    "yellow": f"Yellow : >={thresholds.yellow_percent}%",
                    "red": f"Red : <{thresholds.red_percent}%",
                }
                for prefix, replacement in legend_values.items():
                    if normalized.startswith(prefix):
                        worksheet.Cells(
                            used_range.Row + row_offset,
                            used_range.Column + column_offset,
                        ).Value = replacement
                        break

    for block in blocks:
        pivot_table = worksheet.PivotTables(block.pivot_name)
        table_range = pivot_table.TableRange2
        header_row = int(table_range.Row) + block.header_row_offset
        mode = completion_mode_for_row(
            header_row,
            section_modes,
        )
        thresholds = completion_thresholds(report_date, mode)
        formula_start_row = header_row + 1
        end_row = int(table_range.Row) + int(table_range.Rows.Count) - 1
        percentage_column = int(table_range.Column) + int(table_range.Columns.Count)
        if end_row < formula_start_row:
            continue

        target_range = worksheet.Range(
            worksheet.Cells(formula_start_row, percentage_column),
            worksheet.Cells(end_row, percentage_column),
        )
        if apply_completion_icon_thresholds(target_range, thresholds) == 0:
            raise RuntimeError(
                f"No traffic-light rule found for percentage block '{block.pivot_name}'."
            )


def identify_target_complete_pivots(worksheet):
    pivot_tables = worksheet.PivotTables()
    pivots = [pivot_tables(index) for index in range(1, pivot_tables.Count + 1)]
    summary_candidates = [
        pivot for pivot in pivots if int(pivot.TableRange2.Column) == 1
    ]
    summary_pivot = min(
        summary_candidates,
        key=lambda pivot: int(pivot.TableRange2.Row),
        default=None,
    )

    status_pivot = None
    if summary_pivot is not None:
        summary_range = summary_pivot.TableRange2
        summary_start = int(summary_range.Row)
        summary_end = summary_start + int(summary_range.Rows.Count) - 1
        status_candidates = []
        for pivot in pivots:
            table_range = pivot.TableRange2
            if int(table_range.Column) <= 1:
                continue
            status_start = int(table_range.Row)
            status_end = status_start + int(table_range.Rows.Count) - 1
            if status_start <= summary_end and summary_start <= status_end:
                status_candidates.append(pivot)
        status_pivot = min(
            status_candidates,
            key=lambda pivot: (
                int(pivot.TableRange2.Row),
                int(pivot.TableRange2.Column),
            ),
            default=None,
        )

    if summary_pivot is None or status_pivot is None:
        raise RuntimeError("Could not identify both TARGET COMPLETE PivotTables.")

    return summary_pivot, status_pivot


def validate_target_complete_pivots(worksheet, dataframe: pd.DataFrame) -> None:
    summary_pivot, status_pivot = identify_target_complete_pivots(worksheet)

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
    report_date: date,
    *,
    previous_report_date: date | None = None,
    month_complete_format: str = "MMMM",
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
        target_filter_states = capture_target_pivot_filters(pivot_sheet)
        delivery_filter_states = capture_delivery_pivot_filters(pivot_sheet)
        alignment_models = capture_pivot_alignments(pivot_sheet)
        if not target_filter_states:
            raise RuntimeError("Previous workbook has no target PivotTable page filters.")
        preserve_pivot_refresh_layout(pivot_sheet)

        row_count = len(df) + 1
        column_count = len(df.columns)
        table = all_order_sheet.ListObjects(TABLE_NAME)
        old_column_count = int(table.Range.Columns.Count)
        old_last_row = max(
            int(all_order_sheet.UsedRange.Rows.Count),
            int(table.Range.Rows.Count),
        )
        clear_last_row = max(old_last_row, row_count)
        all_order_sheet.Range(
            all_order_sheet.Cells(2, 1),
            all_order_sheet.Cells(clear_last_row, max(old_column_count, column_count)),
        ).ClearContents()

        target_range = all_order_sheet.Range(
            all_order_sheet.Cells(1, 1),
            all_order_sheet.Cells(row_count, column_count),
        )
        table.Resize(target_range)
        if old_column_count > column_count:
            all_order_sheet.Range(
                all_order_sheet.Cells(1, column_count + 1),
                all_order_sheet.Cells(clear_last_row, old_column_count),
            ).Clear()
        table.TableStyle = TABLE_STYLE_NAME
        year_fab_upload_column = df.columns.get_loc(YEAR_FAB_UPLOAD_HEADER) + 1
        all_order_sheet.Range(
            all_order_sheet.Cells(2, year_fab_upload_column),
            all_order_sheet.Cells(row_count, year_fab_upload_column),
        ).NumberFormat = YEAR_TEXT_FORMAT
        body_range = all_order_sheet.Range(
            all_order_sheet.Cells(2, 1),
            all_order_sheet.Cells(row_count, column_count),
        )
        body_range.NumberFormat = "General"
        apply_text_formats(all_order_sheet, df)
        body_values = tuple(
            tuple(com_value(value) for value in row)
            for row in df.itertuples(index=False, name=None)
        )
        body_range.Value = body_values
        for column_number, header in enumerate(df.columns, start=1):
            all_order_sheet.Cells(1, column_number).Value = str(header)
        apply_na_errors(excel, all_order_sheet, df)
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

        apply_month_complete_formula(all_order_sheet, df, date_format=month_complete_format)

        autofit_with_limits(all_order_sheet, column_count)
        all_order_sheet.Activate()
        excel.ActiveWindow.FreezePanes = False
        excel.ActiveWindow.SplitColumn = 2
        excel.ActiveWindow.SplitRow = 0
        excel.ActiveWindow.FreezePanes = True
        excel.ActiveWindow.ScrollColumn = 3
        excel.ActiveWindow.ScrollRow = 1

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
                if not configure_pivot_cache_for_current_source(pivot_cache):
                    raise RuntimeError("Could not clear obsolete PivotCache items before refresh.")
                pivot_cache.Refresh()
                pivot_table.RefreshTable()
                refreshed_count += 1
            except Exception as exc:
                raise RuntimeError(
                    f"Could not refresh PivotTable '{pivot_table.Name}': {exc}"
                ) from exc

        restore_target_pivot_filters(
            pivot_sheet,
            target_filter_states,
            report_date,
        )
        configure_delivery_pivot_filters(
            pivot_sheet, delivery_filter_states,
            delivery_pivot_policy(previous_report_date, report_date),
            report_date,
        )
        restore_pivot_percentage_blocks(
            pivot_sheet,
            percentage_scratch,
            percentage_blocks,
        )
        percentage_scratch.Visible = -1  # xlSheetVisible
        percentage_scratch.Delete()
        percentage_scratch = None
        repair_pivot_percentage_ranges(pivot_sheet)
        apply_weekly_completion_thresholds(
            pivot_sheet,
            percentage_blocks,
            report_date,
        )
        normalize_pivot_percentage_borders(
            pivot_sheet,
            percentage_blocks,
        )
        update_pivot_period_titles(pivot_sheet, report_date)
        separate_delivery_pivot_sections(pivot_sheet)
        resize_dynamic_pivot_section_banners(
            pivot_sheet,
            title_markers=PIVOT_SECTION_TITLE_MARKERS,
            side_header_markers=PIVOT_SECTION_SIDE_HEADER_MARKERS,
        )
        normalize_pivot_section_row_visibility(pivot_sheet)
        adjust_pivot_column_widths(pivot_sheet)
        restore_pivot_alignments(pivot_sheet, alignment_models)
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


def has_expected_all_order_freeze_panes(worksheet) -> bool:
    pane = worksheet.sheet_view.pane
    if pane is None or pane.state not in {"frozen", "frozenSplit"}:
        return False

    try:
        split_columns = float(pane.xSplit or 0)
        split_rows = float(pane.ySplit or 0)
    except (TypeError, ValueError):
        return False

    return split_columns == 2.0 and split_rows == 0.0


def validate_output(
    output_path: Path,
    expected_quote_ids: list[str],
    expected_pivot_count: int,
    *,
    expected_headers: list[str],
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
        column_count = len(expected_headers)
        normalized_expected_headers = [normalize_header(header) for header in expected_headers]
        actual_headers = [
            normalize_header(worksheet.cell(1, column).value)
            for column in range(1, column_count + 1)
        ]
        if actual_headers != normalized_expected_headers:
            raise ValueError("Output ALL ORDER column names or order do not match the generated data.")
        actual_rows = worksheet.max_row - 1
        expected_rows = len(expected_quote_ids)
        if actual_rows != expected_rows:
            raise ValueError(f"Output has {actual_rows} data rows, expected {expected_rows}.")
        if not has_expected_all_order_freeze_panes(worksheet):
            pane = worksheet.sheet_view.pane
            split_columns = pane.xSplit if pane is not None else None
            split_rows = pane.ySplit if pane is not None else None
            raise ValueError(
                "Output ALL ORDER must freeze columns A:B "
                f"(xSplit=2, ySplit=0); found xSplit={split_columns}, "
                f"ySplit={split_rows}."
            )
        if TABLE_NAME not in worksheet.tables:
            raise ValueError(f"Output is missing Excel table: {TABLE_NAME}")
        expected_ref = f"A1:{get_column_letter(column_count)}{worksheet.max_row}"
        if worksheet.tables[TABLE_NAME].ref != expected_ref:
            raise ValueError(
                f"{TABLE_NAME} range is {worksheet.tables[TABLE_NAME].ref}, expected {expected_ref}."
            )

        headers = {header: column for column, header in enumerate(actual_headers, start=1)}
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
        for pivot in workbook[PIVOT_SHEET_NAME]._pivots:
            cached_headers = [normalize_header(field.name) for field in pivot.cache.cacheFields]
            if cached_headers[:column_count] != normalized_expected_headers:
                raise ValueError(
                    f"PivotTable '{pivot.name}' source columns do not match the generated ALL ORDER."
                )
        return {
            "rows": actual_rows,
            "columns": column_count,
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
    previous = read_previous_dataframe(files.previous)
    raw = read_raw_dataframe(
        files.raw, previous=previous,
        reference_continuation=delivery_pivot_policy(date_from_filename(files.previous), report_date) == "existing-period",
    )
    build = build_all_order(raw, previous, files.collabs, report_date, files.previous)

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
        refreshed_count = update_workbook_via_com(
            temp_path, build.dataframe, report_date,
            previous_report_date=date_from_filename(files.previous),
            month_complete_format=build.month_complete_format,
        )
        validation = validate_output(
            temp_path,
            build.dataframe[QUOTE_ID_HEADER].map(normalize_quote_id).tolist(),
            refreshed_count,
            expected_headers=build.dataframe.columns.tolist(),
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
        "rfs_text_date_policy": DEFAULT_RFS_TEXT_DATE_POLICY,
        "delivery_pivot_policy": delivery_pivot_policy(date_from_filename(files.previous), report_date),
        "collabs_fallback_rows": build.collabs_fallback_count,
        "zero_fallback_rows": build.zero_fallback_count,
        "duplicate_quote_ids": build.duplicate_quote_ids,
        "duplicate_rows_preserved": build.duplicate_rows_preserved,
        "date_fallback_rows": build.date_fallback_count,
        "field_fallback_rows": build.field_fallback_count,
        "invalid_dates": build.invalid_dates,
        "source_date_corrections": build.source_date_corrections,
        "financial_source_policy": build.financial_source_policy,
        "month_complete_format": build.month_complete_format,
        "reference_workflow": build.reference_workflow,
        "historical_status_source": build.historical_status_source,
        "planned_rfs_year_rollovers": build.planned_rfs_year_rollovers,
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
        default=Path(DEFAULT_COLLABS_DIR),
        help=f"Collabs CSV fallback or directory. Defaults to .\\{DEFAULT_COLLABS_DIR}.",
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
        print(f"[1/5] Reading raw data: {files.raw.name}")
        print(f"[2/5] Loading previous mapping: {files.previous.name}")
        print(f"[3/5] Loading OTC/MRC fallback: {files.collabs.name}")
        print("[4/5] Updating ALL ORDER and refreshing PivotTables in Excel")
        result = generate_ide_tracking(files, output_path, report_date)
        print("[5/5] Validating output")
        print(
            f"Wrote {output_path} | rows={result['rows']} | columns={result['columns']} | "
            f"pivots={result['pivot_tables']}"
        )
        print(
            f"Fallbacks: collabs={result['collabs_fallback_rows']}, "
            f"OTC/MRC zero={result['zero_fallback_rows']}, "
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
