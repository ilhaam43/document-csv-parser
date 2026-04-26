#!/usr/bin/env python3
"""Parse and clean CSV files, then export them to Excel workbooks."""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from charset_normalizer import from_path
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


COMMON_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin1")
NULL_LIKE_VALUES = {"", "na", "n/a", "null", "none", "nan", "-"}
EXCEL_MAX_SHEET_NAME_LENGTH = 31
DEFAULT_INPUT_DIR = "input-today"
DEFAULT_OUTPUT_DIR = "output-today"
DEFAULT_VLOOKUP_DIR = "vlookup-yesterday"
TARGET_HEADER_INSERT_AFTER = "Quo"
TARGET_SOURCE_HEADER = "Target + On Hold Duration"
QUO_EXCLUDE_TERM = "XBOT"
DEPT_SD_HEADER = "Dept SD"
PM_HEADER = "PM"
QUO_HEADER = "Quo"
STATUS_HEADER = "Status"
PROCESS_HEADER = "Process"
PROCESS_ADJUSTMENT_HEADER = "Process Adjustment"
PRODUCT_CATEGORY_HEADER = "Product Category"
GROUP_SALES_HEADER = "group_sales"
DIVISION_SALES_HEADER = "Division_sales"
SEGMENT_SALES_HEADER = "segment_sales"
SALES_HEADER = "Sales"
SERVICE_DELIVERY_DIV_HEADER = "Service Delivery Div"
LOOKUP_SERVICE_DELIVERY_DIV_HEADER = "Service Delivery Div."
LOOKUP_GROUP_SALES_HEADER = "Group Sales"
LOOKUP_DIVISION_SALES_HEADER = "Division Sales"
LOOKUP_SEGMENT_SALES_HEADER = "Segment Sales"
PHASE_HEADER_PREFIX = "Phase"
FAB_UPLOAD_HEADER = "FAB Upload"
YEAR_FAB_UPLOAD_HEADER = "YEAR FAB UPLOAD"
RFS_COMMMIT_HEADER = "RFS Commmit"
AGING_OF_RFS_HEADER = "Aging Of RFS"
STATUS_ORDER_HEADER = "STATUS ORDER"
TARGET_SO_COMPLETE_DATE_HEADERS = ("Target SO Complete Date", "Target SO Completion Date")
EXCEL_DATE_DISPLAY_FORMAT = "yyyy-mm-dd"
EXCEL_TABLE_STYLE_NAME = "TableStyleMedium2"
EXCEL_HEADER_FONT_COLOR = "FFFFFF"
EXCEL_RED_FONT_COLOR = "FF0000"
EXCEL_RED_HEADER_FILL = "FF0000"
EXCEL_YELLOW_HEADER_FILL = "FFFF00"
EXCEL_HEADER_ROW_HEIGHT_POINTS = 22.5
VLOOKUP_SHEET_NAME = "ALL ORDER"


@dataclass(frozen=True)
class ConvertOptions:
    output: Path
    delimiter: str | None
    encoding: str | None
    normalize_headers: bool
    keep_empty: bool
    drop_empty_columns: bool
    dedupe: bool
    infer_types: bool
    combine: bool


def detect_encoding(path: Path) -> str:
    """Detect file encoding with a conservative fallback list."""
    result = from_path(path).best()
    if result and result.encoding:
        return result.encoding

    for encoding in COMMON_ENCODINGS:
        try:
            path.read_text(encoding=encoding)
            return encoding
        except UnicodeDecodeError:
            continue

    return "latin1"


def detect_delimiter(path: Path, encoding: str) -> str:
    sample = path.read_text(encoding=encoding, errors="replace")[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        return dialect.delimiter
    except csv.Error:
        return ","


def normalize_column_name(value: object, index: int) -> str:
    raw = str(value).replace("\ufeff", "").replace("_", " ").strip()
    raw = re.sub(r"^#+", "", raw).strip()
    raw = re.sub(r"\s+", " ", raw)
    if not raw:
        return f"Column {index + 1}"

    words = []
    for word in raw.split(" "):
        if word.lower() == "otc":
            words.append("OTC")
        elif word.isupper() or re.fullmatch(r"[A-Z0-9#()/-]+", word):
            words.append(word)
        else:
            words.append(word[:1].upper() + word[1:])

    return " ".join(words)


def make_unique_columns(columns: Iterable[object], normalize: bool) -> list[str]:
    names: list[str] = []
    seen: dict[str, int] = {}

    for index, column in enumerate(columns):
        name = normalize_column_name(column, index) if normalize else str(column).strip()
        name = name or f"Column {index + 1}"

        count = seen.get(name, 0)
        seen[name] = count + 1
        names.append(name if count == 0 else f"{name} ({count + 1})")

    return names


def rename_output_headers(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(
        columns={
            "Group Sales": GROUP_SALES_HEADER,
            "Division Sales": DIVISION_SALES_HEADER,
            "Segment Sales": SEGMENT_SALES_HEADER,
        }
    )


def escape_excel_formula_text(value: object) -> object:
    if isinstance(value, str) and value.startswith("="):
        return f"'{value}"
    return value


def escape_excel_formulas(df: pd.DataFrame) -> pd.DataFrame:
    escaped = df.copy()
    escaped.columns = [escape_excel_formula_text(column) for column in escaped.columns]

    for column in escaped.columns:
        if pd.api.types.is_object_dtype(escaped[column]) or pd.api.types.is_string_dtype(escaped[column]):
            escaped[column] = escaped[column].map(escape_excel_formula_text, na_action="ignore")

    return escaped


def current_target_header() -> str:
    today = date.today()
    return f"TARGET  Detemined as 1 {today.strftime('%B %Y')}"


def current_target_values(source: pd.Series) -> pd.Series:
    today = date.today()
    month_name = today.strftime("%B")
    target_month = pd.Period(today, freq="M")
    source_months = pd.to_datetime(source, errors="coerce", format="mixed").dt.to_period("M")
    values = pd.Series("Target Not Yet Inputted", index=source.index, dtype="string")

    values[source_months < target_month] = f"Before {month_name}"
    values[source_months == target_month] = f"This {month_name}"
    values[source_months > target_month] = f"After {month_name}"

    return values


def add_target_header(df: pd.DataFrame) -> pd.DataFrame:
    if TARGET_HEADER_INSERT_AFTER not in df.columns or TARGET_SOURCE_HEADER not in df.columns:
        return df

    result = df.copy()
    target_header = current_target_header()
    if target_header in result.columns:
        result[target_header] = current_target_values(result[TARGET_SOURCE_HEADER])
        return result

    insert_at = result.columns.get_loc(TARGET_HEADER_INSERT_AFTER) + 1
    result.insert(insert_at, target_header, current_target_values(result[TARGET_SOURCE_HEADER]))
    return result


def drop_excluded_quo_rows(df: pd.DataFrame) -> pd.DataFrame:
    if TARGET_HEADER_INSERT_AFTER not in df.columns:
        return df

    quo_values = df[TARGET_HEADER_INSERT_AFTER].astype("string")
    return df.loc[~quo_values.str.contains(QUO_EXCLUDE_TERM, case=False, na=False)]


def add_year_fab_upload_column(df: pd.DataFrame) -> pd.DataFrame:
    if FAB_UPLOAD_HEADER not in df.columns:
        return df

    result = df.copy()
    fab_upload_dates = pd.to_datetime(result[FAB_UPLOAD_HEADER], errors="coerce", format="mixed")
    year_values = pd.Series(pd.NA, index=result.index, dtype="string")
    has_date = fab_upload_dates.notna()
    year_values.loc[has_date] = fab_upload_dates.loc[has_date].dt.strftime("%Y")

    if YEAR_FAB_UPLOAD_HEADER in result.columns:
        result[YEAR_FAB_UPLOAD_HEADER] = year_values
        return result

    insert_at = result.columns.get_loc(FAB_UPLOAD_HEADER) + 1
    result.insert(insert_at, YEAR_FAB_UPLOAD_HEADER, year_values)
    return result


def reference_date_from_csv_path(csv_path: Path) -> date:
    match = re.search(r"(\d{8})", csv_path.stem)
    if match:
        return pd.to_datetime(match.group(1), format="%Y%m%d").date()

    return date.today()


def output_filename_from_csv_path(csv_path: Path) -> str:
    reference_date = reference_date_from_csv_path(csv_path)
    return f"Daily Tracking {reference_date.day} {reference_date.strftime('%B %Y')}.xlsx"


def add_aging_of_rfs_column(df: pd.DataFrame, reference_date: date) -> pd.DataFrame:
    if RFS_COMMMIT_HEADER not in df.columns:
        return df

    result = df.copy()
    rfs_commit_dates = pd.to_datetime(result[RFS_COMMMIT_HEADER], errors="coerce", format="mixed")
    aging_values = pd.Series(pd.NA, index=result.index, dtype="Int64")
    has_date = rfs_commit_dates.notna()
    aging_values.loc[has_date] = (pd.Timestamp(reference_date) - rfs_commit_dates.loc[has_date]).dt.days

    if AGING_OF_RFS_HEADER in result.columns:
        result[AGING_OF_RFS_HEADER] = aging_values
        return result

    insert_at = result.columns.get_loc(RFS_COMMMIT_HEADER) + 1
    result.insert(insert_at, AGING_OF_RFS_HEADER, aging_values)
    return result


def add_status_order_column(df: pd.DataFrame) -> pd.DataFrame:
    if AGING_OF_RFS_HEADER not in df.columns:
        return df

    result = df.copy()
    aging_values = pd.to_numeric(result[AGING_OF_RFS_HEADER], errors="coerce")
    status_order_values = pd.Series(pd.NA, index=result.index, dtype="string")
    has_aging = aging_values.notna()
    status_order_values.loc[has_aging] = "Delay"
    status_order_values.loc[aging_values < 0] = "Potential Delay"
    status_order_values.loc[aging_values < -5] = "On Track"

    if STATUS_ORDER_HEADER in result.columns:
        result[STATUS_ORDER_HEADER] = status_order_values
        return result

    insert_at = result.columns.get_loc(AGING_OF_RFS_HEADER) + 1
    result.insert(insert_at, STATUS_ORDER_HEADER, status_order_values)
    return result


def clean_dataframe(df: pd.DataFrame, options: ConvertOptions, csv_path: Path) -> pd.DataFrame:
    df = df.copy()
    df.columns = make_unique_columns(df.columns, options.normalize_headers)
    df = rename_output_headers(df)
    df = add_target_header(df)

    for column in df.columns:
        if pd.api.types.is_object_dtype(df[column]) or pd.api.types.is_string_dtype(df[column]):
            cleaned = (
                df[column]
                .astype("string")
                .str.replace("\ufeff", "", regex=False)
                .str.strip()
                .str.replace(r"\s+", " ", regex=True)
            )
            df[column] = cleaned.mask(cleaned.str.lower().isin(NULL_LIKE_VALUES), pd.NA)

    df = drop_excluded_quo_rows(df)
    df = add_year_fab_upload_column(df)
    df = add_aging_of_rfs_column(df, reference_date_from_csv_path(csv_path))
    df = add_status_order_column(df)

    if not options.keep_empty:
        df = df.dropna(how="all")

    if options.drop_empty_columns:
        df = df.dropna(axis="columns", how="all")

    if options.dedupe:
        df = df.drop_duplicates()

    if options.infer_types:
        df = infer_column_types(df)

    return escape_excel_formulas(df).reset_index(drop=True)


def infer_column_types(df: pd.DataFrame) -> pd.DataFrame:
    converted = df.copy()

    for column in converted.columns:
        series = converted[column]
        non_empty = series.dropna()
        if non_empty.empty:
            continue

        numeric_candidate = (
            non_empty.astype("string")
            .str.replace(",", "", regex=False)
            .str.replace("%", "", regex=False)
        )
        numeric = pd.to_numeric(numeric_candidate, errors="coerce")
        if numeric.notna().mean() >= 0.9:
            converted[column] = pd.to_numeric(
                series.astype("string").str.replace(",", "", regex=False).str.replace("%", "", regex=False),
                errors="coerce",
            )
            continue

        parsed_dates = pd.to_datetime(non_empty, errors="coerce", format="mixed")
        if parsed_dates.notna().mean() >= 0.9:
            converted[column] = pd.to_datetime(series, errors="coerce", format="mixed")

    return converted


def read_csv(path: Path, options: ConvertOptions) -> pd.DataFrame:
    encoding = options.encoding or detect_encoding(path)
    delimiter = options.delimiter or detect_delimiter(path, encoding)

    with path.open("r", encoding=encoding, errors="replace", newline="") as handle:
        rows = list(csv.reader(handle, delimiter=delimiter, skipinitialspace=True))

    first_data_row = next((index for index, row in enumerate(rows) if any(cell.strip() for cell in row)), None)
    if first_data_row is None:
        raise ValueError(f"CSV file is empty: {path}")

    header = rows[first_data_row]
    data_rows = rows[first_data_row + 1 :]
    width = max([len(header), *(len(row) for row in data_rows)] or [len(header)])

    if len(header) < width:
        header = [*header, *(f"Extra Column {index + 1}" for index in range(len(header), width))]

    normalized_rows = [
        [*row, *([""] * (width - len(row)))] if len(row) < width else row[:width]
        for row in data_rows
    ]

    return pd.DataFrame(normalized_rows, columns=header, dtype="string")


def sanitize_sheet_name(path: Path, used: set[str]) -> str:
    base = re.sub(r"[\[\]\:\*\?\/\\]", "_", path.stem).strip() or "Sheet"
    base = base[:EXCEL_MAX_SHEET_NAME_LENGTH]
    sheet_name = base
    suffix = 2

    while sheet_name in used:
        suffix_text = f"_{suffix}"
        sheet_name = f"{base[: EXCEL_MAX_SHEET_NAME_LENGTH - len(suffix_text)]}{suffix_text}"
        suffix += 1

    used.add(sheet_name)
    return sheet_name


def resolve_vlookup_workbook() -> Path:
    lookup_dir = Path(__file__).resolve().parent / DEFAULT_VLOOKUP_DIR
    candidates = sorted(
        path for path in lookup_dir.glob("*.xlsx") if path.is_file() and not path.name.startswith("~$")
    )
    if not candidates:
        raise FileNotFoundError(f"No lookup workbook found in: {lookup_dir}")

    return max(candidates, key=lambda path: path.stat().st_mtime)


def phase_lookup_header(lookup_workbook: Path) -> str:
    match = re.search(r"Daily Tracking\s+(.+)$", lookup_workbook.stem, flags=re.IGNORECASE)
    if match:
        return f"{PHASE_HEADER_PREFIX} {match.group(1).strip()}"

    return PHASE_HEADER_PREFIX


def load_lookup_mappings(
    lookup_workbook: Path,
) -> tuple[dict[str, str], dict[str, str], dict[str, str], dict[str, str], dict[str, str], dict[str, str], dict[str, str], dict[str, str], str]:
    lookup_df = pd.read_excel(lookup_workbook, sheet_name=VLOOKUP_SHEET_NAME, dtype="string")
    lookup_df.columns = make_unique_columns(lookup_df.columns, normalize=True)
    if (
        PM_HEADER not in lookup_df.columns
        or DEPT_SD_HEADER not in lookup_df.columns
        or LOOKUP_SERVICE_DELIVERY_DIV_HEADER not in lookup_df.columns
        or QUO_HEADER not in lookup_df.columns
        or PHASE_HEADER_PREFIX not in lookup_df.columns
        or PROCESS_ADJUSTMENT_HEADER not in lookup_df.columns
        or PRODUCT_CATEGORY_HEADER not in lookup_df.columns
        or SALES_HEADER not in lookup_df.columns
        or LOOKUP_GROUP_SALES_HEADER not in lookup_df.columns
        or LOOKUP_DIVISION_SALES_HEADER not in lookup_df.columns
        or LOOKUP_SEGMENT_SALES_HEADER not in lookup_df.columns
    ):
        raise ValueError(
            "Lookup workbook must contain "
            f"'{PM_HEADER}', '{DEPT_SD_HEADER}', '{LOOKUP_SERVICE_DELIVERY_DIV_HEADER}', "
            f"'{QUO_HEADER}', '{PHASE_HEADER_PREFIX}', '{PROCESS_ADJUSTMENT_HEADER}', and "
            f"'{PRODUCT_CATEGORY_HEADER}', '{SALES_HEADER}', '{LOOKUP_GROUP_SALES_HEADER}', and "
            f"'{LOOKUP_DIVISION_SALES_HEADER}', and '{LOOKUP_SEGMENT_SALES_HEADER}' "
            f"columns in sheet '{VLOOKUP_SHEET_NAME}'."
        )

    dept_sd_mapping: dict[str, str] = {}
    service_delivery_div_mapping: dict[str, str] = {}
    phase_mapping: dict[str, str] = {}
    process_adjustment_mapping: dict[str, str] = {}
    product_category_mapping: dict[str, str] = {}
    group_sales_mapping: dict[str, str] = {}
    division_sales_mapping: dict[str, str] = {}
    segment_sales_mapping: dict[str, str] = {}
    for _, row in lookup_df[
        [
            PM_HEADER,
            DEPT_SD_HEADER,
            LOOKUP_SERVICE_DELIVERY_DIV_HEADER,
            QUO_HEADER,
            PHASE_HEADER_PREFIX,
            PROCESS_ADJUSTMENT_HEADER,
            PRODUCT_CATEGORY_HEADER,
            SALES_HEADER,
            LOOKUP_GROUP_SALES_HEADER,
            LOOKUP_DIVISION_SALES_HEADER,
            LOOKUP_SEGMENT_SALES_HEADER,
        ]
    ].dropna(
        subset=[PM_HEADER]
    ).iterrows():
        pm_value = str(row[PM_HEADER]).strip()
        if not pm_value:
            continue

        dept_sd_value = row[DEPT_SD_HEADER]
        if pm_value not in dept_sd_mapping and not pd.isna(dept_sd_value):
            dept_sd_mapping[pm_value] = str(dept_sd_value).strip()

        service_delivery_div_value = row[LOOKUP_SERVICE_DELIVERY_DIV_HEADER]
        if pm_value not in service_delivery_div_mapping and not pd.isna(service_delivery_div_value):
            service_delivery_div_mapping[pm_value] = str(service_delivery_div_value).strip()

        quo_value = str(row[QUO_HEADER]).strip()
        phase_value = row[PHASE_HEADER_PREFIX]
        if quo_value and quo_value not in phase_mapping and not pd.isna(phase_value):
            phase_mapping[quo_value] = str(phase_value).strip()

        process_adjustment_value = row[PROCESS_ADJUSTMENT_HEADER]
        if quo_value and quo_value not in process_adjustment_mapping and not pd.isna(process_adjustment_value):
            process_adjustment_mapping[quo_value] = str(process_adjustment_value).strip()

        product_category_value = row[PRODUCT_CATEGORY_HEADER]
        if quo_value and quo_value not in product_category_mapping and not pd.isna(product_category_value):
            product_category_mapping[quo_value] = str(product_category_value).strip()

        sales_value = str(row[SALES_HEADER]).strip()
        group_sales_value = row[LOOKUP_GROUP_SALES_HEADER]
        if sales_value and sales_value not in group_sales_mapping and not pd.isna(group_sales_value):
            group_sales_mapping[sales_value] = str(group_sales_value).strip()

        division_sales_value = row[LOOKUP_DIVISION_SALES_HEADER]
        if sales_value and sales_value not in division_sales_mapping and not pd.isna(division_sales_value):
            division_sales_mapping[sales_value] = str(division_sales_value).strip()

        segment_sales_value = row[LOOKUP_SEGMENT_SALES_HEADER]
        if sales_value and sales_value not in segment_sales_mapping and not pd.isna(segment_sales_value):
            segment_sales_mapping[sales_value] = str(segment_sales_value).strip()

    return (
        dept_sd_mapping,
        service_delivery_div_mapping,
        phase_mapping,
        process_adjustment_mapping,
        product_category_mapping,
        group_sales_mapping,
        division_sales_mapping,
        segment_sales_mapping,
        phase_lookup_header(lookup_workbook),
    )


def apply_lookup_values(
    df: pd.DataFrame,
    dept_sd_lookup: dict[str, str],
    service_delivery_div_lookup: dict[str, str],
    phase_lookup: dict[str, str],
    process_adjustment_lookup: dict[str, str],
    product_category_lookup: dict[str, str],
    group_sales_lookup: dict[str, str],
    division_sales_lookup: dict[str, str],
    segment_sales_lookup: dict[str, str],
    phase_header: str,
) -> pd.DataFrame:
    if PM_HEADER not in df.columns and QUO_HEADER not in df.columns:
        return df

    result = df.copy()
    pm_values = result[PM_HEADER].astype("string").str.strip() if PM_HEADER in result.columns else None

    if DEPT_SD_HEADER in result.columns and pm_values is not None:
        looked_up_dept_sd = pm_values.map(dept_sd_lookup)
        result[DEPT_SD_HEADER] = pd.Series(looked_up_dept_sd, index=result.index, dtype="string")

    if pm_values is not None:
        looked_up_service_delivery_div = pd.Series(pm_values.map(service_delivery_div_lookup), index=result.index, dtype="string")
        if SERVICE_DELIVERY_DIV_HEADER in result.columns:
            result[SERVICE_DELIVERY_DIV_HEADER] = looked_up_service_delivery_div
        else:
            insert_at = result.columns.get_loc(DEPT_SD_HEADER) + 1 if DEPT_SD_HEADER in result.columns else len(result.columns)
            result.insert(insert_at, SERVICE_DELIVERY_DIV_HEADER, looked_up_service_delivery_div)

    if SALES_HEADER in result.columns:
        looked_up_group_sales = pd.Series(
            result[SALES_HEADER].astype("string").str.strip().map(group_sales_lookup),
            index=result.index,
            dtype="string",
        )
        if GROUP_SALES_HEADER in result.columns:
            result[GROUP_SALES_HEADER] = looked_up_group_sales

        looked_up_division_sales = pd.Series(
            result[SALES_HEADER].astype("string").str.strip().map(division_sales_lookup),
            index=result.index,
            dtype="string",
        )
        if DIVISION_SALES_HEADER in result.columns:
            result[DIVISION_SALES_HEADER] = looked_up_division_sales
        else:
            insert_at = result.columns.get_loc(GROUP_SALES_HEADER) + 1 if GROUP_SALES_HEADER in result.columns else len(result.columns)
            result.insert(insert_at, DIVISION_SALES_HEADER, looked_up_division_sales)

        looked_up_segment_sales = pd.Series(
            result[SALES_HEADER].astype("string").str.strip().map(segment_sales_lookup),
            index=result.index,
            dtype="string",
        )
        if SEGMENT_SALES_HEADER in result.columns:
            result[SEGMENT_SALES_HEADER] = looked_up_segment_sales

    if QUO_HEADER in result.columns:
        quo_values = result[QUO_HEADER].astype("string").str.strip()
        looked_up_phase = pd.Series(quo_values.map(phase_lookup), index=result.index, dtype="string")
        if phase_header in result.columns:
            result[phase_header] = looked_up_phase
        else:
            insert_at = result.columns.get_loc(STATUS_HEADER) + 1 if STATUS_HEADER in result.columns else len(result.columns)
            result.insert(insert_at, phase_header, looked_up_phase)

        looked_up_process_adjustment = pd.Series(
            quo_values.map(process_adjustment_lookup),
            index=result.index,
            dtype="string",
        )
        if PROCESS_ADJUSTMENT_HEADER in result.columns:
            result[PROCESS_ADJUSTMENT_HEADER] = looked_up_process_adjustment
        else:
            insert_at = result.columns.get_loc(PROCESS_HEADER) + 1 if PROCESS_HEADER in result.columns else len(result.columns)
            result.insert(insert_at, PROCESS_ADJUSTMENT_HEADER, looked_up_process_adjustment)

        if PRODUCT_CATEGORY_HEADER in result.columns:
            result[PRODUCT_CATEGORY_HEADER] = pd.Series(
                quo_values.map(product_category_lookup),
                index=result.index,
                dtype="string",
            )

    return result


def write_excel_sheet(
    writer: pd.ExcelWriter,
    df: pd.DataFrame,
    sheet_name: str,
    dept_sd_lookup: dict[str, str],
    service_delivery_div_lookup: dict[str, str],
    phase_lookup: dict[str, str],
    process_adjustment_lookup: dict[str, str],
    product_category_lookup: dict[str, str],
    group_sales_lookup: dict[str, str],
    division_sales_lookup: dict[str, str],
    segment_sales_lookup: dict[str, str],
    phase_header: str,
) -> None:
    df = apply_lookup_values(
        df,
        dept_sd_lookup,
        service_delivery_div_lookup,
        phase_lookup,
        process_adjustment_lookup,
        product_category_lookup,
        group_sales_lookup,
        division_sales_lookup,
        segment_sales_lookup,
        phase_header,
    )
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    apply_target_so_complete_date_filter_format(worksheet, df)
    apply_worksheet_presentation(worksheet)


def apply_target_so_complete_date_filter_format(worksheet, df: pd.DataFrame) -> None:
    header_positions = {
        str(cell.value).strip(): cell.column
        for cell in worksheet[1]
        if cell.value is not None and str(cell.value).strip()
    }
    target_header = next(
        (header for header in TARGET_SO_COMPLETE_DATE_HEADERS if header in df.columns and header in header_positions),
        None,
    )
    if target_header is None:
        return

    target_dates = pd.to_datetime(df[target_header], errors="coerce", format="mixed")
    target_column = header_positions[target_header]

    for row_index, target_date in enumerate(target_dates, start=2):
        if pd.isna(target_date):
            continue

        cell = worksheet.cell(row=row_index, column=target_column)
        cell.value = target_date.to_pydatetime()
        cell.number_format = EXCEL_DATE_DISPLAY_FORMAT


def apply_worksheet_presentation(worksheet) -> None:
    add_excel_table(worksheet)
    autofit_worksheet_columns(worksheet)
    style_header_row(worksheet)


def add_excel_table(worksheet) -> None:
    table_name = re.sub(r"[^A-Za-z0-9_]", "_", f"Table_{worksheet.title}") or "Table_1"
    if table_name[0].isdigit():
        table_name = f"Table_{table_name}"

    table = Table(displayName=table_name, ref=worksheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name=EXCEL_TABLE_STYLE_NAME,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)


def autofit_worksheet_columns(worksheet) -> None:
    for column_cells in worksheet.iter_cols():
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)

        for cell in column_cells:
            value = cell.value
            if value is None:
                continue

            if isinstance(value, datetime):
                rendered = value.strftime("%Y-%m-%d")
            elif isinstance(value, date):
                rendered = value.strftime("%Y-%m-%d")
            else:
                rendered = str(value)

            max_length = max(max_length, len(rendered))

        worksheet.column_dimensions[column_letter].width = max_length + 2


def style_header_row(worksheet) -> None:
    worksheet.row_dimensions[1].height = EXCEL_HEADER_ROW_HEIGHT_POINTS
    for cell in worksheet[1]:
        fill_color = header_fill_color(cell.value)
        font_color = EXCEL_RED_FONT_COLOR if fill_color == EXCEL_YELLOW_HEADER_FILL else EXCEL_HEADER_FONT_COLOR
        cell.font = Font(color=font_color, bold=True)
        if fill_color is not None:
            cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)


def header_fill_color(header_value: object) -> str | None:
    normalized = normalize_header_key(header_value)
    if not normalized:
        return None

    if (
        normalized in {
            normalize_header_key(DEPT_SD_HEADER),
            normalize_header_key(SERVICE_DELIVERY_DIV_HEADER),
            normalize_header_key(STATUS_ORDER_HEADER),
            normalize_header_key(GROUP_SALES_HEADER),
            normalize_header_key(DIVISION_SALES_HEADER),
            normalize_header_key(SEGMENT_SALES_HEADER),
            normalize_header_key(SALES_HEADER),
            normalize_header_key("Target SO Completion Date"),
            normalize_header_key("Target SO Complete Date"),
        }
        or normalized.startswith("targetdeterminedas")
        or normalized.startswith("targetdeteminedas")
    ):
        return EXCEL_RED_HEADER_FILL

    if (
        normalized in {
            normalize_header_key(PROCESS_ADJUSTMENT_HEADER),
            normalize_header_key(AGING_OF_RFS_HEADER),
            normalize_header_key(YEAR_FAB_UPLOAD_HEADER),
        }
        or normalized.startswith("phase") and normalized != normalize_header_key(PHASE_HEADER_PREFIX)
    ):
        return EXCEL_YELLOW_HEADER_FILL

    return None


def normalize_header_key(header_value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(header_value).lower())


def resolve_csv_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path]

    if not input_path.is_dir():
        raise FileNotFoundError(f"Input path does not exist: {input_path}")

    files = sorted(input_path.glob("*.csv"))
    if not files:
        raise FileNotFoundError(f"No CSV files found in: {input_path}")

    return files


def default_output_path(input_path: Path, combine: bool) -> Path:
    if input_path.is_file():
        return input_path.parent / output_filename_from_csv_path(input_path)

    output_dir = input_path.parent / DEFAULT_OUTPUT_DIR

    if combine:
        return output_dir / "combined.xlsx"

    return output_dir


def resolve_output_path(input_path: Path, csv_files: list[Path], args: argparse.Namespace) -> Path:
    output = args.output or default_output_path(input_path, args.combine)

    if input_path.is_file() and (output.exists() and output.is_dir()):
        return output / output_filename_from_csv_path(input_path)

    if input_path.is_dir() and args.combine and output.suffix.lower() != ".xlsx":
        return output / "combined.xlsx"

    if input_path.is_dir() and not args.combine and output.suffix.lower() == ".xlsx":
        raise ValueError("Use --combine when writing multiple CSV files to one .xlsx output file.")

    if len(csv_files) == 1 and input_path.is_dir() and args.combine and output.suffix.lower() != ".xlsx":
        return output / "combined.xlsx"

    return output


def convert_one(csv_path: Path, output_path: Path, options: ConvertOptions) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df = clean_dataframe(read_csv(csv_path, options), options, csv_path)
    (
        dept_sd_lookup,
        service_delivery_div_lookup,
        phase_lookup,
        process_adjustment_lookup,
        product_category_lookup,
        group_sales_lookup,
        division_sales_lookup,
        segment_sales_lookup,
        phase_header,
    ) = load_lookup_mappings(resolve_vlookup_workbook())
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        write_excel_sheet(
            writer,
            df,
            VLOOKUP_SHEET_NAME,
            dept_sd_lookup,
            service_delivery_div_lookup,
            phase_lookup,
            process_adjustment_lookup,
            product_category_lookup,
            group_sales_lookup,
            division_sales_lookup,
            segment_sales_lookup,
            phase_header,
        )
    print(f"Wrote {output_path}")


def convert_many(csv_files: list[Path], options: ConvertOptions) -> None:
    (
        dept_sd_lookup,
        service_delivery_div_lookup,
        phase_lookup,
        process_adjustment_lookup,
        product_category_lookup,
        group_sales_lookup,
        division_sales_lookup,
        segment_sales_lookup,
        phase_header,
    ) = load_lookup_mappings(resolve_vlookup_workbook())
    if options.combine:
        options.output.parent.mkdir(parents=True, exist_ok=True)
        used_sheet_names: set[str] = set()
        with pd.ExcelWriter(options.output, engine="openpyxl") as writer:
            for csv_path in csv_files:
                df = clean_dataframe(read_csv(csv_path, options), options, csv_path)
                sheet_name = sanitize_sheet_name(csv_path, used_sheet_names)
                write_excel_sheet(
                    writer,
                    df,
                    sheet_name,
                    dept_sd_lookup,
                    service_delivery_div_lookup,
                    phase_lookup,
                    process_adjustment_lookup,
                    product_category_lookup,
                    group_sales_lookup,
                    division_sales_lookup,
                    segment_sales_lookup,
                    phase_header,
                )
        print(f"Wrote {options.output}")
        return

    options.output.mkdir(parents=True, exist_ok=True)
    for csv_path in csv_files:
        convert_one(csv_path, options.output / output_filename_from_csv_path(csv_path), options)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Clean CSV data and export it to Excel (.xlsx).",
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        default=Path(DEFAULT_INPUT_DIR),
        help=f"CSV file or directory containing CSV files. Defaults to .\\{DEFAULT_INPUT_DIR}.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help=f"Output .xlsx file or output directory. Defaults to .\\{DEFAULT_OUTPUT_DIR} for directory input.",
    )
    parser.add_argument("--delimiter", help="CSV delimiter. Auto-detected when omitted.")
    parser.add_argument("--encoding", help="CSV encoding. Auto-detected when omitted.")
    parser.add_argument(
        "--keep-empty",
        action="store_true",
        help="Keep fully empty rows. Empty columns are kept by default.",
    )
    parser.add_argument(
        "--drop-empty-columns",
        action="store_true",
        help="Drop columns where all values are empty.",
    )
    parser.add_argument(
        "--no-normalize-headers",
        action="store_true",
        help="Keep original column names instead of converting underscores to readable headers.",
    )
    parser.add_argument(
        "--dedupe",
        action="store_true",
        help="Remove duplicate rows after cleaning.",
    )
    parser.add_argument(
        "--infer-types",
        action="store_true",
        help="Infer numeric and date columns. By default, values are preserved as text.",
    )
    parser.add_argument(
        "--combine",
        action="store_true",
        help="When input is a directory, write all CSV files to one workbook with one sheet per file.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    input_path = args.input.resolve()

    try:
        csv_files = resolve_csv_files(input_path)
        output_path = resolve_output_path(input_path, csv_files, args).resolve()
        options = ConvertOptions(
            output=output_path,
            delimiter=args.delimiter,
            encoding=args.encoding,
            normalize_headers=not args.no_normalize_headers,
            keep_empty=args.keep_empty,
            drop_empty_columns=args.drop_empty_columns,
            dedupe=args.dedupe,
            infer_types=args.infer_types,
            combine=args.combine,
        )

        if len(csv_files) == 1 and input_path.is_file():
            convert_one(csv_files[0], output_path, options)
        else:
            convert_many(csv_files, options)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
