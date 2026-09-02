#!/usr/bin/env python3
"""Generate an iPhone tracking workbook from a daily tracking workbook."""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import time
from datetime import date
from pathlib import Path

import pandas as pd

from excel_pivot_filters import (
    apply_month_target_filter,
    configure_pivot_cache_for_current_source,
    find_named_pivot_field,
    set_visible_items_exact,
)
from excel_pivot_layout import (
    align_completion_threshold_legends,
    resize_dynamic_pivot_section_banners,
    update_dynamic_pivot_section_titles,
)


ALL_ORDER_SHEET_NAME = "ALL ORDER"
IPHONE_SHEET_NAME = "ALL ORDER IPHONE"
PIVOT_SHEET_NAME = "PIVOT"
ON_PROGRESS_PREFIX = "ALL ORDER ON PROGRESS"
ALL_ORDER_TABLE_NAME = "Table57"
IPHONE_TABLE_NAME = "Table27"
PRODUCT_CATEGORY_HEADER = "Product Category"
QUO_HEADER = "quo"
AGING_OF_RFS_HEADER = "Aging Of RFS"
YEAR_FAB_UPLOAD_HEADER = "YEAR FAB UPLOAD"
YEAR_FAB_UPLOAD_EXCLUDED_VALUES = {"", "1900", "nan", "none", "n/a", "(blank)", "null"}
DEFAULT_INPUT_DIR = "input-iphone"
DEFAULT_REFERENCE_DIR = "vlookup-iphone"
DEFAULT_OUTPUT_DIR = "output-iphone"
PREVIOUS_IPHONE_MAPPING_HEADERS = (
    PRODUCT_CATEGORY_HEADER,
    "PM",
    "Dept SD",
    "Service Delivery Div. ",
)
EXCEL_TABLE_STYLE_NAME = "TableStyleMedium2"
# Keep each COM SAFEARRAY comfortably below Excel's allocation limit on the
# Windows server, where workbooks can already consume substantial memory.
EXCEL_COM_ROW_BATCH_SIZE = 250
PERCENTAGE_COMPLETION_HEADER = "Percentage of Completion (SO Complete, Cancel & Change Target)"
PHASE_PIVOT_ORDER = (
    "00-New",
    "01-Presales",
    "02-Survey",
    "03-Allocation",
    "04-Pre Installation",
    "05-Customer Preparation",
    "06-Installation",
    "07-UAT On Hold",
    "Cancel",
    "SO Complete",
)
IPHONE_PIVOT_TARGET_FILTERS = {
    "PivotTable1": "target",
    "PivotTable2": "target",
    "PivotTable3": "target",
    "PivotTable4": "target",
    "PivotTable8": "before",
    "PivotTable9": "before",
    "PivotTable12": "not_yet",
    "PivotTable5": "all",
    "PivotTable6": "all",
}
IPHONE_COMPLETE_PHASES = (
    "04-Pre Installation",
    "05-Customer Preparation",
    "07-UAT On Hold",
    "Cancel",
    "SO Complete",
)
IPHONE_ACTIVE_PHASES = (
    "04-Pre Installation",
    "05-Customer Preparation",
    "07-UAT On Hold",
)


def profile_log(label: str, start: float) -> None:
    if os.getenv("IPHONE_PROFILE", "").strip().lower() in {"1", "true", "yes"}:
        print(f"[profile] {label}: {time.perf_counter() - start:.2f}s", file=sys.stderr)


def reference_date_from_workbook_path(workbook_path: Path) -> date:
    match = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", workbook_path.stem)
    if not match:
        raise ValueError(f"Could not infer report date from workbook name: {workbook_path.name}")

    return pd.to_datetime(" ".join(match.groups()), format="%d %B %Y").date()


def default_output_path(input_workbook: Path) -> Path:
    report_date = reference_date_from_workbook_path(input_workbook)
    return input_workbook.with_name(
        f"Daily Tracking Iphone {report_date.day} {report_date.strftime('%B %Y')}.xlsx"
    )


def newest_workbook_in_dir(directory: Path) -> Path:
    if not directory.is_dir():
        raise FileNotFoundError(f"Directory not found: {directory}")

    candidates = sorted(
        path for path in directory.glob("*.xlsx") if path.is_file() and not path.name.startswith("~$")
    )
    if not candidates:
        raise FileNotFoundError(f"No .xlsx workbook found in: {directory}")

    return max(candidates, key=lambda path: path.stat().st_mtime)


def resolve_workbook_path(path: Path) -> Path:
    return newest_workbook_in_dir(path) if path.is_dir() else path


def output_path_for_directory(output_dir: Path, input_workbook: Path) -> Path:
    return output_dir / default_output_path(input_workbook).name


def progress_sheet_name(report_date: date) -> str:
    return f"{ON_PROGRESS_PREFIX} {report_date.day} {report_date.strftime('%b')}"


def column_index_by_header(sheet, header_name: str) -> int:
    target = re.sub(r"[^a-z0-9]+", "", header_name.lower())
    used_columns = max(sheet.UsedRange.Columns.Count, 300)
    for column_index in range(1, used_columns + 1):
        value = sheet.Cells(1, column_index).Value
        normalized = re.sub(r"[^a-z0-9]+", "", str(value).lower())
        if normalized == target:
            return column_index

    raise ValueError(f"Could not find header '{header_name}' on sheet '{sheet.Name}'")


def sheet_by_name(workbook, sheet_name: str):
    for sheet in workbook.Worksheets:
        if sheet.Name == sheet_name:
            return sheet
    return None


def first_progress_sheet(workbook):
    for sheet in workbook.Worksheets:
        if str(sheet.Name).startswith(ON_PROGRESS_PREFIX):
            return sheet
    return None


def reset_sheet(sheet) -> None:
    try:
        while sheet.ListObjects.Count:
            sheet.ListObjects(1).Delete()
    except Exception:
        pass
    sheet.Cells.Clear()


def copy_used_range(source_sheet, target_sheet, xl) -> tuple[int, int]:
    reset_sheet(target_sheet)
    used_range = source_sheet.UsedRange
    row_count = used_range.Rows.Count
    column_count = used_range.Columns.Count
    used_range.Copy()
    target_sheet.Range("A1").PasteSpecial(Paste=-4104)  # xlPasteAll
    target_sheet.Range("A1").PasteSpecial(Paste=8)  # xlPasteColumnWidths
    xl.CutCopyMode = False
    return row_count, column_count


def ensure_table(sheet, table_name: str, row_count: int, column_count: int) -> None:
    if row_count < 1 or column_count < 1:
        raise ValueError(f"Cannot create table '{table_name}' on empty sheet '{sheet.Name}'")

    data_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(row_count, column_count))
    if sheet.ListObjects.Count:
        table = sheet.ListObjects(1)
        table.Resize(data_range)
    else:
        table = sheet.ListObjects.Add(1, data_range, None, 1)

    table.Name = table_name
    table.TableStyle = EXCEL_TABLE_STYLE_NAME


def iphone_header_caption(header_value: object) -> object:
    text = str(header_value or "").strip()
    if normalize_header_key(text) == normalize_header_key(QUO_HEADER):
        return QUO_HEADER
    if normalize_header_key(text) == normalize_header_key("Service Delivery Div"):
        return "Service Delivery Div. "

    target_match = re.match(
        r"^(TARGET\s+Detemined as 1\s+[A-Za-z]+)\s+(\d{4})$",
        text,
        flags=re.IGNORECASE,
    )
    if target_match:
        return f"{target_match.group(1)} {target_match.group(2)[-2:]}"

    return header_value


def is_dated_phase_header(header_value: object) -> bool:
    return bool(re.match(r"^phase\s+\d{1,2}\s+[A-Za-z]+\s+\d{4}$", str(header_value).strip(), flags=re.IGNORECASE))


def capture_sheet_column_model(sheet, column_count: int) -> dict[str, object]:
    template_sheet = sheet.Parent.Worksheets.Add()
    template_sheet.Visible = 0
    try:
        source_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, column_count))
        target_range = template_sheet.Range(template_sheet.Cells(1, 1), template_sheet.Cells(1, column_count))
        source_range.Copy()
        target_range.PasteSpecial(Paste=-4122)  # xlPasteFormats
        widths = {
            column_index: sheet.Columns(column_index).ColumnWidth
            for column_index in range(1, column_count + 1)
        }
        captions = {
            column_index: sheet.Cells(1, column_index).Value
            for column_index in range(1, column_count + 1)
        }
        return {
            "sheet": template_sheet,
            "widths": widths,
            "captions": captions,
            "row_height": sheet.Rows(1).RowHeight,
        }
    except Exception:
        try:
            template_sheet.Delete()
        except Exception:
            pass
        raise


def apply_sheet_column_model(sheet, model: dict[str, object], column_count: int, xl) -> None:
    template_sheet = model.get("sheet")
    if template_sheet is not None:
        try:
            template_sheet.Range(
                template_sheet.Cells(1, 1),
                template_sheet.Cells(1, column_count),
            ).Copy()
            sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, column_count)).PasteSpecial(Paste=-4122)
            xl.CutCopyMode = False
        except Exception:
            pass

    widths = model.get("widths", {})
    if isinstance(widths, dict):
        for column_index in range(1, column_count + 1):
            width = widths.get(column_index)
            if width is None:
                continue
            try:
                sheet.Columns(column_index).ColumnWidth = width
            except Exception:
                pass

    captions = model.get("captions", {})
    if isinstance(captions, dict):
        for column_index in range(1, column_count + 1):
            current_caption = sheet.Cells(1, column_index).Value
            template_caption = captions.get(column_index)
            if template_caption is not None and normalize_header_key(template_caption) == normalize_header_key(
                current_caption
            ):
                sheet.Cells(1, column_index).Value = template_caption
            else:
                sheet.Cells(1, column_index).Value = iphone_header_caption(current_caption)

    if isinstance(widths, dict) and isinstance(captions, dict):
        for column_index in range(1, column_count + 1):
            template_caption = captions.get(column_index)
            current_caption = sheet.Cells(1, column_index).Value
            template_width = widths.get(column_index)
            if template_width is None or not is_dated_phase_header(template_caption) or not is_dated_phase_header(
                current_caption
            ):
                continue

            try:
                width_delta = (len(str(current_caption).strip()) - len(str(template_caption).strip())) * 0.423
                sheet.Columns(column_index).ColumnWidth = max(8.09, float(template_width) + width_delta)
            except Exception:
                pass

    row_height = model.get("row_height")
    if row_height is not None:
        try:
            sheet.Rows(1).RowHeight = row_height
        except Exception:
            pass


def delete_sheet_model(model: dict[str, object]) -> None:
    template_sheet = model.get("sheet")
    if template_sheet is None:
        return
    try:
        template_sheet.Delete()
    except Exception:
        pass


def capture_column_widths(sheet, column_count: int) -> dict[int, float]:
    widths: dict[int, float] = {}
    for column_index in range(1, column_count + 1):
        try:
            widths[column_index] = float(sheet.Columns(column_index).ColumnWidth)
        except Exception:
            pass
    return widths


def capture_header_captions(sheet, column_count: int) -> dict[int, object]:
    return {column_index: sheet.Cells(1, column_index).Value for column_index in range(1, column_count + 1)}


def apply_column_widths(sheet, widths: dict[int, float]) -> None:
    for column_index, width in widths.items():
        try:
            sheet.Columns(column_index).ColumnWidth = width
        except Exception:
            pass


def apply_dated_phase_width_adjustments(
    sheet,
    template_widths: dict[int, float],
    template_captions: dict[int, object],
) -> None:
    for column_index, template_caption in template_captions.items():
        current_caption = sheet.Cells(1, column_index).Value
        template_width = template_widths.get(column_index)
        if template_width is None or not is_dated_phase_header(template_caption) or not is_dated_phase_header(
            current_caption
        ):
            continue

        try:
            width_delta = (len(str(current_caption).strip()) - len(str(template_caption).strip())) * 0.423
            sheet.Columns(column_index).ColumnWidth = max(8.09, float(template_width) + width_delta)
        except Exception:
            pass


def build_iphone_sheet(all_order_sheet, iphone_sheet, xl) -> tuple[int, int]:
    source_range = all_order_sheet.UsedRange
    source_row_count = source_range.Rows.Count
    column_count = source_range.Columns.Count
    header_values = all_order_sheet.Range(
        all_order_sheet.Cells(1, 1),
        all_order_sheet.Cells(1, column_count),
    ).Value2
    if not isinstance(header_values, tuple) or not header_values:
        raise ValueError(f"Sheet '{all_order_sheet.Name}' is empty")

    headers = tuple(header_values[0] if isinstance(header_values[0], tuple) else header_values)
    product_category_index = next(
        (
            index
            for index, header in enumerate(headers)
            if normalize_header_key(header) == normalize_header_key(PRODUCT_CATEGORY_HEADER)
        ),
        None,
    )
    if product_category_index is None:
        raise ValueError(f"Could not find header '{PRODUCT_CATEGORY_HEADER}' on sheet '{all_order_sheet.Name}'")

    iphone_rows = [headers]
    for start_row in range(2, source_row_count + 1, EXCEL_COM_ROW_BATCH_SIZE):
        end_row = min(source_row_count, start_row + EXCEL_COM_ROW_BATCH_SIZE - 1)
        batch_values = all_order_sheet.Range(
            all_order_sheet.Cells(start_row, 1),
            all_order_sheet.Cells(end_row, column_count),
        ).Value2
        if not isinstance(batch_values, tuple):
            continue
        if batch_values and not isinstance(batch_values[0], tuple):
            batch_values = (batch_values,)
        for row in batch_values:
            product_category = str(row[product_category_index] or "").strip()
            if re.search(r"iphone", product_category, flags=re.IGNORECASE):
                iphone_rows.append(tuple(row))

    row_count = len(iphone_rows)
    column_count = len(headers)
    column_model = capture_sheet_column_model(iphone_sheet, column_count)
    reset_sheet(iphone_sheet)
    for start_index in range(0, row_count, EXCEL_COM_ROW_BATCH_SIZE):
        row_batch = tuple(iphone_rows[start_index : start_index + EXCEL_COM_ROW_BATCH_SIZE])
        first_target_row = start_index + 1
        last_target_row = first_target_row + len(row_batch) - 1
        iphone_sheet.Range(
            iphone_sheet.Cells(first_target_row, 1),
            iphone_sheet.Cells(last_target_row, column_count),
        ).Value2 = row_batch

    apply_sheet_column_model(iphone_sheet, column_model, column_count, xl)

    aging_column = next(
        (
            index
            for index, header in enumerate(headers, start=1)
            if normalize_header_key(header) == normalize_header_key(AGING_OF_RFS_HEADER)
        ),
        None,
    )
    ensure_table(iphone_sheet, IPHONE_TABLE_NAME, row_count, column_count)
    apply_sheet_column_model(iphone_sheet, column_model, column_count, xl)
    if aging_column is not None and row_count > 1:
        iphone_sheet.Range(
            iphone_sheet.Cells(2, aging_column),
            iphone_sheet.Cells(row_count, aging_column),
        ).NumberFormat = all_order_sheet.Cells(2, aging_column).NumberFormat
    delete_sheet_model(column_model)
    return row_count, column_count


def normalize_header_key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def previous_iphone_mappings(reference_workbook: Path) -> dict[str, dict[str, str]]:
    df = pd.read_excel(
        reference_workbook,
        sheet_name=IPHONE_SHEET_NAME,
        dtype="string",
        engine="openpyxl",
    )
    if df.empty:
        return {}

    column_by_key = {normalize_header_key(column): column for column in df.columns}
    quo_column = column_by_key.get(normalize_header_key(QUO_HEADER))
    if quo_column is None:
        raise ValueError(f"Reference sheet '{IPHONE_SHEET_NAME}' is missing quo column")

    source_columns = {
        header: column_by_key[normalize_header_key(header)]
        for header in PREVIOUS_IPHONE_MAPPING_HEADERS
        if normalize_header_key(header) in column_by_key
    }

    mappings: dict[str, dict[str, str]] = {}
    for _, row in df[[quo_column, *source_columns.values()]].iterrows():
        quo_value = str(row[quo_column]).strip()
        if not quo_value:
            continue

        mapped_values = {
            target_header: str(row[source_header]).strip()
            for target_header, source_header in source_columns.items()
            if not pd.isna(row[source_header]) and str(row[source_header]).strip()
        }
        if mapped_values:
            mappings[quo_value] = mapped_values

    return mappings


def apply_previous_iphone_mappings(all_order_sheet, mappings: dict[str, dict[str, str]]) -> int:
    if not mappings:
        return 0

    quo_column = column_index_by_header(all_order_sheet, QUO_HEADER)
    target_columns = {
        header: column_index_by_header(all_order_sheet, header)
        for header in PREVIOUS_IPHONE_MAPPING_HEADERS
    }
    row_count = all_order_sheet.UsedRange.Rows.Count
    updated_count = 0

    if row_count < 2:
        return 0

    quo_range = all_order_sheet.Range(all_order_sheet.Cells(2, quo_column), all_order_sheet.Cells(row_count, quo_column))
    quo_values = quo_range.Value
    if not isinstance(quo_values, tuple):
        quo_rows = [(quo_values,)]
    elif quo_values and not isinstance(quo_values[0], tuple):
        quo_rows = [(value,) for value in quo_values]
    else:
        quo_rows = list(quo_values)

    target_column_values: dict[str, list[list[object]]] = {}
    for header, column_index in target_columns.items():
        value_range = all_order_sheet.Range(
            all_order_sheet.Cells(2, column_index),
            all_order_sheet.Cells(row_count, column_index),
        )
        values = value_range.Value
        if not isinstance(values, tuple):
            rows = [[values]]
        elif values and not isinstance(values[0], tuple):
            rows = [[value] for value in values]
        else:
            rows = [list(row) for row in values]
        target_column_values[header] = rows

    touched_headers: set[str] = set()
    for row_offset, row_value in enumerate(quo_rows):
        quo_value = str(row_value[0] or "").strip()
        row_mapping = mappings.get(quo_value)
        if not row_mapping:
            continue

        for header, mapped_value in row_mapping.items():
            current_value = target_column_values[header][row_offset][0]
            if str(current_value or "").strip() != mapped_value:
                target_column_values[header][row_offset][0] = mapped_value
                touched_headers.add(header)
                updated_count += 1

    for header in touched_headers:
        column_index = target_columns[header]
        value_range = all_order_sheet.Range(
            all_order_sheet.Cells(2, column_index),
            all_order_sheet.Cells(row_count, column_index),
        )
        value_range.Value = tuple(tuple(row) for row in target_column_values[header])

    return updated_count


def rename_progress_sheet(workbook, report_date: date) -> str:
    target_name = progress_sheet_name(report_date)
    progress_sheet = first_progress_sheet(workbook)
    if progress_sheet is None:
        raise ValueError(f"Could not find '{ON_PROGRESS_PREFIX} ...' sheet in reference workbook")

    if progress_sheet.Name != target_name:
        existing = sheet_by_name(workbook, target_name)
        if existing is not None:
            existing.Delete()
        progress_sheet.Name = target_name

    return target_name


def refresh_pivots(workbook) -> int:
    pivot_tables_to_refresh = []
    for sheet in workbook.Worksheets:
        pivot_tables = sheet.PivotTables()
        for pivot_index in range(1, pivot_tables.Count + 1):
            pivot_tables_to_refresh.append(pivot_tables(pivot_index))

    refreshed_cache_indexes: set[int] = set()
    for pivot_table in pivot_tables_to_refresh:
        try:
            cache_index = int(pivot_table.CacheIndex)
            if cache_index in refreshed_cache_indexes:
                continue
            pivot_cache = pivot_table.PivotCache()
            configure_pivot_cache_for_current_source(pivot_cache)
            pivot_cache.Refresh()
            refreshed_cache_indexes.add(cache_index)
        except Exception:
            try:
                pivot_table.RefreshTable()
            except Exception:
                pass

    for pivot_table in pivot_tables_to_refresh:
        try:
            pivot_table.RefreshTable()
        except Exception:
            pass

    for pivot_table in pivot_tables_to_refresh:
        apply_year_fab_upload_filter(pivot_table)

    for pivot_table in pivot_tables_to_refresh:
        try:
            pivot_table.Update()
        except Exception:
            pass

    return len(pivot_tables_to_refresh)


def apply_iphone_monthly_pivot_specs(workbook, report_date: date) -> None:
    """Restore the target/process/phase filters that a cache refresh can discard."""
    pivot_sheet = sheet_by_name(workbook, PIVOT_SHEET_NAME)
    if pivot_sheet is not None:
        try:
            pivot_tables = pivot_sheet.PivotTables()
        except Exception:
            pivot_tables = None
        if pivot_tables is not None:
            for pivot_name, filter_mode in IPHONE_PIVOT_TARGET_FILTERS.items():
                try:
                    pivot_table = pivot_tables(pivot_name)
                except Exception:
                    continue
                try:
                    pivot_table.ManualUpdate = False
                except Exception:
                    pass
                apply_month_target_filter(
                    pivot_table,
                    report_date,
                    filter_mode,
                    orientation=3,
                )
                phase_field = find_named_pivot_field(pivot_table, {"Phase"})
                allowed_phases = (
                    IPHONE_ACTIVE_PHASES
                    if pivot_name in {"PivotTable8", "PivotTable9", "PivotTable12"}
                    else IPHONE_COMPLETE_PHASES
                )
                set_visible_items_exact(phase_field, allowed_phases)
                try:
                    pivot_table.RefreshTable()
                except Exception:
                    pass

    progress_sheet = first_progress_sheet(workbook)
    if progress_sheet is None:
        return
    try:
        progress_pivots = progress_sheet.PivotTables()
    except Exception:
        return

    progress_specs = {
        "PivotTable9": ("target", IPHONE_COMPLETE_PHASES),
        "PivotTable2": ("before_not_yet", IPHONE_ACTIVE_PHASES),
    }
    for pivot_name, (target_mode, allowed_phases) in progress_specs.items():
        try:
            pivot_table = progress_pivots(pivot_name)
        except Exception:
            continue
        try:
            pivot_table.ManualUpdate = False
        except Exception:
            pass
        apply_month_target_filter(
            pivot_table,
            report_date,
            target_mode,
            orientation=1,
            position=1,
        )
        process_field = find_named_pivot_field(pivot_table, {"Process Adjustment"})
        set_visible_items_exact(process_field, {"New Reg", "Non New Reg"})
        phase_field = find_named_pivot_field(pivot_table, {"Phase"})
        set_visible_items_exact(phase_field, allowed_phases)
        try:
            pivot_table.RefreshTable()
        except Exception:
            pass

    try:
        left_pivot = progress_pivots("PivotTable9")
        if int(left_pivot.TableRange2.Row) != 2:
            left_pivot.TableRange2.Cut(progress_sheet.Range("A2"))
    except Exception:
        pass


def pivot_group_max_height(pivot_sheet, pivot_names: tuple[str, ...]) -> int:
    heights = []
    for pivot_name in pivot_names:
        try:
            heights.append(int(pivot_sheet.PivotTables(pivot_name).TableRange2.Rows.Count))
        except Exception:
            pass
    return max(heights, default=0)


def find_pivot_title_row(pivot_sheet, marker: str) -> int | None:
    marker_key = normalize_header_key(marker)
    try:
        used_range = pivot_sheet.UsedRange
        values = used_range.Value
    except Exception:
        return None
    if not values:
        return None
    if not isinstance(values, tuple):
        value_rows = ((values,),)
    elif values and not isinstance(values[0], tuple):
        value_rows = (values,)
    else:
        value_rows = values
    for row_offset, row_values in enumerate(value_rows, start=used_range.Row):
        if not isinstance(row_values, tuple):
            row_values = (row_values,)
        for value in row_values:
            if marker_key in normalize_header_key(value):
                return row_offset
    return None


def shift_iphone_lower_sections_for_month_rollover(pivot_sheet, template_top_height: int) -> tuple[int, int | None]:
    """Move lower sections when the refreshed top block grows at month rollover."""
    current_top_height = pivot_group_max_height(
        pivot_sheet,
        ("PivotTable1", "PivotTable2", "PivotTable3", "PivotTable4"),
    )
    growth = current_top_height - template_top_height
    if growth <= 0:
        return 0, None

    title_row = find_pivot_title_row(pivot_sheet, "delay completion")
    if title_row is None:
        return 0, None

    row_shift = growth + 1
    try:
        pivot_sheet.Rows(f"{title_row}:{title_row + row_shift - 1}").Insert()
    except Exception:
        return 0, None
    return row_shift, title_row


def shifted_row_mapping(mapping: dict[int, object], start_row: int | None, row_shift: int) -> dict[int, object]:
    if start_row is None or row_shift == 0:
        return mapping
    return {
        (row + row_shift if row >= start_row else row): value
        for row, value in mapping.items()
    }


def is_excluded_year_fab_upload_value(value: object) -> bool:
    normalized = str(value).strip()
    if normalized.lower() in YEAR_FAB_UPLOAD_EXCLUDED_VALUES:
        return True

    if re.fullmatch(r"\d{4}(?:\.0+)?", normalized):
        return int(float(normalized)) < 2023

    return False


def apply_year_fab_upload_filter(pivot_table) -> None:
    try:
        field = pivot_table.PivotFields(YEAR_FAB_UPLOAD_HEADER)
        field.EnableMultiplePageItems = True
        try:
            field.ClearAllFilters()
        except Exception:
            pass

        for item in field.PivotItems():
            if not is_excluded_year_fab_upload_value(item.Value):
                try:
                    item.Visible = True
                except Exception:
                    pass

        for item in field.PivotItems():
            if is_excluded_year_fab_upload_value(item.Value):
                try:
                    item.Visible = False
                except Exception:
                    pass
    except Exception:
        pass


def column_letter(column_number: int) -> str:
    result = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def header_column(sheet, header_row: int, header_text: str, start_col: int = 5, end_col: int = 20) -> int | None:
    normalized = normalize_header_key(header_text)
    for column_index in range(start_col, end_col + 1):
        if normalize_header_key(sheet.Cells(header_row, column_index).Value) == normalized:
            return column_index
    return None


def contiguous_left_pivot_last_row(pivot_sheet, header_row: int) -> int:
    for row_index in range(header_row + 1, header_row + 120):
        has_label = pivot_sheet.Cells(row_index, 1).Value not in (None, "")
        has_count = pivot_sheet.Cells(row_index, 3).Value not in (None, "")
        if not has_label and not has_count:
            return row_index - 1
    return header_row


def pivot_phase_block_needs_repair(pivot_sheet, header_row: int) -> bool:
    if header_column(pivot_sheet, header_row, "SO Complete") is not None:
        return False

    for row_index in range(max(1, header_row - 2), header_row + 2):
        for column_index in range(5, 13):
            text = str(pivot_sheet.Cells(row_index, column_index).Text or "").strip().upper()
            if text == "#SPILL!":
                return True

    return pivot_sheet.Cells(header_row, 1).Value not in (None, "") and pivot_sheet.Cells(header_row, 5).Value in (None, "")


def sheet_header_map(sheet) -> dict[str, int]:
    return {
        normalize_header_key(sheet.Cells(1, column_index).Value): column_index
        for column_index in range(1, sheet.UsedRange.Columns.Count + 1)
        if sheet.Cells(1, column_index).Value not in (None, "")
    }


def column_values(sheet, column_index: int, row_count: int) -> list[object]:
    values = sheet.Range(sheet.Cells(2, column_index), sheet.Cells(row_count, column_index)).Value
    if row_count <= 1:
        return []
    if not isinstance(values, tuple):
        return [values]
    return [row[0] if isinstance(row, tuple) else row for row in values]


def capture_phase_pivot_formats(workbook, xl):
    pivot_sheet = sheet_by_name(workbook, PIVOT_SHEET_NAME)
    if pivot_sheet is None:
        return None, {}

    format_sheet = None
    format_sources = {}
    for header_row in (9, 56, 167):
        if format_sheet is None:
            format_sheet = workbook.Worksheets.Add()
            format_sheet.Name = "__iphone_phase_fmt"
            format_sheet.Visible = 0

        destination_row = (len(format_sources) * 81) + 1
        source_range = pivot_sheet.Range(
            pivot_sheet.Cells(max(1, header_row - 1), 5),
            pivot_sheet.Cells(header_row + 79, 20),
        )
        target_range = format_sheet.Range(
            format_sheet.Cells(destination_row, 1),
            format_sheet.Cells(destination_row + 80, 16),
        )
        source_range.Copy()
        target_range.PasteSpecial(Paste=-4122)  # xlPasteFormats
        xl.CutCopyMode = False
        format_sources[header_row] = target_range

    return format_sheet, format_sources


def repair_spilled_phase_pivot_blocks(workbook, xl=None, format_sources=None, header_rows=None) -> None:
    """Rebuild only collapsed/#SPILL phase pivot blocks; leave normal pivots untouched."""
    pivot_sheet = sheet_by_name(workbook, PIVOT_SHEET_NAME)
    iphone_sheet = sheet_by_name(workbook, IPHONE_SHEET_NAME)
    if pivot_sheet is None or iphone_sheet is None:
        return

    format_sources = format_sources or {}
    header_map = sheet_header_map(iphone_sheet)
    required_headers = {
        "process": normalize_header_key("Process"),
        "service": normalize_header_key("Service Delivery Div. "),
        "dept": normalize_header_key("Dept SD"),
        "phase": normalize_header_key("Phase"),
        "year": normalize_header_key(YEAR_FAB_UPLOAD_HEADER),
    }
    if any(header_key not in header_map for header_key in required_headers.values()):
        return

    row_count = iphone_sheet.UsedRange.Rows.Count
    source_columns = {
        name: column_values(iphone_sheet, header_map[header_key], row_count)
        for name, header_key in required_headers.items()
    }
    source_rows = []
    for index in range(max(0, row_count - 1)):
        year_value = source_columns["year"][index] if index < len(source_columns["year"]) else None
        if is_excluded_year_fab_upload_value(year_value):
            continue

        source_rows.append(
            {
                "process": str(source_columns["process"][index] or "").strip(),
                "service": str(source_columns["service"][index] or "").strip(),
                "dept": str(source_columns["dept"][index] or "").strip(),
                "phase": str(source_columns["phase"][index] or "").strip(),
            }
        )

    process_names = {row["process"] for row in source_rows if row["process"]}
    service_names = {row["service"] for row in source_rows if row["service"]}
    dept_names = {row["dept"] for row in source_rows if row["dept"]}
    active_phases = {
        row["phase"]
        for row in source_rows
        if row["phase"] and any(row["phase"] == phase for phase in PHASE_PIVOT_ORDER)
    }

    phase_columns = [phase for phase in PHASE_PIVOT_ORDER if phase in active_phases]
    if "SO Complete" not in phase_columns:
        phase_columns.append("SO Complete")
    if not phase_columns:
        return

    def counts_for(current_process: str | None, current_service: str | None, current_dept: str | None) -> dict[str, int]:
        counts = {phase: 0 for phase in phase_columns}
        for row in source_rows:
            if current_process is not None and row["process"] != current_process:
                continue
            if current_service is not None and row["service"] != current_service:
                continue
            if current_dept is not None and row["dept"] != current_dept:
                continue
            if row["phase"] in counts:
                counts[row["phase"]] += 1
        return counts

    for header_row in header_rows or (9, 56, 167):
        if not pivot_phase_block_needs_repair(pivot_sheet, header_row):
            continue

        last_row = contiguous_left_pivot_last_row(pivot_sheet, header_row)
        if last_row <= header_row:
            continue

        pivot_sheet.Range(
            pivot_sheet.Cells(max(1, header_row - 1), 5),
            pivot_sheet.Cells(last_row, 20),
        ).ClearContents()

        format_source = format_sources.get(header_row)
        if format_source is not None:
            try:
                format_source.Copy()
                pivot_sheet.Range(
                    pivot_sheet.Cells(max(1, header_row - 1), 5),
                    pivot_sheet.Cells(header_row + 79, 20),
                ).PasteSpecial(Paste=-4122)  # xlPasteFormats
                if xl is not None:
                    xl.CutCopyMode = False
            except Exception:
                pass

        pivot_sheet.Cells(header_row - 1, 5).Value = "Count of quo"
        pivot_sheet.Cells(header_row - 1, 6).Value = "Column Labels"
        pivot_sheet.Cells(header_row, 5).Value = "Div./Dept."
        for offset, phase in enumerate(phase_columns, start=6):
            pivot_sheet.Cells(header_row, offset).Value = phase
        last_phase_col = 5 + len(phase_columns)

        current_process = None
        current_service = None
        for row_index in range(header_row + 1, last_row + 1):
            label = str(pivot_sheet.Cells(row_index, 1).Value or "").strip()
            if not label:
                continue

            pivot_sheet.Cells(row_index, 5).Value = label
            if normalize_header_key(label) == normalize_header_key("Grand Total"):
                row_counts = counts_for(None, None, None)
            elif label in process_names:
                current_process = label
                current_service = None
                row_counts = counts_for(current_process, None, None)
            elif label in service_names:
                current_service = label
                row_counts = counts_for(current_process, current_service, None)
            elif label in dept_names:
                row_counts = counts_for(current_process, current_service, label)
            else:
                row_counts = {phase: 0 for phase in phase_columns}

            for offset, phase in enumerate(phase_columns, start=6):
                count = row_counts.get(phase, 0)
                pivot_sheet.Cells(row_index, offset).Value = count if count else None

        try:
            phase_format_source = None
            if format_source is not None:
                phase_format_source = format_source.Columns(2)
            else:
                phase_format_source = pivot_sheet.Range(
                    pivot_sheet.Cells(max(1, header_row - 1), 6),
                    pivot_sheet.Cells(header_row + 79, 6),
                )

            for column_index in range(6, last_phase_col + 1):
                phase_format_source.Copy()
                pivot_sheet.Range(
                    pivot_sheet.Cells(max(1, header_row - 1), column_index),
                    pivot_sheet.Cells(header_row + 79, column_index),
                ).PasteSpecial(Paste=-4122)  # xlPasteFormats
            if xl is not None:
                xl.CutCopyMode = False
            else:
                pivot_sheet.Parent.Application.CutCopyMode = False
        except Exception:
            pass

        try:
            pivot_sheet.Range(
                pivot_sheet.Cells(header_row + 1, 6),
                pivot_sheet.Cells(last_row, last_phase_col),
            ).NumberFormat = "0"
        except Exception:
            pass


def formula_column_offsets(formula_r1c1: str) -> list[int]:
    offsets: list[int] = []
    for match in re.finditer(r"RC(?:\[(-?\d+)\]|(\d+))?", formula_r1c1 or ""):
        relative_offset, absolute_column = match.groups()
        if absolute_column is not None:
            continue

        offsets.append(int(relative_offset) if relative_offset is not None else 0)

    return offsets


def capture_percentage_completion_formats(workbook, xl):
    pivot_sheet = sheet_by_name(workbook, PIVOT_SHEET_NAME)
    if pivot_sheet is None:
        return None, {}, {}, {}

    format_sheet = None
    format_sources = {}
    column_widths = {}
    formula_offsets = {}
    grand_total_format_offsets = {}
    for header_row in (9, 56, 167):
        percentage_col = header_column(pivot_sheet, header_row, PERCENTAGE_COMPLETION_HEADER)
        if percentage_col is None:
            continue

        if format_sheet is None:
            format_sheet = workbook.Worksheets.Add()
            format_sheet.Name = "__iphone_pct_fmt"
            format_sheet.Visible = 0

        destination_column = len(format_sources) + 1
        source_range = pivot_sheet.Range(
            pivot_sheet.Cells(header_row, percentage_col),
            pivot_sheet.Cells(header_row + 79, percentage_col),
        )
        target_range = format_sheet.Range(
            format_sheet.Cells(1, destination_column),
            format_sheet.Cells(80, destination_column),
        )
        source_range.Copy()
        target_range.PasteSpecial(Paste=-4122)  # xlPasteFormats
        xl.CutCopyMode = False
        format_sources[header_row] = target_range
        column_widths[header_row] = pivot_sheet.Columns(percentage_col).ColumnWidth
        formula_offsets[header_row] = [
            offset
            for offset in formula_column_offsets(str(pivot_sheet.Cells(header_row + 1, percentage_col).FormulaR1C1))
            if percentage_col + offset != 3
        ]
        grand_total_format_offsets[header_row] = max(
            2,
            contiguous_left_pivot_last_row(pivot_sheet, header_row) - header_row + 1,
        )

    return format_sheet, format_sources, column_widths, formula_offsets, grand_total_format_offsets


def clear_stale_percentage_completion_columns(workbook) -> None:
    """Free old side-calculation columns so refreshed PivotTables can expand normally."""
    pivot_sheet = sheet_by_name(workbook, PIVOT_SHEET_NAME)
    if pivot_sheet is None:
        return

    for header_row in (9, 56, 167):
        percentage_col = header_column(pivot_sheet, header_row, PERCENTAGE_COMPLETION_HEADER, end_col=25)
        if percentage_col is None:
            continue

        try:
            pivot_sheet.Range(
                pivot_sheet.Cells(header_row, percentage_col),
                pivot_sheet.Cells(header_row + 79, percentage_col),
            ).Clear()
        except Exception:
            pass


def repair_percentage_completion_columns(
    workbook,
    xl,
    format_sources=None,
    column_widths=None,
    formula_offsets=None,
    grand_total_format_offsets=None,
    header_rows=None,
) -> None:
    pivot_sheet = sheet_by_name(workbook, PIVOT_SHEET_NAME)
    if pivot_sheet is None:
        return

    format_sources = format_sources or {}
    column_widths = column_widths or {}
    formula_offsets = formula_offsets or {}
    grand_total_format_offsets = grand_total_format_offsets or {}

    for header_row in header_rows or (9, 56, 167):
        pivot_table = next(
            (
                pivot_sheet.PivotTables()(index)
                for index in range(1, pivot_sheet.PivotTables().Count + 1)
                if int(pivot_sheet.PivotTables()(index).TableRange1.Row) <= header_row
                <= int(
                    pivot_sheet.PivotTables()(index).TableRange1.Row
                    + pivot_sheet.PivotTables()(index).TableRange1.Rows.Count
                    - 1
                )
                and int(pivot_sheet.PivotTables()(index).TableRange1.Column) >= 5
            ),
            None,
        )
        if pivot_table is None:
            continue

        table_start_col = int(pivot_table.TableRange1.Column)
        table_end_col = table_start_col + int(pivot_table.TableRange1.Columns.Count) - 1
        so_col = header_column(pivot_sheet, header_row, "SO Complete", table_start_col, table_end_col)
        cancel_col = header_column(pivot_sheet, header_row, "Cancel", table_start_col, table_end_col)
        formula_col = table_end_col + 1
        data_row_start = header_row + 1
        data_row_end = int(pivot_table.TableRange1.Row + pivot_table.TableRange1.Rows.Count - 1)

        for column_index in range(5, 20):
            if normalize_header_key(pivot_sheet.Cells(header_row, column_index).Value) == normalize_header_key(
                PERCENTAGE_COMPLETION_HEADER
            ):
                pivot_sheet.Range(
                    pivot_sheet.Cells(header_row, column_index),
                    pivot_sheet.Cells(data_row_end, column_index),
                ).Clear()

        try:
            pivot_sheet.Range(
                pivot_sheet.Cells(header_row, formula_col),
                pivot_sheet.Cells(header_row + 79, formula_col),
            ).Clear()
        except Exception:
            pass

        pivot_sheet.Cells(header_row, formula_col).Value = PERCENTAGE_COMPLETION_HEADER
        formula_letter = column_letter(formula_col)
        numerator_columns = [column for column in (cancel_col, so_col) if column is not None]

        for row_index in range(data_row_start, data_row_end + 1):
            if numerator_columns:
                numerator = "+".join(
                    f"{column_letter(column)}{row_index}" for column in numerator_columns
                )
                formula = f"=({numerator})/C{row_index}"
            else:
                formula = "=0"
            pivot_sheet.Range(f"{formula_letter}{row_index}").Formula = formula
            pivot_sheet.Range(f"{formula_letter}{row_index}").NumberFormat = "0.00%"

        format_source = format_sources.get(header_row)
        if format_source is not None:
            try:
                format_source.Cells(1, 1).Copy()
                pivot_sheet.Cells(header_row, formula_col).PasteSpecial(Paste=-4122)  # xlPasteFormats

                body_last_row = data_row_end - 1
                if body_last_row >= data_row_start:
                    format_source.Cells(2, 1).Copy()
                    body_range = pivot_sheet.Range(
                        pivot_sheet.Cells(data_row_start, formula_col),
                        pivot_sheet.Cells(body_last_row, formula_col),
                    )
                    body_range.PasteSpecial(Paste=-4122)  # xlPasteFormats
                    body_range.Borders(12).LineStyle = -4142  # xlInsideHorizontal / xlLineStyleNone

                grand_total_offset = min(
                    max(grand_total_format_offsets.get(header_row, 2), 2),
                    int(format_source.Rows.Count),
                )
                format_source.Cells(grand_total_offset, 1).Copy()
                pivot_sheet.Cells(data_row_end, formula_col).PasteSpecial(Paste=-4122)  # xlPasteFormats
                xl.CutCopyMode = False
            except Exception:
                pass
        if header_row in column_widths:
            try:
                pivot_sheet.Columns(formula_col).ColumnWidth = column_widths[header_row]
            except Exception:
                pass


def generate_iphone_tracking(input_workbook: Path, reference_workbook: Path, output_workbook: Path) -> None:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("pywin32 is required because the iPhone workbook refreshes real Excel pivots.") from exc

    if not input_workbook.is_file():
        raise FileNotFoundError(f"Input workbook not found: {input_workbook}")
    if not reference_workbook.is_file():
        raise FileNotFoundError(f"Reference workbook not found: {reference_workbook}")

    report_date = reference_date_from_workbook_path(input_workbook)
    output_workbook.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(reference_workbook, output_workbook)

    xl = None
    source_wb = None
    target_wb = None
    try:
        total_start = time.perf_counter()
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            xl.ScreenUpdating = False
        except Exception:
            pass
        try:
            xl.EnableEvents = False
        except Exception:
            pass

        phase_start = time.perf_counter()
        source_wb = xl.Workbooks.Open(str(input_workbook.resolve()))
        profile_log("open source workbook", phase_start)
        phase_start = time.perf_counter()
        target_wb = xl.Workbooks.Open(str(output_workbook.resolve()))
        profile_log("open target workbook", phase_start)

        source_all_order = sheet_by_name(source_wb, ALL_ORDER_SHEET_NAME)
        target_all_order = sheet_by_name(target_wb, ALL_ORDER_SHEET_NAME)
        target_iphone = sheet_by_name(target_wb, IPHONE_SHEET_NAME)
        target_pivot_sheet = sheet_by_name(target_wb, PIVOT_SHEET_NAME)
        if source_all_order is None or target_all_order is None or target_iphone is None or target_pivot_sheet is None:
            raise ValueError("Input/reference workbook must contain ALL ORDER and ALL ORDER IPHONE sheets.")
        template_top_height = pivot_group_max_height(
            target_pivot_sheet,
            ("PivotTable1", "PivotTable2", "PivotTable3", "PivotTable4"),
        )
        reference_report_date = reference_date_from_workbook_path(reference_workbook)
        is_month_rollover = (
            reference_report_date.year,
            reference_report_date.month,
        ) != (report_date.year, report_date.month)

        phase_start = time.perf_counter()
        (
            pct_format_sheet,
            pct_format_sources,
            pct_column_widths,
            pct_formula_offsets,
            pct_grand_total_format_offsets,
        ) = (
            capture_percentage_completion_formats(target_wb, xl)
        )
        profile_log("capture percentage formats", phase_start)
        phase_start = time.perf_counter()
        phase_format_sheet, phase_format_sources = capture_phase_pivot_formats(target_wb, xl)
        profile_log("capture phase pivot formats", phase_start)
        phase_start = time.perf_counter()
        source_column_count = source_all_order.UsedRange.Columns.Count
        target_all_order_widths = capture_column_widths(target_all_order, source_column_count)
        target_iphone_widths = capture_column_widths(target_iphone, source_column_count)
        target_all_order_captions = capture_header_captions(target_all_order, source_column_count)
        target_iphone_captions = capture_header_captions(target_iphone, source_column_count)
        all_order_column_model = capture_sheet_column_model(target_all_order, source_column_count)
        profile_log("capture sheet models", phase_start)
        phase_start = time.perf_counter()
        row_count, column_count = copy_used_range(source_all_order, target_all_order, xl)
        profile_log("copy all order", phase_start)
        phase_start = time.perf_counter()
        apply_sheet_column_model(target_all_order, all_order_column_model, column_count, xl)
        delete_sheet_model(all_order_column_model)
        ensure_table(target_all_order, ALL_ORDER_TABLE_NAME, row_count, column_count)
        profile_log("restore all order model", phase_start)
        phase_start = time.perf_counter()
        mappings = previous_iphone_mappings(reference_workbook)
        profile_log("load previous mappings", phase_start)
        phase_start = time.perf_counter()
        updated_count = apply_previous_iphone_mappings(target_all_order, mappings)
        profile_log("apply previous mappings", phase_start)
        phase_start = time.perf_counter()
        iphone_rows, _ = build_iphone_sheet(target_all_order, target_iphone, xl)
        profile_log("build iphone sheet", phase_start)
        phase_start = time.perf_counter()
        progress_name = rename_progress_sheet(target_wb, report_date)
        clear_stale_percentage_completion_columns(target_wb)
        refreshed_count = refresh_pivots(target_wb)
        apply_iphone_monthly_pivot_specs(target_wb, report_date)
        row_shift = 0
        shift_start_row = None
        if is_month_rollover:
            row_shift, shift_start_row = shift_iphone_lower_sections_for_month_rollover(
                target_pivot_sheet,
                template_top_height,
            )
        if row_shift:
            pct_format_sources = shifted_row_mapping(pct_format_sources, shift_start_row, row_shift)
            pct_column_widths = shifted_row_mapping(pct_column_widths, shift_start_row, row_shift)
            pct_formula_offsets = shifted_row_mapping(pct_formula_offsets, shift_start_row, row_shift)
            pct_grand_total_format_offsets = shifted_row_mapping(
                pct_grand_total_format_offsets,
                shift_start_row,
                row_shift,
            )
            phase_format_sources = shifted_row_mapping(phase_format_sources, shift_start_row, row_shift)
        percentage_header_rows = tuple(sorted(pct_format_sources)) or (9, 56, 167 + row_shift)
        phase_header_rows = tuple(sorted(phase_format_sources)) or percentage_header_rows
        month_upper = report_date.strftime("%B").upper()
        update_dynamic_pivot_section_titles(
            target_wb.Worksheets(PIVOT_SHEET_NAME),
            {
                "target complete": (
                    f" TARGET COMPLETE {month_upper} {report_date.year} "
                    f"(Based on Dashboard {month_upper} {report_date.year})"
                ),
                "delay completion": (
                    f"DELAY COMPLETION (SHOULD BE COMPLETE BEFORE "
                    f"{month_upper} {report_date.year})"
                ),
                "all target": f"ALL TARGET {month_upper}",
            },
        )
        profile_log("refresh pivots", phase_start)
        phase_start = time.perf_counter()
        repair_spilled_phase_pivot_blocks(
            target_wb,
            xl,
            phase_format_sources,
            phase_header_rows,
        )
        profile_log("repair spilled phase pivots", phase_start)
        phase_start = time.perf_counter()
        repair_percentage_completion_columns(
            target_wb,
            xl,
            pct_format_sources,
            pct_column_widths,
            pct_formula_offsets,
            pct_grand_total_format_offsets,
            percentage_header_rows,
        )
        align_completion_threshold_legends(
            target_pivot_sheet,
            report_date,
            (
                (13, 2, "weekly"),
                (11, 3, "weekly"),
                (22, 1, "weekly"),
            ),
        )
        profile_log("repair percentage columns", phase_start)
        phase_start = time.perf_counter()
        resize_dynamic_pivot_section_banners(
            target_wb.Worksheets(PIVOT_SHEET_NAME),
            title_markers=(
                "target complete",
                "delay completion",
                "all target",
                "target not inputted",
            ),
            side_header_markers=(PERCENTAGE_COMPLETION_HEADER,),
        )
        profile_log("resize pivot section banners", phase_start)
        phase_start = time.perf_counter()
        apply_column_widths(target_all_order, target_all_order_widths)
        apply_column_widths(target_iphone, target_iphone_widths)
        apply_dated_phase_width_adjustments(target_all_order, target_all_order_widths, target_all_order_captions)
        apply_dated_phase_width_adjustments(target_iphone, target_iphone_widths, target_iphone_captions)
        profile_log("restore column widths", phase_start)

        if pct_format_sheet is not None:
            try:
                pct_format_sheet.Delete()
            except Exception:
                pass
        if phase_format_sheet is not None:
            try:
                phase_format_sheet.Delete()
            except Exception:
                pass

        phase_start = time.perf_counter()
        target_wb.Save()
        profile_log("save workbook", phase_start)
        profile_log("total iphone generation", total_start)
        print(f"[info] Applied {updated_count} previous iPhone mappings")
        print(f"[info] Wrote {iphone_rows - 1} iPhone rows to '{IPHONE_SHEET_NAME}'")
        print(f"[info] Refreshed {refreshed_count} PivotTables, progress sheet '{progress_name}'")
    finally:
        if source_wb is not None:
            try:
                source_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if target_wb is not None:
            try:
                target_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate Daily Tracking iPhone workbook from Daily Tracking and previous iPhone workbook.",
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        default=Path(DEFAULT_INPUT_DIR),
        help=f"Daily Tracking workbook or directory. Defaults to .\\{DEFAULT_INPUT_DIR}.",
    )
    parser.add_argument(
        "-r",
        "--reference",
        type=Path,
        default=Path(DEFAULT_REFERENCE_DIR),
        help=f"Previous iPhone workbook or directory. Defaults to .\\{DEFAULT_REFERENCE_DIR}.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path(DEFAULT_OUTPUT_DIR),
        help=f"Output workbook or directory. Defaults to .\\{DEFAULT_OUTPUT_DIR}.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    input_path = args.input.resolve()
    reference_path = args.reference.resolve()
    input_workbook = resolve_workbook_path(input_path).resolve()
    reference_workbook = resolve_workbook_path(reference_path).resolve()

    output_arg = args.output.resolve()
    if output_arg.suffix.lower() == ".xlsx":
        output_workbook = output_arg
    else:
        output_workbook = output_path_for_directory(output_arg, input_workbook).resolve()

    try:
        generate_iphone_tracking(input_workbook, reference_workbook, output_workbook)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote {output_workbook}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
