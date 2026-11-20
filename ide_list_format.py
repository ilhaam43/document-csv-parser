"""IDE List export rules and native Excel writer, isolated from legacy IDE reports."""

from __future__ import annotations

import calendar
import logging
import math
import os
import re
import shutil
import tempfile
from collections import Counter
from dataclasses import dataclass, replace
from datetime import date, datetime
from pathlib import Path
from time import perf_counter
from types import SimpleNamespace
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import BUILTIN_FORMATS

from ide_pivot_presentation import pivot_row_identity
from excel_pivot_filters import is_dynamic_target_header
from ide_workbook_package import conditional_formats, validate_table_and_caches, worksheet_part


LOGGER = logging.getLogger(__name__)


STATUS = {
    "quotation": "1.Quotation", "pra-delivery": "2.Pra-Delivery",
    "delivery": "3.Delivery", "waiting aos": "4.Waiting AOS",
    "so complete": "5.SO Complete", "cancel": "6.Cancel",
}
REQUIRED = {"Key", "Quote ID", "Quote Name", "Status", "Target RFS", "Actual RFS Date",
            "FAB Upload Date", "SO Completion Date", "Revenue", "Department", "PM", "APM"}
PREFIX = (
    "Key", "Quote ID", "Quote Name", "SO Number", "Customer", "Department", "PM", "APM",
    "MS Contributor", "Progress", "Progress status", "Actual RFS Date", "Target RFS",
    "{target}", "Aging Of RFS", "STATUS ORDER", "ON TIME DELIVERY", "Product", "Updated",
    "Resolved", "Admin", "Components", "Labels", "Parent", "Created", "CID", "Nomor PO",
    "PIC Customer", "Customer Address", "Sales", "Segment Sales", "Project Type", "Brand",
    "Partner", "Catalog Status", "FAB Sign Date", "FAB Upload Date", "YEAR FAB UPLOAD",
    "Cash Date", "SO Date", "Service Ready Date", "ORDER TYPE", "PROCESS", "AOS Sent Date",
    "AOS Sign Date", "SO Completion Date", "Status", "Yesterday Status ", "MONTH COMPLETE",
    "Revenue", "Revenue Upside", "Penalty BC", "CoS Actual", "Risk Actual", "CoS BC", "Risk BC",
    "Penalty Actual", "Status HOTO D2", "Cost Planned", "Actual Cost", "CEI", "Month Complete2",
    "Revenue (Juta)", "Nodin Number", "SimCard QTY", "OTD", "Nodin Date",
)
DATE_FIELDS = {"Target RFS", "Actual RFS Date", "Updated", "Resolved", "Created", "FAB Sign Date",
               "FAB Upload Date", "Cash Date", "SO Date", "Service Ready Date", "AOS Sent Date",
               "AOS Sign Date", "SO Completion Date", "Nodin Date"}
FINANCIAL_FIELDS = {
    "Revenue", "Revenue Upside", "Penalty BC", "CoS Actual", "Risk Actual", "CoS BC",
    "Risk BC", "Penalty Actual", "Cost Planned", "Actual Cost", "Revenue (Juta)",
}
PROGRESS_BAR_COLOR = 0xCAD37F  # Excel OLE color for the manual's #7FD3CA.
RAW_ORDER_TYPE = "Order Type (Raw)"


@dataclass(frozen=True)
class ColumnPresentation:
    number_format: str = "General"
    alignment: tuple = (None, "top", False, False, 0, 0)
    header_alignment: tuple = ("center", "center", True, False, 0, 0)
    font_name: str | None = None
    font_size: float | None = None


def presentation_key(name: str, *, raw: bool = False) -> str:
    name = header(name).casefold()
    if raw and name == "order type":
        return RAW_ORDER_TYPE.casefold()
    if is_dynamic_target_header(name):
        return "__monthly_target__"
    if re.fullmatch(r"month complete\d+", name) or (raw and name == "month complete"):
        return "__source_month_complete__"
    return name


def read_column_presentations(path: Path, sheet_name: str, *, records: list[dict] | None = None) -> dict[str, ColumnPresentation]:
    """Use populated cells, not row 2: sparse Progress/date columns start blank."""
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if records is not None and sheet_name not in book.sheetnames:
            raise ValueError(f"Workbook {path.name!r} must contain worksheet {sheet_name!r}.")
        rows = book[sheet_name].iter_rows()
        headings = next(rows)
        names = [header(cell.value) for cell in headings]
        nonblank = [name.casefold() for name in names if name]
        if len(nonblank) != len(set(nonblank)):
            raise ValueError(f"Worksheet {sheet_name!r} contains duplicate column headers.")
        formats = [Counter() for _ in headings]
        alignments = [Counter() for _ in headings]

        def alignment(cell):
            value = cell.alignment
            return (value.horizontal, value.vertical, bool(value.wrapText),
                    bool(value.shrinkToFit), value.indent, value.textRotation)

        for row in rows:
            if records is not None:
                values = [cell.value for cell in row]
                if any(value is not None for value in values):
                    records.append(dict(zip(names, values)))
            for index, cell in enumerate(row[:len(headings)]):
                if cell.value is not None:
                    formats[index][cell.number_format] += 1
                    alignments[index][alignment(cell)] += 1
        result = {}
        for index, cell in enumerate(headings):
            if cell.value is None:
                continue
            result[presentation_key(str(cell.value), raw=sheet_name == "List")] = ColumnPresentation(
                formats[index].most_common(1)[0][0] if formats[index] else "General",
                alignments[index].most_common(1)[0][0] if alignments[index] else ColumnPresentation().alignment,
                alignment(cell),
                getattr(getattr(cell, "font", None), "name", None),
                getattr(getattr(cell, "font", None), "sz", None),
            )
        return result
    finally:
        book.close()


def column_presentations(
    columns, reference: dict[str, ColumnPresentation], source: dict[str, ColumnPresentation],
) -> dict[str, ColumnPresentation]:
    result = {}
    financial_format = reference.get("revenue", ColumnPresentation()).number_format
    if financial_format in {"General", "@"}:
        financial_format = BUILTIN_FORMATS[44]
    for name in columns:
        key = presentation_key(name)
        fallback = reference.get("status") if key == "yesterday status" else None
        model = reference.get(key, source.get(key, fallback or ColumnPresentation()))
        if name == "Progress":
            model = replace(model, number_format="0.0%")
        elif name == "Aging Of RFS":
            model = replace(model, number_format="0")
        elif name in FINANCIAL_FIELDS:
            model = replace(model, number_format=financial_format)
        elif name in DATE_FIELDS:
            model = replace(model, number_format=r"d\ mmm\ yyyy")
        if name in {"Aging Of RFS", "STATUS ORDER", "ON TIME DELIVERY", "MONTH COMPLETE"}:
            values = list(model.alignment)
            values[2] = False
            model = replace(model, alignment=tuple(values))
        elif name in {"ORDER TYPE", "PROCESS"}:
            values = list(model.alignment)
            values[2] = True
            model = replace(model, alignment=tuple(values))
        heading = list(model.header_alignment)
        heading[:2] = ["center", "center"]
        heading[4] = 0
        result[name] = replace(model, header_alignment=tuple(heading))
    return result


def apply_column_alignment(target, values: tuple) -> None:
    horizontal, vertical, wrap, shrink, indent, rotation = values
    horizontal_value = {
        None: 1, "general": 1, "left": -4131, "center": -4108, "right": -4152,
        "fill": 5, "justify": -4130, "centerContinuous": 7, "distributed": -4117,
    }[horizontal]
    vertical_value = {
        None: -4107, "bottom": -4107, "top": -4160, "center": -4108,
        "justify": -4130, "distributed": -4117,
    }[vertical]
    for name, value in {
        "HorizontalAlignment": horizontal_value, "VerticalAlignment": vertical_value,
        "WrapText": wrap, "ShrinkToFit": shrink, "IndentLevel": int(indent or 0),
        "Orientation": -4166 if rotation == 255 else int(rotation or 0),
    }.items():
        if getattr(target, name, None) != value:
            setattr(target, name, value)


def fit_unwrapped_text_columns(worksheet, body, frame: pd.DataFrame, presentations: dict) -> None:
    """Keep reference widths as a minimum; fit long names without growing unbounded."""
    for index, name in enumerate(frame.columns, 1):
        if presentations[name].alignment[2]:
            continue
        width = float(worksheet.Columns(index).ColumnWidth)
        longest = max((len(value) for value in frame[name] if isinstance(value, str)), default=0)
        if longest <= width:
            continue
        body.Columns(index).AutoFit()
        fitted = float(worksheet.Columns(index).ColumnWidth)
        worksheet.Columns(index).ColumnWidth = max(width, min(48, fitted))


def header(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\ufeff", "").strip())


def blank(value: object) -> bool:
    return value is None or value is pd.NA or (isinstance(value, float) and math.isnan(value))


def read_records(path: Path, sheet: str, *, presentations: dict | None = None) -> list[dict]:
    if presentations is not None:
        records = []
        presentations.update(read_column_presentations(path, sheet, records=records))
        return records
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in book.sheetnames:
            raise ValueError(f"Workbook {path.name!r} must contain worksheet {sheet!r}.")
        rows = book[sheet].iter_rows(values_only=True)
        names = [header(value) for value in next(rows)]
        nonblank = [name.casefold() for name in names if name]
        if len(nonblank) != len(set(nonblank)):
            raise ValueError(f"Worksheet {sheet!r} contains duplicate column headers.")
        return [dict(zip(names, row)) for row in rows if any(value is not None for value in row)]
    finally:
        book.close()


def is_list_export(path: Path) -> bool:
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if "List" not in book.sheetnames:
            return False
        names = {header(value) for value in next(book["List"].iter_rows(values_only=True))}
        missing = REQUIRED - names
        if missing:
            raise ValueError("IDE List export is missing columns: " + ", ".join(sorted(missing)))
        return True
    finally:
        book.close()


def normalize_status(value: object) -> str | None:
    if blank(value) or not str(value).strip():
        return None
    text = re.sub(r"^\d+\.\s*", "", str(value).strip())
    try:
        return STATUS[text.casefold()]
    except KeyError as exc:
        raise ValueError(f"Unrecognized IDE List Status: {value!r}") from exc


def target_bucket(value: object, report_date: date) -> str:
    if blank(value):
        return "Target Not Yet Inputted"
    if not isinstance(value, (date, datetime)):
        raise ValueError(f"Target RFS must be an Excel date, not {value!r}.")
    source, period = (value.year, value.month), (report_date.year, report_date.month)
    prefix = "Before" if source < period else "After" if source > period else "Target"
    return f"{prefix} {calendar.month_name[report_date.month]}"


def build_rows(raw: list[dict], previous: list[dict], report_date: date) -> tuple[pd.DataFrame, list[dict]]:
    from generate_ide_tracking import target_header

    if not raw:
        raise ValueError("IDE List export contains no records.")
    missing = REQUIRED - raw[0].keys()
    if missing:
        raise ValueError("IDE List export is missing columns: " + ", ".join(sorted(missing)))
    if not previous or not {"Quote ID", "Status", "ORDER TYPE"} <= previous[0].keys():
        raise ValueError("The IDE List export requires a new-format IDE reference with Status and ORDER TYPE.")
    keys = [str(row.get("Key") or "").strip() for row in raw]
    if not all(keys) or len(set(keys)) != len(keys):
        raise ValueError("IDE List Key must be populated and unique; duplicate Quote IDs are allowed.")
    history = {}
    for row in previous:
        quote = str(row.get("Quote ID") or "").strip()
        if quote:
            history.setdefault(quote, row)
    columns = [name.format(target=target_header(report_date)) for name in PREFIX]
    copied = set(columns) | {"Month Complete"}
    additional_columns = []
    for name in raw[0]:
        if name.casefold() == "order type":
            # Excel Table1 treats this raw header as the same as report ORDER TYPE.
            additional_columns.append((RAW_ORDER_TYPE, name))
        elif name and name not in copied:
            additional_columns.append((name, name))
    used_headers = {header(name).casefold() for name in columns}
    for name, _ in additional_columns:
        key = header(name).casefold()
        if key in used_headers:
            raise ValueError(f"IDE List column {name!r} conflicts with an ALL ORDER column name.")
        used_headers.add(key)
        columns.append(name)
    records, warnings = [], []
    for index, source in enumerate(raw, 2):
        row = dict(source)
        quote = str(source.get("Quote ID") or "").strip()
        if not quote:
            raise ValueError(f"Missing Quote ID at List row {index}.")
        old = history.get(quote)
        for name in DATE_FIELDS:
            value = source.get(name)
            if not blank(value) and not isinstance(value, (datetime, date)):
                raise ValueError(f"List row {index}: {name} must be an Excel date, not {value!r}.")
            if isinstance(value, (date, datetime)) and not 1900 <= value.year <= report_date.year + 10:
                warnings.append({"key": source["Key"], "field": name, "value": str(value),
                                 "reason": "Unusual source year preserved; no verified replacement available."})
        target, actual = source.get("Target RFS"), source.get("Actual RFS Date")
        target_date = target.date() if isinstance(target, datetime) else target
        aging = None if blank(target_date) else (report_date - target_date).days
        order_type = None if old is None else old.get("ORDER TYPE")
        if blank(order_type) or not str(order_type).strip():
            order_type = "Renewal" if re.search(r"\bRENEW", str(source.get("Quote Name") or ""), re.I) else "New Registration"
        row.update({
            target_header(report_date): target_bucket(target, report_date),
            "Aging Of RFS": aging,
            "STATUS ORDER": None if aging is None else "On Track" if aging < -5 else "Potential Delay" if aging < 0 else "Delay",
            "ON TIME DELIVERY": None if blank(target) or blank(actual) else "DELAY" if actual > target else "ONTIME",
            "Status": normalize_status(source.get("Status")),
            "Yesterday Status ": "#N/A" if old is None else normalize_status(old.get("Status")),
            "ORDER TYPE": order_type,
            "PROCESS": "New Reg" if order_type == "New Registration" else "Non New Reg",
            "YEAR FAB UPLOAD": "1900" if blank(source.get("FAB Upload Date")) else str(source["FAB Upload Date"].year),
            "MONTH COMPLETE": "Jan-00" if blank(source.get("SO Completion Date")) else source["SO Completion Date"].strftime("%b-%y"),
            "Month Complete2": source.get("Month Complete"),
        })
        row.update({name: source.get(source_name) for name, source_name in additional_columns})
        records.append({name: row.get(name) for name in columns})
    return pd.DataFrame(records, columns=columns, dtype=object), warnings


def _select(field, names: set[str] | None) -> None:
    field.ClearAllFilters()
    if names is None:
        return
    field.EnableMultiplePageItems = True
    items = list(field.PivotItems())
    available = {str(item.Name) for item in items}
    selected = available & names
    if not selected:
        # Excel cannot hide every item. Keep a missing-category filter in the
        # cache rather than silently showing unrelated records.
        raise ValueError(f"Pivot field {field.Name!r} has no requested items: {sorted(names)}")
    for item in items:
        if str(item.Name) in selected:
            item.Visible = True
    for item in items:
        if str(item.Name) not in selected:
            item.Visible = False
    if len(selected) == 1:
        field.EnableMultiplePageItems = False
        field.CurrentPage = next(iter(selected))


def _capture_filters(sheet) -> dict:
    states = {}
    for pivot in sheet.PivotTables():
        fields = {}
        for field in [*list(pivot.PageFields), *list(pivot.ColumnFields)]:
            if not isinstance(field.SourceName, str):
                continue  # Excel's synthetic Values axis is not a source field.
            name = str(field.SourceName)
            if name == "Data":
                continue
            if int(field.Orientation) == 3 and not field.EnableMultiplePageItems:
                current = str(field.CurrentPage)
                fields[name] = None if current == "(All)" else {current}
            else:
                items = list(field.PivotItems())
                fields[name] = None if all(item.Visible for item in items) else {str(item.Name) for item in items if item.Visible}
        states[str(pivot.Name)] = fields
    return states


def capture_progress_bar(table):
    progress = table.ListColumns("Progress").DataBodyRange
    for rule in progress.FormatConditions:
        if int(rule.Type) == 4:  # xlDataBar
            return rule
    # Older generated references left this 0-100% bar on a date column when
    # Progress moved. Recover only that recognizable rule, not unrelated CF.
    for rule in table.DataBodyRange.FormatConditions:
        if int(rule.Type) != 4:
            continue
        area = rule.AppliesTo
        index = int(area.Column) - int(table.Range.Column) + 1
        if (int(area.Columns.Count) == 1 and 1 <= index <= int(table.ListColumns.Count)
                and str(table.ListColumns(index).Name) in DATE_FIELDS
                and int(rule.MinPoint.Type) == int(rule.MaxPoint.Type) == 0
                and float(rule.MinPoint.Value) == 0 and float(rule.MaxPoint.Value) == 1
                and int(rule.BarColor.Color) == PROGRESS_BAR_COLOR):
            return rule
    return None


def restore_progress_bar(table, rule) -> None:
    progress = table.ListColumns("Progress").DataBodyRange
    if rule is None:
        rule = progress.FormatConditions.AddDatabar()
        rule.BarColor.Color = PROGRESS_BAR_COLOR
        rule.BarFillType = 0  # xlDataBarFillSolid
    else:
        rule.ModifyAppliesToRange(progress)
    rule.MinPoint.Modify(0, 0)  # xlConditionValueNumber, fixed 0-100% scale.
    rule.MaxPoint.Modify(0, 1)
    rule.PercentMin, rule.PercentMax, rule.ShowValue = 0, 100, True


def validate_progress_bar(sheet, column: int, row_count: int) -> None:
    expected = f"{get_column_letter(column)}2:{get_column_letter(column)}{row_count + 1}"
    rules = [(area, rule) for area in sheet.conditional_formatting
             for rule in sheet.conditional_formatting[area] if rule.type == "dataBar"
             and any(region.min_col <= column <= region.max_col for region in area.sqref.ranges)]
    if len(rules) != 1 or str(rules[0][0].sqref) != expected:
        raise RuntimeError(f"Saved IDE List Progress data bar must cover {expected} exactly.")
    points = rules[0][1].dataBar.cfvo
    if len(points) != 2 or [(point.type, point.val) for point in points] != [("num", 0), ("num", 1)]:
        raise RuntimeError("Saved IDE List Progress data bar must use a fixed 0-100% scale.")


def is_target_after_filter(fields: dict) -> bool:
    from generate_ide_tracking import target_item_kind

    return any(is_dynamic_target_header(name) and selected
               and {target_item_kind(value) for value in selected} == {"after"}
               for name, selected in fields.items())


def percentage_helper_formula(status_cells: dict[str, list[str]], denominator: str, *,
                              target_after: bool = False) -> str:
    complete = status_cells.get("5.SO Complete", [])
    cancel = status_cells.get("6.Cancel", [])
    if target_after:
        # Target After intentionally follows the confirmed manual precedence.
        completed = f'SUM({",".join(complete)})' if complete else "0"
        cancelled = f'SUM({",".join(cancel)})' if cancel else "0"
        expression = f"{completed}+{cancelled}/{denominator}"
    else:
        expression = f'({"+".join(complete + cancel) or "0"})/{denominator}'
    return f"=IF({denominator}=0,0,{expression})"


def _restore_percentages(sheet, scratch, blocks) -> None:
    from generate_ide_tracking import restore_pivot_percentage_blocks

    restore_pivot_percentage_blocks(sheet, scratch, blocks)
    filters = _capture_filters(sheet)
    for block in blocks:
        phase = sheet.PivotTables(block.pivot_name)
        target_after = is_target_after_filter(filters[str(phase.Name)])
        region = phase.TableRange1
        summaries = [p for p in sheet.PivotTables() if int(p.TableRange1.Column) < int(region.Column)
                     and {str(f.SourceName) for f in p.RowFields} == {str(f.SourceName) for f in phase.RowFields}
                     and any(str(f.SourceName) == "Quote ID" and int(f.Function) == -4112 for f in p.DataFields)]
        if not summaries:
            raise RuntimeError(f"No quote-count summary for {phase.Name}.")
        summary = min(summaries, key=lambda p: abs(int(p.TableRange1.Row) - int(region.Row)))
        denominators = {}
        summary_fields, phase_fields = {}, {}
        for row in summary.DataBodyRange.Rows:
            for cell in row.Cells:
                if str(cell.PivotCell.DataField.SourceName) == "Quote ID":
                    denominators[pivot_row_identity(cell, field_names=summary_fields)] = cell.Address
        body = phase.DataBodyRange
        percentage_column = int(region.Column + region.Columns.Count)
        if body is None:
            target = sheet.Cells(int(region.Row + region.Rows.Count - 1), percentage_column)
            target.Formula = "=0"
            target.NumberFormat = "0.00%"
            continue
        status_columns = None
        for row in body.Rows:
            if status_columns is None:
                status_columns = {}
                for cell in row.Cells:
                    items = cell.PivotCell.ColumnItems
                    statuses = [str(items(i).Name) for i in range(1, int(items.Count) + 1)]
                    for status in ("5.SO Complete", "6.Cancel"):
                        if status in statuses:
                            status_columns.setdefault(status, []).append(int(cell.Column))
            row_number = int(row.Row)
            status_cells = {status: [f"${get_column_letter(column)}${row_number}" for column in columns]
                            for status, columns in status_columns.items()}
            identity = pivot_row_identity(row.Cells(1, 1), field_names=phase_fields)
            denominator = denominators.get(identity)
            if denominator is None:
                raise RuntimeError(f"Missing percentage denominator for {phase.Name}: {identity!r}")
            target = sheet.Cells(int(row.Row), percentage_column)
            target.Formula = percentage_helper_formula(
                status_cells, denominator, target_after=target_after)
            target.NumberFormat = "0.00%"


def _validate_pivot_totals(sheet, frame: pd.DataFrame) -> None:
    """Reconcile native pivot totals independently against their current filters."""
    records = frame.to_dict("records")
    states = _capture_filters(sheet)
    for pivot in sheet.PivotTables():
        selected = records
        for name, allowed in states[str(pivot.Name)].items():
            if allowed is not None:
                selected = [row for row in selected if ("(blank)" if blank(row.get(name)) or row.get(name) == ""
                                                       else str(row[name])) in allowed]
        body = pivot.DataBodyRange
        if body is None:
            if selected:
                raise RuntimeError(f"Pivot {pivot.Name} has no body for {len(selected)} eligible rows.")
            continue
        totals = body.Rows(int(body.Rows.Count))
        values = [cell.Value2 for cell in totals.Cells]
        measures = list(pivot.DataFields)
        if len(measures) == 2:
            revenue = [row.get("Revenue") for row in selected if not blank(row.get("Revenue"))]
            expected = [sum(revenue) if revenue else None, len(selected) or None]
            equal = all((a is None and b in (None, 0)) or (isinstance(a, (int, float)) and isinstance(b, (int, float))
                        and math.isclose(a, b, rel_tol=1e-12, abs_tol=1e-5)) for a, b in zip(expected, values))
            if len(expected) != len(values) or not equal:
                raise RuntimeError(f"Pivot {pivot.Name} totals {values!r} differ from source {expected!r}.")
        elif int(measures[0].Calculation) == 8:
            # The rightmost total is 100%; each phase is independently checked.
            for cell in totals.Cells:
                items = cell.PivotCell.ColumnItems
                categories = [str(items(i).Name) for i in range(1, int(items.Count) + 1)]
                numerator = sum(row.get("ON TIME DELIVERY") in categories for row in selected) if categories else len(selected)
                expected = numerator / len(selected) if selected else None
                if expected is None and cell.Value2 in (None, 0):
                    continue
                if expected is None or not math.isclose(float(cell.Value2 or 0), expected, rel_tol=1e-12, abs_tol=1e-10):
                    raise RuntimeError(f"Pivot {pivot.Name} percentage does not reconcile with source rows.")
        else:
            actual = values[-1] if pivot.RowGrand else sum(value or 0 for value in values)
            if (actual or 0) != len(selected):
                raise RuntimeError(f"Pivot {pivot.Name} count {actual} differs from source count {len(selected)}.")


def _write_native(path: Path, frame: pd.DataFrame, report_date: date, *, raw_path: Path | None = None,
                  reference_presentations: dict | None = None, source_presentations: dict | None = None,
                  timings: dict | None = None) -> int:
    import pythoncom
    import win32com.client
    from generate_ide_tracking import (
        capture_pivot_percentage_blocks, com_value, contiguous_row_ranges, normalize_pivot_section_row_visibility,
        update_pivot_period_titles, target_header, normalize_pivot_percentage_borders,
    )
    from excel_pivot_layout import resize_dynamic_pivot_section_banners
    from ide_pivot_presentation import read_presentation_grid, capture_presentation, restore_presentation

    stage_start = perf_counter()

    def mark_stage(name):
        nonlocal stage_start
        now = perf_counter()
        if timings is not None:
            timings[name] = now - stage_start
        stage_start = now

    if reference_presentations is None:
        reference_presentations = read_column_presentations(path, "ALL ORDER")
    if source_presentations is None:
        source_presentations = read_column_presentations(raw_path, "List") if raw_path is not None else {}
    presentations = column_presentations(frame.columns, reference_presentations, source_presentations)
    report_font = source_presentations.get("key", reference_presentations.get("key", ColumnPresentation()))
    pivot_grid, pivot_rows = read_presentation_grid(path)
    mark_stage("read_pivot_presentation")
    pythoncom.CoInitialize()
    excel = book = anchor = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = excel.DisplayAlerts = excel.ScreenUpdating = excel.EnableEvents = False
        anchor = excel.Workbooks.Add()
        excel.Calculation = -4135
        book = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        mark_stage("open_excel")
        data, sheet = book.Worksheets("ALL ORDER"), book.Worksheets("PIVOT")
        if not data.ListObjects.Count or str(data.ListObjects(1).Name) != "Table1":
            raise ValueError("The new IDE format requires a reference with Excel table Table1.")
        table = data.ListObjects("Table1")
        progress_bar = capture_progress_bar(table)
        filters = _capture_filters(sheet)
        pivot_presentation = capture_presentation(sheet, pivot_grid, pivot_rows)
        widths = {i: float(sheet.Columns(i).ColumnWidth) for i in range(1, int(sheet.UsedRange.Columns.Count) + 1)}
        scratch, blocks = capture_pivot_percentage_blocks(book, sheet)
        mark_stage("capture_reference")
        old_headers = [str(cell.Value2) for cell in table.HeaderRowRange.Cells]
        header_styles = {
            presentation_key(str(cell.Value2)): (cell.Interior.Color, cell.Font.Color, cell.Font.Bold)
            for cell in table.HeaderRowRange.Cells
        }
        old_target = next(name for name in old_headers if is_dynamic_target_header(name))
        new_target = target_header(report_date)
        # Bind by field name before renaming the monthly field. Reusing positional
        # cache indices would attach pivots to different columns after reordering.
        write_headers = [old_target if name == new_target else name for name in frame.columns]
        write_header_keys = {header(name).casefold() for name in write_headers}
        migration_headers = [name for name in old_headers if header(name).casefold() not in write_header_keys]
        cache_headers = write_headers + migration_headers
        styles = {presentation_key(name): data.Cells(1, i).EntireColumn.ColumnWidth
                  for i, name in enumerate(old_headers, 1)}
        data.Range(data.Cells(2, 1), data.Cells(max(int(table.Range.Rows.Count), len(frame) + 1),
                                             max(len(old_headers), len(frame.columns)))).ClearContents()
        table.Resize(data.Range(data.Cells(1, 1), data.Cells(len(frame) + 1, len(cache_headers))))
        table.HeaderRowRange.Value2 = (tuple(cache_headers),)
        body = table.DataBodyRange
        body.NumberFormat = "General"
        for i, name in enumerate(frame.columns, 1):
            column = body.Columns(i)
            if any(isinstance(value, str) for value in frame[name] if not blank(value)):
                column.NumberFormat = "@"
            elif name in DATE_FIELDS:
                column.NumberFormat = "dd/mm/yyyy hh:mm" if name in {"Updated", "Resolved", "Created"} else "dd/mm/yyyy"
            key = presentation_key(name)
            fallback_width = styles.get("month complete", 18) if key == "yesterday status" else 18
            data.Columns(i).ColumnWidth = styles.get(key, fallback_width)
        body.Value2 = tuple(tuple(com_value(value) for value in row) + (None,) * len(migration_headers)
                            for row in frame.itertuples(index=False, name=None))
        col = {name: i for i, name in enumerate(frame.columns, 1)}
        # History is deliberately #N/A only for quotes absent from the reference.
        for i, value in enumerate(frame["Yesterday Status "], 2):
            if value == "#N/A":
                data.Cells(i, col["Yesterday Status "]).NumberFormat = "General"
                data.Cells(i, col["Yesterday Status "]).Formula = "=NA()"
        for name, formula in {
            "YEAR FAB UPLOAD": f'=TEXT(RC{col["FAB Upload Date"]},"yyyy")',
            "MONTH COMPLETE": f'=TEXT(RC{col["SO Completion Date"]},"MMM-YY")',
            "Aging Of RFS": f'=IF(RC{col["Target RFS"]}="","",DAYS(DATE({report_date.year},{report_date.month},{report_date.day}),RC{col["Target RFS"]}))',
            "STATUS ORDER": f'=IF(RC{col["Aging Of RFS"]}="","",IF(RC{col["Aging Of RFS"]}<-5,"On Track",IF(RC{col["Aging Of RFS"]}<0,"Potential Delay","Delay")))',
            "ON TIME DELIVERY": f'=IF(OR(RC{col["Actual RFS Date"]}="",RC{col["Target RFS"]}=""),"",IF(RC{col["Actual RFS Date"]}>RC{col["Target RFS"]},"DELAY","ONTIME"))',
        }.items():
            body.Columns(col[name]).NumberFormat = "General"
            body.Columns(col[name]).FormulaR1C1 = formula
        # A formula returning "" is a different PivotCache item from a genuinely
        # empty cell. Preserve the manual workbook's (blank) category.
        for name in ("Aging Of RFS", "STATUS ORDER", "ON TIME DELIVERY"):
            missing_rows = [i for i, value in enumerate(frame[name], 2) if blank(value)]
            for start, end in contiguous_row_ranges(missing_rows):
                data.Range(data.Cells(start, col[name]), data.Cells(end, col[name])).ClearContents()
        # Text protection is only for ingestion. Restore display formats after
        # writing so General/month labels cannot trigger Excel's date conversion.
        for name, presentation in presentations.items():
            column = body.Columns(col[name])
            column.NumberFormat = presentation.number_format
            apply_column_alignment(column, presentation.alignment)
            header_cell = data.Cells(1, col[name])
            apply_column_alignment(header_cell, presentation.header_alignment)
            header_cell_font = header_cell.Font
            key = presentation_key(name)
            style = header_styles.get(key)
            if style is None:
                style = header_styles.get("month complete" if key == "yesterday status" else "key")
            if style is not None:
                header_cell.Interior.Color, header_cell_font.Color, header_cell_font.Bold = style
            header_font = source_presentations.get(key)
            if header_font is None:
                header_font = report_font if key == "yesterday status" else reference_presentations.get(key)
            if header_font is not None:
                if header_font.font_name and header_cell_font.Name != header_font.font_name:
                    header_cell_font.Name = header_font.font_name
                if header_font.font_size and header_cell_font.Size != header_font.font_size:
                    header_cell_font.Size = header_font.font_size
        if report_font.font_name and body.Font.Name != report_font.font_name:
            body.Font.Name = report_font.font_name
        if report_font.font_size and body.Font.Size != report_font.font_size:
            body.Font.Size = report_font.font_size
        fit_unwrapped_text_columns(data, body, frame, presentations)
        data.Calculate()
        mark_stage("write_data_and_formats")
        for pivot in sheet.PivotTables():
            pivot.PreserveFormatting, pivot.HasAutoFormat = True, False
            pivot.ManualUpdate = True
        cache = book.PivotCaches().Create(SourceType=1, SourceData="Table1")
        cache.MissingItemsLimit = 0
        cache.RefreshOnFileOpen = False
        for index, pivot in enumerate(list(sheet.PivotTables())):
            if index == 0:
                pivot.ChangePivotCache(cache)
            else:
                # Excel materializes a new cache on the first binding. Subsequent
                # pivots must share its live index, which Excel can renumber.
                pivot.CacheIndex = int(cache.Index)
            cache = pivot.PivotCache()
        if migration_headers:
            table.Resize(data.Range(data.Cells(1, 1), data.Cells(len(frame) + 1, len(frame.columns))))
            data.Range(data.Cells(1, len(frame.columns) + 1),
                       data.Cells(len(frame) + 1, len(cache_headers))).Clear()
        if old_target != new_target:
            table.ListColumns(old_target).Name = new_target
        restore_progress_bar(table, progress_bar)
        cache.Refresh()
        for pivot in sheet.PivotTables():
            for old_name, selected in filters[str(pivot.Name)].items():
                name = new_target if old_name == old_target else old_name
                if old_name == old_target and selected is not None:
                    from generate_ide_tracking import target_item_kind, target_item_label
                    selected = {target_item_label(target_item_kind(value), report_date) for value in selected}
                _select(pivot.PivotFields(name), selected)
            pivot.ManualUpdate = False
            pivot.RefreshTable()
        mark_stage("refresh_pivots")
        _validate_pivot_totals(sheet, frame)
        _restore_percentages(sheet, scratch, blocks)
        mark_stage("validate_pivots_and_percentages")
        normalize_pivot_percentage_borders(sheet, blocks)
        for block in blocks:
            pivot = sheet.PivotTables(block.pivot_name)
            region = pivot.TableRange2
            header_row = int(region.Row) + block.header_row_offset
            percentage_column = int(region.Column + region.Columns.Count)
            sheet.Cells(header_row, percentage_column).Borders(9).LineStyle = -4142
            border = sheet.Cells(header_row+1, percentage_column).Borders(8)
            border.LineStyle, border.Weight, border.Color = 1, -4138, 255
        scratch.Visible = -1
        scratch.Delete()
        update_pivot_period_titles(sheet, report_date)
        used = sheet.UsedRange
        for r, row in enumerate(used.Value2, int(used.Row)):
            for c, value in enumerate(row, int(used.Column)):
                if isinstance(value, str) and value.startswith("TARGET RFS ORDER AGING"):
                    sheet.Cells(r, c).Value2 = ("TARGET RFS ORDER AGING START FROM PRA-DELIVERY - WAITING AOS STATUS "
                                   f"{calendar.month_name[report_date.month].upper()} {report_date:%y}")
        normalize_pivot_section_row_visibility(sheet)
        for i, width in widths.items():
            sheet.Columns(i).ColumnWidth = width
        restore_presentation(sheet, pivot_presentation)
        resize_dynamic_pivot_section_banners(
            sheet,
            title_markers=("target complete", "target rfs order aging", "ontime delivery",
                           "delay completion", "target after", "target not inputted"),
            side_header_markers=("percentage of completion",),
        )
        mark_stage("restore_layout")
        book.Activate()
        window = book.Windows(1)
        window.Activate()
        data.Activate()
        window.ScrollRow = 1
        window.ScrollColumn = 5
        excel.Calculation = -4105
        excel.Calculate()
        try:
            errors = sheet.UsedRange.SpecialCells(-4123, 16)
        except Exception:
            errors = None
        if errors is not None:
            raise RuntimeError(f"IDE List PIVOT contains formula errors: {errors.Address}")
        count = int(sheet.PivotTables().Count)
        book.Save()
        mark_stage("calculate_and_save")
        return count
    finally:
        for owned_book in (book, anchor):
            if owned_book is not None:
                try:
                    owned_book.Close(SaveChanges=False)
                except Exception:
                    pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        mark_stage("close_excel")


def validate_saved(path: Path, frame: pd.DataFrame, pivot_count: int) -> dict:
    rows = read_records(path, "ALL ORDER")
    if len(rows) != len(frame):
        raise RuntimeError("Saved IDE List record count differs from input.")
    for index, (saved, expected) in enumerate(zip(rows, frame.to_dict("records")), 2):
        for name, value in expected.items():
            actual = saved.get(header(name))
            if (blank(value) or value == "") and (blank(actual) or actual == ""):
                continue
            if isinstance(value, (int, float)) and isinstance(actual, (int, float)):
                equal = math.isclose(value, actual, rel_tol=1e-12, abs_tol=1e-8)
            else:
                equal = value == actual
            if not equal:
                raise RuntimeError(f"Saved IDE List mismatch at {get_column_letter(frame.columns.get_loc(name)+1)}{index} ({name}): {actual!r} != {value!r}")
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        from ide_pivot_presentation import validate_centered_header

        for cell in book["ALL ORDER"][1][:len(frame.columns)]:
            validate_centered_header(cell)
    finally:
        book.close()
    with ZipFile(path) as archive:
        data_part = worksheet_part(archive, "ALL ORDER")
        data_root = ET.fromstring(archive.read(data_part))
        validate_progress_bar(SimpleNamespace(conditional_formatting=conditional_formats(data_root)),
                              frame.columns.get_loc("Progress") + 1, len(frame))
        validate_table_and_caches(archive, data_root, data_part,
                                  f"A1:{get_column_letter(len(frame.columns))}{len(frame)+1}", len(frame), pivot_count)
    return {"rows": len(rows), "columns": len(frame.columns), "pivot_tables": pivot_count}


def generate(raw_path: Path, reference: Path, output: Path, report_date: date) -> dict:
    started = perf_counter()
    raw_presentations, reference_presentations, timings = {}, {}, {}
    raw = read_records(raw_path, "List", presentations=raw_presentations)
    previous = read_records(reference, "ALL ORDER", presentations=reference_presentations)
    timings["read_inputs"] = perf_counter() - started
    stage_start = perf_counter()
    frame, warnings = build_rows(raw, previous, report_date)
    timings["build_rows"] = perf_counter() - stage_start
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.resolve() in {raw_path.resolve(), reference.resolve()}:
        raise ValueError("Output must not overwrite the raw or reference workbook.")
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=output.parent, prefix="ide_list_", suffix=".xlsx", delete=False) as handle:
            temporary = Path(handle.name)
        shutil.copy2(reference, temporary)
        count = _write_native(temporary, frame, report_date, raw_path=raw_path,
                              reference_presentations=reference_presentations, source_presentations=raw_presentations,
                              timings=timings)
        stage_start = perf_counter()
        validation = validate_saved(temporary, frame, count)
        timings["validate_saved"] = perf_counter() - stage_start
        os.replace(temporary, output)
        temporary = None
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    quotes = Counter(frame["Quote ID"])
    timings["total"] = perf_counter() - started
    LOGGER.info("IDE List stage timings (seconds): %s", timings)
    return {**validation, "format_profile": "ide-list", "refreshed_pivots": count, "timings_seconds": timings,
            "financial_source_policy": "raw-revenue", "source_warnings": warnings,
            "collabs_fallback_rows": 0, "zero_fallback_rows": 0, "date_fallback_rows": 0,
            "field_fallback_rows": 0, "invalid_dates": {},
            "duplicate_quote_ids": [quote for quote, count in quotes.items() if count > 1],
            "duplicate_rows_preserved": sum(count - 1 for count in quotes.values()),
            "historical_status_source": "first-quote-reference-status", "aging_policy": "report-date-target-rfs"}
