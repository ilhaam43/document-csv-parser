"""Write weekly datasets into the existing Excel template with native Excel COM."""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import time
import uuid
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from excel_automation_lock import excel_process_lock
from pipeline_stage_runner import wait_for_xlsx

from .constants import (
    DATE_COLUMNS,
    DECIMAL_COLUMNS,
    INTEGER_COLUMNS,
    MASTER_COLUMNS,
    MASTER_ADJUSTMENT_REVIEW_PIVOTS,
    MASTER_SHEET_NAME,
    MASTER_TABLE_NAME,
    MOBILE_ONLY_REMOVED_PIVOTS,
    MOBILE_ADJUSTMENT_REVIEW_PIVOTS,
    PIVOT_PERFORMANCE_SHEET_NAME,
)
from .dates import excel_week_number
from .io import iter_master_csv
from .slide import completion_month, load_slide_review, reporting_week_start, week_label
from .pivot_review import PivotReviewDecision, load_pivot_review


EXCEL_LOCK_TIMEOUT_SECONDS = 60 * 60
WRITE_CHUNK_ROWS = 2_000
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
EXCEL_EPOCH = datetime(1899, 12, 30)
TEXT_COLUMNS = frozenset(MASTER_COLUMNS) - DATE_COLUMNS - INTEGER_COLUMNS - DECIMAL_COLUMNS


def _excel_serial(value: date | datetime) -> float:
    if isinstance(value, datetime):
        resolved = value.replace(tzinfo=None)
    else:
        resolved = datetime.combine(value, datetime.min.time())
    return (resolved - EXCEL_EPOCH).total_seconds() / 86_400


def _com_value(column: str, value: Any) -> Any:
    if value is None:
        return None
    if column in DATE_COLUMNS and isinstance(value, (date, datetime)):
        # Passing datetime directly through pywin32 applies the Windows timezone
        # offset. Excel serials preserve the source date and time exactly.
        return _excel_serial(value)
    if isinstance(value, Decimal):
        return float(value)
    return value


def _sheet_by_name(workbook, name: str):
    for sheet in workbook.Worksheets:
        if str(sheet.Name).casefold() == name.casefold():
            return sheet
    return None


def _pivot_manifest(workbook) -> dict[str, set[str]]:
    manifest: dict[str, set[str]] = {}
    for sheet in workbook.Worksheets:
        tables = sheet.PivotTables()
        manifest[str(sheet.Name)] = {
            str(tables(index).Name) for index in range(1, tables.Count + 1)
        }
    return manifest


def _flatten_range_values(values: Any) -> tuple[Any, ...]:
    if values is None:
        return ()
    if not isinstance(values, tuple):
        return (values,)
    flattened: list[Any] = []
    for value in values:
        if isinstance(value, tuple):
            flattened.extend(value)
        else:
            flattened.append(value)
    return tuple(flattened)


def _changed_adjustments(
    old_adjustments: tuple[Any, ...],
    old_achievements: tuple[Any, ...],
    records: list[dict[str, Any]],
) -> set[str]:
    changed: set[str] = set()
    for index, record in enumerate(records):
        adjustment = record.get("ADJUSTMENT")
        if adjustment in (None, ""):
            continue
        old_adjustment = old_adjustments[index] if index < len(old_adjustments) else None
        old_achievement = old_achievements[index] if index < len(old_achievements) else None
        if (
            str(adjustment) != str(old_adjustment or "")
            or str(record.get("ACHIEVEMENT") or "") != str(old_achievement or "")
        ):
            changed.add(str(adjustment))
    return changed


def _latest_report_week(records: list[dict[str, Any]]) -> int:
    fab_dates = [
        value
        for record in records
        if isinstance((value := record.get("fab_upload")), (date, datetime))
    ]
    if not fab_dates:
        raise ValueError("Cannot infer report week because fab_upload has no dates.")
    return excel_week_number(max(fab_dates))


def _write_table(sheet, records: list[dict[str, Any]]) -> set[str]:
    table = sheet.ListObjects(MASTER_TABLE_NAME)
    old_adjustments = _flatten_range_values(
        table.ListColumns("ADJUSTMENT").DataBodyRange.Value2
    )
    old_achievements = _flatten_range_values(
        table.ListColumns("ACHIEVEMENT").DataBodyRange.Value2
    )
    changed_adjustments = _changed_adjustments(
        old_adjustments,
        old_achievements,
        records,
    )
    column_count = len(MASTER_COLUMNS)
    target_last_row = len(records) + 1
    old_last_row = int(table.Range.Row + table.Range.Rows.Count - 1)
    clear_last_row = max(old_last_row, target_last_row)
    sheet.Range(sheet.Cells(2, 1), sheet.Cells(clear_last_row, column_count)).ClearContents()

    target_range = sheet.Range(
        sheet.Cells(1, 1),
        sheet.Cells(target_last_row, column_count),
    )
    table.Resize(target_range)
    for column_number, header in enumerate(MASTER_COLUMNS, start=1):
        sheet.Cells(1, column_number).Value = header

    original_number_formats = {
        column_number: sheet.Cells(2, column_number).NumberFormat
        for column_number in range(1, column_count + 1)
    }
    if records:
        for column_number, column in enumerate(MASTER_COLUMNS, start=1):
            if column not in TEXT_COLUMNS:
                continue
            sheet.Range(
                sheet.Cells(2, column_number),
                sheet.Cells(target_last_row, column_number),
            ).NumberFormat = "@"

    for start in range(0, len(records), WRITE_CHUNK_ROWS):
        chunk = records[start : start + WRITE_CHUNK_ROWS]
        values = tuple(
            tuple(_com_value(column, record.get(column)) for column in MASTER_COLUMNS)
            for record in chunk
        )
        first_row = start + 2
        last_row = first_row + len(chunk) - 1
        sheet.Range(
            sheet.Cells(first_row, 1),
            sheet.Cells(last_row, column_count),
        ).Value = values

    if records:
        for column_number, column in enumerate(MASTER_COLUMNS, start=1):
            if column not in TEXT_COLUMNS:
                continue
            sheet.Range(
                sheet.Cells(2, column_number),
                sheet.Cells(target_last_row, column_number),
            ).NumberFormat = original_number_formats[column_number]
    return changed_adjustments


def _remove_mobile_only_pivots(workbook) -> None:
    sheet = _sheet_by_name(workbook, PIVOT_PERFORMANCE_SHEET_NAME)
    if sheet is None:
        raise RuntimeError(f"Template has no {PIVOT_PERFORMANCE_SHEET_NAME!r} sheet.")
    for pivot_name in MOBILE_ONLY_REMOVED_PIVOTS:
        try:
            pivot = sheet.PivotTables(pivot_name)
        except Exception:
            continue
        pivot.TableRange2.Clear()


def _include_new_adjustments(
    workbook,
    values: set[str],
    *,
    mobile: bool,
) -> set[tuple[str, str]]:
    if not values and not mobile:
        return set()
    target_pivots = (
        MOBILE_ADJUSTMENT_REVIEW_PIVOTS if mobile else MASTER_ADJUSTMENT_REVIEW_PIVOTS
    )
    changed_pivots: set[tuple[str, str]] = set()
    for sheet in workbook.Worksheets:
        pivot_tables = sheet.PivotTables()
        for pivot_number in range(1, pivot_tables.Count + 1):
            pivot = pivot_tables(pivot_number)
            pivot_key = (str(sheet.Name), str(pivot.Name))
            if pivot_key not in target_pivots:
                continue
            try:
                field = pivot.PivotFields("ADJUSTMENT")
            except Exception as exc:
                raise RuntimeError(
                    f"Adjustment pivot {sheet.Name!r}/{pivot.Name!r} has no ADJUSTMENT field."
                ) from exc
            if mobile:
                try:
                    if not bool(field.AllItemsVisible):
                        field.ClearAllFilters()
                        changed_pivots.add(pivot_key)
                except Exception as exc:
                    raise RuntimeError(
                        f"Cannot reset Mobile adjustment filter on {sheet.Name!r}/{pivot.Name!r}."
                    ) from exc
                continue
            try:
                if not bool(field.EnableMultiplePageItems):
                    continue
            except Exception:
                continue
            for value in values:
                try:
                    field.PivotItems(value).Visible = True
                    changed_pivots.add(pivot_key)
                except Exception:
                    continue
    return changed_pivots


def _week_number(value: Any) -> int | None:
    match = re.fullmatch(r"W(\d{1,2})", str(value or "").strip(), flags=re.IGNORECASE)
    if match is None:
        return None
    week = int(match.group(1))
    return week if 1 <= week <= 53 else None


def _apply_mobile_week_window(
    workbook,
    report_week: int,
) -> set[tuple[str, str]]:
    first_visible_week = max(1, report_week - 3)
    changed_pivots: set[tuple[str, str]] = set()
    sheet = _sheet_by_name(workbook, PIVOT_PERFORMANCE_SHEET_NAME)
    if sheet is None:
        raise RuntimeError(f"Template has no {PIVOT_PERFORMANCE_SHEET_NAME!r} sheet.")
    pivot_tables = sheet.PivotTables()
    for pivot_number in range(1, pivot_tables.Count + 1):
        pivot = pivot_tables(pivot_number)
        try:
            field = pivot.PivotFields("Week Completion")
            if int(field.Orientation) not in (1, 2):
                continue
            items = field.PivotItems()
        except Exception:
            continue
        desired_items: list[tuple[Any, bool]] = []
        for item_number in range(1, items.Count + 1):
            item = items(item_number)
            week = _week_number(getattr(item, "Name", None))
            if week is not None:
                desired_items.append((item, first_visible_week <= week <= report_week))
        if not desired_items:
            continue
        changed = False
        for item, desired_visible in desired_items:
            if not desired_visible:
                continue
            try:
                if not bool(item.Visible):
                    item.Visible = True
                    changed = True
            except Exception:
                continue
        for item, desired_visible in desired_items:
            if desired_visible:
                continue
            try:
                if bool(item.Visible):
                    item.Visible = False
                    changed = True
            except Exception:
                continue
        if changed:
            changed_pivots.add((str(sheet.Name), str(pivot.Name)))
    return changed_pivots


def _order_week_items(workbook) -> set[tuple[str, str]]:
    changed_pivots: set[tuple[str, str]] = set()
    for sheet in workbook.Worksheets:
        pivot_tables = sheet.PivotTables()
        for pivot_number in range(1, pivot_tables.Count + 1):
            pivot = pivot_tables(pivot_number)
            for field_name in ("Week FAB", "Week Completion"):
                try:
                    field = pivot.PivotFields(field_name)
                    if int(field.Orientation) not in (1, 2):
                        continue
                    items = field.PivotItems()
                except Exception:
                    continue
                week_items: list[tuple[int, int, Any]] = []
                for item_number in range(1, items.Count + 1):
                    item = items(item_number)
                    try:
                        if not bool(item.Visible):
                            continue
                    except Exception:
                        continue
                    week = _week_number(getattr(item, "Name", None))
                    if week is None:
                        continue
                    try:
                        position = int(item.Position)
                    except Exception:
                        position = item_number
                    week_items.append((position, week, item))
                current_order = [week for _, week, _ in sorted(week_items)]
                if current_order == sorted(current_order):
                    continue
                field_changed = False
                for position, (_, _, item) in enumerate(
                    sorted(week_items, key=lambda value: value[1]),
                    start=1,
                ):
                    try:
                        item.Position = position
                        field_changed = True
                    except Exception:
                        continue
                if field_changed:
                    changed_pivots.add((str(sheet.Name), str(pivot.Name)))
    return changed_pivots


def _refresh_selected_pivots(workbook, pivot_keys: set[tuple[str, str]]) -> None:
    for sheet_name, pivot_name in sorted(pivot_keys):
        sheet = _sheet_by_name(workbook, sheet_name)
        if sheet is None:
            raise RuntimeError(f"Pivot sheet disappeared after refresh: {sheet_name!r}.")
        sheet.PivotTables(pivot_name).RefreshTable()


def _leadtime_item_visible(value: Any) -> bool:
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return False
    return number.is_finite() and number >= 0


def _apply_nonnegative_leadtime_filters(workbook) -> set[tuple[str, str]]:
    changed: set[tuple[str, str]] = set()
    for sheet in workbook.Worksheets:
        for index in range(1, sheet.PivotTables().Count + 1):
            pivot = sheet.PivotTables(index)
            try:
                field = pivot.PivotFields("LEADTIME FAB-CASH")
                if int(field.Orientation) != 3 or not bool(field.EnableMultiplePageItems):
                    continue
            except Exception:
                continue
            items = field.PivotItems()
            desired = []
            valid_records = False
            for item_number in range(1, items.Count + 1):
                item = items(item_number)
                visible = _leadtime_item_visible(item.Name)
                current = bool(item.Visible)
                if visible and current:
                    valid_records = True
                if current != visible and int(item.RecordCount) > 0:
                    desired.append((item, visible))
                    valid_records = valid_records or visible
            if not desired or not valid_records:
                continue
            # Reveal valid numbers first: Excel cannot hide a field's last item.
            original_manual = bool(pivot.ManualUpdate)
            pivot.ManualUpdate = True
            try:
                for item, visible in sorted(desired, key=lambda pair: not pair[1]):
                    try:
                        item.Visible = visible
                    except Exception as exc:
                        raise RuntimeError(
                            f"Cannot apply lead-time filter on {sheet.Name!r}/{pivot.Name!r}: {item.Name!r}."
                        ) from exc
                    changed.add((str(sheet.Name), str(pivot.Name)))
            finally:
                pivot.ManualUpdate = original_manual
    return changed


def _master_month_rollover(workbook, report_year: int, report_week: int) -> bool:
    sheet = _sheet_by_name(workbook, "SLIDE GH MEETING")
    if sheet is None:
        return False
    expanded = set()
    for index in range(1, sheet.PivotTables().Count + 1):
        pivot = sheet.PivotTables(index)
        try:
            if int(pivot.PivotFields("Week Completion").Orientation) != 1:
                continue
        except Exception:
            continue
        area = pivot.TableRange1
        labels = _flatten_range_values(sheet.Range(
            sheet.Cells(int(area.Row), int(area.Column)),
            sheet.Cells(int(area.Row + area.Rows.Count - 1), int(area.Column)),
        ).Value2)
        parent_month = None
        for value in labels:
            label = str(value or "").strip()
            month = completion_month(label)
            if month is not None:
                parent_month = month
            elif _week_number(label) is not None and parent_month is not None:
                expanded.add(parent_month)
    target = reporting_week_start(report_year, report_week).replace(day=1)
    return len(expanded) == 1 and next(iter(expanded)) < target


@dataclass(frozen=True)
class PivotHelperFormula:
    sheet: str
    pivot: str
    column: int
    formula: str
    number_format: str
    first_row: int = 0
    last_row: int = 0


def _uniform_relative_formula(values: tuple[Any, ...]) -> str | None:
    if any(value not in (None, "") and not (isinstance(value, str) and value.startswith("="))
           for value in values):
        return None
    formulas = {value for value in values if isinstance(value, str) and value.startswith("=")}
    return next(iter(formulas)) if len(formulas) == 1 else None


def _capture_pivot_helpers(workbook) -> list[PivotHelperFormula]:
    helpers: list[PivotHelperFormula] = []
    for sheet in workbook.Worksheets:
        for index in range(1, sheet.PivotTables().Count + 1):
            pivot = sheet.PivotTables(index)
            body = pivot.DataBodyRange
            if body is None:
                continue
            area = pivot.TableRange1
            column = int(area.Column + area.Columns.Count)
            first_row = int(body.Row)
            last_row = int(area.Row + area.Rows.Count - 1)
            cells = sheet.Range(sheet.Cells(first_row, column), sheet.Cells(last_row, column))
            values = _flatten_range_values(cells.FormulaR1C1)
            formula = _uniform_relative_formula(values)
            if formula is not None:
                sample = next(i for i, value in enumerate(values) if value == formula)
                helpers.append(PivotHelperFormula(
                    str(sheet.Name), str(pivot.Name), column, formula,
                    str(sheet.Cells(first_row + sample, column).NumberFormat),
                    first_row, last_row,
                ))
    return helpers


def _restore_pivot_helpers(workbook, helpers: list[PivotHelperFormula]) -> None:
    for helper in helpers:
        sheet = _sheet_by_name(workbook, helper.sheet)
        pivot = sheet.PivotTables(helper.pivot)
        area = pivot.TableRange1
        first_row = int(pivot.DataBodyRange.Row)
        last_row = int(area.Row + area.Rows.Count - 1)
        if helper.column < int(area.Column + area.Columns.Count):
            raise RuntimeError(f"Pivot {helper.pivot!r} expanded over its helper column.")
        for row in range(max(1, helper.first_row), helper.last_row + 1):
            if row < first_row or row > last_row:
                cell = sheet.Cells(row, helper.column)
                if cell.FormulaR1C1 == helper.formula:
                    cell.ClearContents()
        cells = sheet.Range(sheet.Cells(first_row, helper.column), sheet.Cells(last_row, helper.column))
        cells.FormulaR1C1 = helper.formula
        cells.NumberFormat = helper.number_format


def _apply_pivot_review(
    workbook, decisions: list[PivotReviewDecision]
) -> set[tuple[str, str]]:
    if not decisions:
        return set()
    sheet = _sheet_by_name(workbook, PIVOT_PERFORMANCE_SHEET_NAME)
    if sheet is None:
        raise RuntimeError("Weekly pivot review requires PIVOT PERFORMANCE.")
    changed: set[tuple[str, str]] = set()
    for decision in decisions:
        try:
            pivot = sheet.PivotTables(decision.pivot)
            field = pivot.PivotFields(decision.field)
            item = field.PivotItems(decision.item)
            if decision.visible is not None:
                item.Visible = decision.visible
            else:
                item.Position = decision.position
        except Exception as exc:
            raise ValueError(
                f"Pivot review cannot set {decision.pivot!r} "
                f"{decision.field} item {decision.item!r}."
            ) from exc
        changed.add((PIVOT_PERFORMANCE_SHEET_NAME, decision.pivot))
    return changed


def _refresh_pivots(
    workbook,
    new_adjustments: set[str],
    *,
    mobile: bool,
    report_week: int | None,
    pivot_review: list[PivotReviewDecision],
    reset_master_leadtime_filter: bool = False,
) -> None:
    refreshed_caches: set[int] = set()
    for sheet in workbook.Worksheets:
        pivot_tables = sheet.PivotTables()
        for pivot_number in range(1, pivot_tables.Count + 1):
            pivot = pivot_tables(pivot_number)
            cache_index = int(pivot.CacheIndex)
            if cache_index not in refreshed_caches:
                cache = pivot.PivotCache()
                cache.Refresh()
                refreshed_caches.add(cache_index)
            pivot.RefreshTable()
    changed_pivots = _include_new_adjustments(
        workbook,
        new_adjustments,
        mobile=mobile,
    )
    if not mobile and reset_master_leadtime_filter:
        # Mobile templates have curated, service-specific value selections.
        changed_pivots.update(_apply_nonnegative_leadtime_filters(workbook))
    if mobile:
        if report_week is None:
            raise ValueError("Mobile PivotTable refresh requires a report week.")
        changed_pivots.update(_apply_mobile_week_window(workbook, report_week))
    changed_pivots.update(_order_week_items(workbook))
    changed_pivots.update(_apply_pivot_review(workbook, pivot_review))
    _refresh_selected_pivots(workbook, changed_pivots)


def _refresh_slide_summary(
    workbook,
    *,
    report_year: int,
    report_week: int,
    mobile: bool,
    reviewed_values: dict[str, str | float],
) -> int:
    sheet = _sheet_by_name(workbook, "SLIDE GH MEETING")
    if sheet is None:
        raise RuntimeError("Template has no 'SLIDE GH MEETING' sheet.")
    pivot = None
    for index in range(1, sheet.PivotTables().Count + 1):
        candidate = sheet.PivotTables(index)
        try:
            if int(candidate.PivotFields("Week Completion").Orientation) == 1:
                pivot = candidate
                break
        except Exception:
            continue
    if pivot is None:
        raise RuntimeError("Slide summary has no Week Completion PivotTable.")

    report_month = reporting_week_start(report_year, report_week).replace(day=1)
    month_field = pivot.PivotFields("Month Completion")
    month_items = month_field.PivotItems()
    original_manual = bool(pivot.ManualUpdate)
    pivot.ManualUpdate = True
    try:
        for index in range(1, month_items.Count + 1):
            item = month_items(index)
            month = completion_month(str(item.Name))
            if month is not None and bool(item.Visible) and bool(item.ShowDetail) != (month == report_month):
                item.ShowDetail = month == report_month
    finally:
        pivot.ManualUpdate = original_manual
    pivot.RefreshTable()

    pivot_range = pivot.TableRange2
    first_row = int(pivot_range.Row)
    last_row = first_row + int(pivot_range.Rows.Count) - 1
    first_column = int(pivot_range.Column)
    last_column = first_column + int(pivot_range.Columns.Count) - 1
    header_row = next(
        (
            row for row in range(first_row, last_row + 1)
            if str(sheet.Cells(row, first_column).Value or "").strip() == "Row Labels"
        ),
        None,
    )
    if header_row is None:
        raise RuntimeError("Slide PivotTable has no Row Labels header.")
    pivot_columns = {
        str(sheet.Cells(header_row, column).Value or "").strip().casefold(): column
        for column in range(first_column + 1, last_column + 1)
    }
    week_rows = {
        str(sheet.Cells(row, first_column).Value or "").strip(): row
        for row in range(header_row + 1, last_row + 1)
    }
    if f"W{report_week}" not in week_rows:
        raise RuntimeError("Slide PivotTable has no current report week.")

    static_columns = {
        column: str(sheet.Cells(9, column).Value or "").strip().casefold()
        for column in range(20, 23)
        if sheet.Cells(9, column).Value not in (None, "")
    }
    old_labels = {}
    old_total_rows = []
    for row in range(10, 101):
        label = str(sheet.Cells(row, 19).Value or "").strip()
        if (month := completion_month(label)) is not None:
            old_labels[month] = label
        if label == "AVG Total":
            old_total_rows.append(row)
    if len(old_total_rows) != 1:
        raise RuntimeError("Slide chart source must contain exactly one AVG Total row.")

    summary_rows = []
    for label, pivot_row in week_rows.items():
        month = completion_month(label)
        week = _week_number(label)
        if month is not None and month < report_month and month.year == report_year:
            summary_rows.append((old_labels.get(month, label), pivot_row, False))
        elif (week is not None and week <= report_week
              and reporting_week_start(report_year, week).replace(day=1) == report_month):
            summary_rows.append((week_label(report_year, week), pivot_row, mobile))
    current_row = 9 + len(summary_rows)
    total_row = current_row + 1
    old_total_row = old_total_rows[0]
    # Keep the template's number formats and total styling while rebuilding values.
    sheet.Range(sheet.Cells(old_total_row, 19), sheet.Cells(old_total_row, 22)).Copy(
        sheet.Cells(total_row, 19)
    )
    sheet.Range(sheet.Cells(10, 19), sheet.Cells(max(old_total_row, total_row), 22)).ClearContents()
    sheet.Range(sheet.Cells(10, 19), sheet.Cells(total_row, 19)).NumberFormat = "@"
    for static_row, (label, pivot_row, rounded) in enumerate(summary_rows, start=10):
        sheet.Cells(static_row, 19).Value = label
        for static_column, service in static_columns.items():
            if service not in pivot_columns:
                raise RuntimeError(f"Slide PivotTable is missing service {service!r}.")
            pivot_value = sheet.Cells(pivot_row, pivot_columns[service]).Value2
            sheet.Cells(static_row, static_column).Value = (
                int(Decimal(str(pivot_value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                if rounded and isinstance(pivot_value, (int, float))
                else pivot_value
            )
    sheet.Cells(total_row, 19).Value = "AVG Total"
    for static_column, service in static_columns.items():
        value = sheet.Cells(last_row, pivot_columns[service]).Value2
        sheet.Cells(total_row, static_column).Value = (
            int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
            if mobile and isinstance(value, (int, float)) else value
        )

    for cell, value in reviewed_values.items():
        sheet.Range(cell).Value = value

    charts = sheet.ChartObjects()
    if int(charts.Count) != 1:
        raise RuntimeError("Slide sheet must contain exactly one summary chart.")
    series = charts(1).Chart.SeriesCollection()
    if int(series.Count) != len(static_columns):
        raise RuntimeError("Slide chart series do not match service columns.")
    for index, static_column in enumerate(static_columns, start=1):
        item = series(index)
        item.XValues = sheet.Range(sheet.Cells(10, 19), sheet.Cells(current_row, 19))
        item.Values = sheet.Range(
            sheet.Cells(10, static_column), sheet.Cells(current_row, static_column)
        )
    if str(sheet.Cells(total_row, 19).Value or "").strip() != "AVG Total":
        raise RuntimeError("Slide summary total row was lost.")
    return current_row


def _hidden_pivot_tail_range(
    is_hidden, is_blank, first_row: int, last_row: int
) -> tuple[int, int] | None:
    if not is_hidden(last_row):
        return None
    start = last_row
    while start > first_row and is_hidden(start - 1):
        start -= 1
    end = last_row
    for row in (last_row + 1, last_row + 2):
        if not is_hidden(row) or not is_blank(row):
            break
        end = row
    return start, end


def _reveal_hidden_pivot_tails(workbook) -> None:
    for sheet in workbook.Worksheets:
        used_range = sheet.UsedRange
        last_column = int(used_range.Column) + int(used_range.Columns.Count) - 1
        for index in range(1, sheet.PivotTables().Count + 1):
            pivot = sheet.PivotTables(index)
            data_range = pivot.TableRange1
            first_row = int(data_range.Row)
            last_row = first_row + int(data_range.Rows.Count) - 1

            def is_blank(row: int) -> bool:
                cells = sheet.Range(sheet.Cells(row, 1), sheet.Cells(row, last_column))
                return int(sheet.Application.WorksheetFunction.CountA(cells)) == 0

            reveal = _hidden_pivot_tail_range(
                lambda row: bool(sheet.Rows(row).Hidden), is_blank, first_row, last_row
            )
            if reveal is not None:
                sheet.Rows(f"{reveal[0]}:{reveal[1]}").Hidden = False


def _fit_column_to_range(sheet, cell_range, *, padding: float = 0, limit: float | None = None) -> None:
    column = sheet.Columns(int(cell_range.Column))
    original_width = float(column.ColumnWidth)
    cell_range.Columns.AutoFit()
    fitted_width = float(column.ColumnWidth) + padding
    if limit is not None:
        fitted_width = min(fitted_width, limit)
    column.ColumnWidth = max(original_width, fitted_width)


def _fit_weekly_summary_columns(workbook) -> None:
    for sheet in workbook.Worksheets:
        if str(sheet.Name) == MASTER_SHEET_NAME:
            continue
        for index in range(1, sheet.PivotTables().Count + 1):
            pivot = sheet.PivotTables(index)
            body = pivot.DataBodyRange
            if body is not None:
                for column_index in range(1, int(body.Columns.Count) + 1):
                    _fit_column_to_range(sheet, body.Columns(column_index))

            page_range = pivot.TableRange2
            data_range = pivot.TableRange1
            first_row = int(page_range.Row)
            header_row = int(data_range.Row)
            first_column = int(page_range.Column)
            is_service_summary = (
                str(sheet.Name) == PIVOT_PERFORMANCE_SHEET_NAME
                and first_row <= 5
                and first_column >= 7
            )
            is_monitoring_summary = (
                str(sheet.Name) == "MONITORING ORDER - MOBILE"
                and first_row <= 5
                and first_column == 6
            )
            if header_row <= first_row or not (is_service_summary or is_monitoring_summary):
                continue
            for offset, padding in ((0, 3.0), (1, 5.5)):
                column_number = first_column + offset
                page_column = sheet.Range(
                    sheet.Cells(first_row, column_number),
                    sheet.Cells(header_row, column_number),
                )
                _fit_column_to_range(sheet, page_column, padding=padding, limit=32)

    slide = _sheet_by_name(workbook, "SLIDE GH MEETING")
    if slide is None:
        return
    for index in range(1, slide.PivotTables().Count + 1):
        pivot = slide.PivotTables(index)
        pivot_range = pivot.TableRange2
        if int(pivot_range.Row) > 5 or int(pivot_range.Column) != 1:
            continue
        data_range = pivot.TableRange1
        first_data_row = int(data_range.Row) + 1
        last_data_row = int(data_range.Row) + int(data_range.Rows.Count) - 1
        helper_column = int(pivot_range.Column) + int(pivot_range.Columns.Count)
        if first_data_row <= last_data_row:
            helper_values = slide.Range(
                slide.Cells(first_data_row, helper_column),
                slide.Cells(last_data_row, helper_column),
            )
            _fit_column_to_range(slide, helper_values)
        break


def _validate_workbook(
    workbook,
    records: list[dict[str, Any]],
    expected_pivots: dict[str, set[str]],
) -> None:
    data_sheet = _sheet_by_name(workbook, MASTER_SHEET_NAME)
    if data_sheet is None:
        raise RuntimeError(f"Generated workbook has no {MASTER_SHEET_NAME!r} sheet.")
    table = data_sheet.ListObjects(MASTER_TABLE_NAME)
    actual_rows = int(table.ListRows.Count)
    actual_columns = int(table.ListColumns.Count)
    if actual_rows != len(records) or actual_columns != len(MASTER_COLUMNS):
        raise RuntimeError(
            "Generated Table1 shape mismatch: "
            f"expected {len(records)}x{len(MASTER_COLUMNS)}, "
            f"got {actual_rows}x{actual_columns}."
        )
    actual_pivots = _pivot_manifest(workbook)
    if actual_pivots != expected_pivots:
        raise RuntimeError("Generated PivotTable names differ from the template contract.")
    for sheet in workbook.Worksheets:
        for pivot_name in actual_pivots[str(sheet.Name)]:
            pivot = sheet.PivotTables(pivot_name)
            if str(pivot.PivotCache().SourceData) != MASTER_TABLE_NAME:
                raise RuntimeError(
                    f"PivotTable {sheet.Name!r}/{pivot_name!r} no longer reads Table1."
                )
            if int(pivot.TableRange2.Rows.Count) < 1:
                raise RuntimeError(
                    f"PivotTable {sheet.Name!r}/{pivot_name!r} has no rendered rows."
                )
            data_range = pivot.TableRange1
            last_row = int(data_range.Row) + int(data_range.Rows.Count) - 1
            if bool(sheet.Rows(last_row).Hidden):
                raise RuntimeError(
                    f"PivotTable {sheet.Name!r}/{pivot_name!r} ends on hidden row {last_row}."
                )
    sample_rows = ((2, records[0]), (len(records) + 1, records[-1])) if records else ()
    for row_number, record in sample_rows:
        for column_name in ("Month of FAB Upload", "YEAR FAB Upload", "quote_num"):
            column_number = MASTER_COLUMNS.index(column_name) + 1
            actual_value = data_sheet.Cells(row_number, column_number).Value2
            if actual_value != record.get(column_name):
                raise RuntimeError(
                    f"Generated {column_name!r} changed type or value at row {row_number}."
                )
        for column_name in ("fab_upload", "cash_dt", "order_completion_date"):
            column_number = MASTER_COLUMNS.index(column_name) + 1
            actual_value = data_sheet.Cells(row_number, column_number).Value2
            expected_value = _com_value(column_name, record.get(column_name))
            if expected_value is None:
                if actual_value is not None:
                    raise RuntimeError(
                        f"Generated {column_name!r} should be blank at row {row_number}."
                    )
            elif not isinstance(actual_value, (int, float)) or abs(
                actual_value - expected_value
            ) > 1e-7:
                raise RuntimeError(
                    f"Generated {column_name!r} changed date at row {row_number}."
                )


def write_weekly_workbook(
    *,
    data_csv: str | Path,
    template_workbook: str | Path,
    output_workbook: str | Path,
    mobile: bool = False,
    refresh_pivots: bool = True,
    report_week: int | None = None,
    report_date: date | None = None,
    slide_review: str | Path | None = None,
    pivot_review: str | Path | None = None,
) -> Path:
    if sys.platform != "win32":
        raise RuntimeError("Weekly workbook generation requires Windows and Microsoft Excel.")
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for weekly workbook generation.") from exc

    data_path = Path(data_csv).resolve()
    template_path = Path(template_workbook).resolve()
    output_path = Path(output_workbook).resolve()
    if not data_path.is_file():
        raise FileNotFoundError(f"Weekly data CSV not found: {data_path}")
    if not template_path.is_file():
        raise FileNotFoundError(f"Weekly template workbook not found: {template_path}")
    if output_path == template_path:
        raise ValueError("Output workbook must not overwrite the template workbook.")

    records = list(iter_master_csv(data_path))
    resolved_report_week = report_week
    if mobile and resolved_report_week is None:
        resolved_report_week = _latest_report_week(records)
    if resolved_report_week is not None and not 1 <= resolved_report_week <= 53:
        raise ValueError("Report week must be between 1 and 53.")
    reviewed_slide_values = load_slide_review(slide_review)["mobile" if mobile else "master"]
    reviewed_pivot_values = load_pivot_review(pivot_review)["mobile" if mobile else "master"]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    working_dir = output_path.parent / f".{output_path.stem}.{uuid.uuid4().hex}.work"
    working_path = working_dir / output_path.name
    working_dir.mkdir(parents=True, exist_ok=False)
    shutil.copy2(template_path, working_path)

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        with excel_process_lock(EXCEL_LOCK_TIMEOUT_SECONDS):
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            try:
                excel.Calculation = XL_CALCULATION_MANUAL
            except Exception:
                pass
            workbook = excel.Workbooks.Open(
                str(working_path),
                UpdateLinks=0,
                ReadOnly=False,
            )
            data_sheet = _sheet_by_name(workbook, MASTER_SHEET_NAME)
            if data_sheet is None:
                raise RuntimeError(f"Template has no {MASTER_SHEET_NAME!r} sheet.")
            expected_pivots = _pivot_manifest(workbook)
            if mobile:
                expected_pivots[PIVOT_PERFORMANCE_SHEET_NAME] -= MOBILE_ONLY_REMOVED_PIVOTS
            new_adjustments = _write_table(data_sheet, records)
            if mobile:
                _remove_mobile_only_pivots(workbook)
            pivot_helpers = _capture_pivot_helpers(workbook) if refresh_pivots else []
            if refresh_pivots:
                report_year = (
                    (report_date or max(
                        value.date() if isinstance(value, datetime) else value
                        for record in records
                        if isinstance((value := record.get("fab_upload")), (date, datetime))
                    )).year if resolved_report_week is not None else None
                )
                reset_master_filter = (
                    not mobile and resolved_report_week is not None
                    and _master_month_rollover(workbook, report_year, resolved_report_week)
                )
                _refresh_pivots(
                    workbook,
                    new_adjustments,
                    mobile=mobile,
                    report_week=resolved_report_week,
                    pivot_review=reviewed_pivot_values,
                    reset_master_leadtime_filter=reset_master_filter,
                )
                _restore_pivot_helpers(workbook, pivot_helpers)
                if resolved_report_week is not None:
                    _refresh_slide_summary(
                        workbook,
                        report_year=report_year,
                        report_week=resolved_report_week,
                        mobile=mobile,
                        reviewed_values=reviewed_slide_values,
                    )
            try:
                excel.Calculation = XL_CALCULATION_AUTOMATIC
            except Exception:
                pass
            excel.CalculateFullRebuild()
            _reveal_hidden_pivot_tails(workbook)
            _fit_weekly_summary_columns(workbook)
            _validate_workbook(workbook, records, expected_pivots)
            workbook.Save()
            workbook.Close(SaveChanges=False)
            workbook = None
            excel.Quit()
            excel = None

        wait_for_xlsx(working_path, timeout_seconds=60)
        working_path.replace(output_path)
        wait_for_xlsx(output_path, timeout_seconds=60)
        return output_path
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        try:
            working_path.unlink(missing_ok=True)
        except OSError:
            pass
        try:
            working_dir.rmdir()
        except OSError:
            pass


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-csv", required=True, type=Path)
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--mobile", action="store_true")
    parser.add_argument("--report-week", type=int)
    parser.add_argument("--report-date", type=date.fromisoformat)
    parser.add_argument("--slide-review", type=Path)
    parser.add_argument("--pivot-review", type=Path)
    parser.add_argument("--skip-pivot-refresh", action="store_true")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    started = time.perf_counter()
    output = write_weekly_workbook(
        data_csv=args.data_csv,
        template_workbook=args.template,
        output_workbook=args.output,
        mobile=args.mobile,
        refresh_pivots=not args.skip_pivot_refresh,
        report_week=args.report_week,
        report_date=args.report_date,
        slide_review=args.slide_review,
        pivot_review=args.pivot_review,
    )
    print(f"Generated {output} in {time.perf_counter() - started:.1f}s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
