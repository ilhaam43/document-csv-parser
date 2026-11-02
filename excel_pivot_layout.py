"""Shared Excel COM helpers for dynamic PivotTable section layouts."""

from __future__ import annotations

import re
from collections.abc import Iterable
from datetime import date
from openpyxl.utils import get_column_letter

from completion_threshold_rules import (
    CompletionThresholds,
    STANDARD_PROFILE,
    TARGET_AFTER_MODE,
    WEEKLY_MODE,
    completion_legend_values_for,
    completion_thresholds_for,
)


EXCEL_MAX_ROW = 1_048_576
PERCENTAGE_COMPLETION_HEADER = "Percentage of Completion (SO Complete, Cancel & Change Target)"
COMPLETION_SECTION_MARKERS = (
    "target complete",
    "new registration",
    "non new registration",
    "delay completion",
    "target after",
    "all target",
    "target not inputted",
    "target order on going process",
    "all target order",
)


def normalize_layout_text(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def pivot_row_identity(cell) -> tuple[tuple[str, str], ...]:
    """Identify a pivot row by categories and subtotal depth, not its position."""
    items = cell.PivotCell.RowItems
    return tuple(sorted(
        (normalize_layout_text(items(i).Parent.SourceName), str(items(i).Name))
        for i in range(1, int(items.Count) + 1)
    ))


def completion_denominator_references(pivot_table) -> dict[tuple, str]:
    body = pivot_table.DataBodyRange
    if body is None:
        return {}
    count_column = None
    for column in range(1, int(body.Columns.Count) + 1):
        field = body.Cells(1, column).PivotCell.DataField
        if normalize_layout_text(field.SourceName) == "quo" and int(field.Function) == -4112:
            count_column = column
            break
    if count_column is None:
        raise RuntimeError(f"PivotTable '{pivot_table.Name}' has no quote-count data column.")
    references = {}
    for row in range(1, int(body.Rows.Count) + 1):
        cell = body.Cells(row, count_column)
        identity = pivot_row_identity(cell)
        if identity in references:
            raise RuntimeError(f"PivotTable '{pivot_table.Name}' has duplicate summary row {identity!r}.")
        references[identity] = f"{get_column_letter(int(cell.Column))}{int(cell.Row)}"
    return references


def completion_percentage_formula(numerator_cells: list[str], denominator_cell: str) -> str:
    numerator = "+".join(numerator_cells) or "0"
    return f'=IF({denominator_cell}=0,0,({numerator})/{denominator_cell})'


def restore_dashboard_pivot_presentation(pivot_sheet, column_widths: dict[int, float]) -> None:
    for pivot in pivot_sheet.PivotTables():
        pivot.PreserveFormatting = True
        pivot.HasAutoFormat = False
        if pivot.RowRange is not None:
            pivot.RowRange.EntireRow.Hidden = False
    for column, width in column_widths.items():
        pivot_sheet.Columns(column).ColumnWidth = width


def completion_summary_pivot(pivot_sheet, phase_pivot):
    expected_fields = {normalize_layout_text(field.SourceName) for field in phase_pivot.RowFields}
    candidates = []
    for candidate in pivot_sheet.PivotTables():
        if int(candidate.TableRange1.Column) >= int(phase_pivot.TableRange1.Column):
            continue
        if {normalize_layout_text(field.SourceName) for field in candidate.RowFields} != expected_fields:
            continue
        if not any(normalize_layout_text(f.SourceName) == "quo" and int(f.Function) == -4112 for f in candidate.DataFields):
            continue
        candidates.append(candidate)
    if not candidates:
        raise RuntimeError(f"PivotTable '{phase_pivot.Name}' has no adjacent quote-count summary.")
    return min(candidates, key=lambda candidate: abs(int(candidate.RowRange.Row) - int(phase_pivot.RowRange.Row)))


def completion_legend_values(
    report_date: date,
    mode: str = WEEKLY_MODE,
    profile: str = STANDARD_PROFILE,
) -> tuple[str, str, str]:
    return completion_legend_values_for(
        report_date,
        profile=profile,
        mode=mode,
    )


def completion_section_modes(
    value_rows,
    first_row: int,
) -> list[tuple[int, str]]:
    """Return dynamic normal/Target After section modes from visible labels."""
    normalized_markers = tuple(
        normalize_layout_text(marker) for marker in COMPLETION_SECTION_MARKERS
    )
    modes_by_row: dict[int, str] = {}
    for row_index, row_values in enumerate(value_rows, start=first_row):
        if not isinstance(row_values, tuple):
            row_values = (row_values,)
        for value in row_values:
            normalized = normalize_layout_text(value)
            if not normalized or not any(
                marker in normalized for marker in normalized_markers
            ):
                continue
            mode = (
                TARGET_AFTER_MODE
                if "targetafter" in normalized
                else WEEKLY_MODE
            )
            if mode == TARGET_AFTER_MODE or row_index not in modes_by_row:
                modes_by_row[row_index] = mode
    return sorted(modes_by_row.items())


def completion_mode_for_row(
    row_number: int,
    section_modes: Iterable[tuple[int, str]],
    fallback: str = WEEKLY_MODE,
) -> str:
    if fallback == TARGET_AFTER_MODE:
        return TARGET_AFTER_MODE
    preceding_modes = [
        (section_row, mode)
        for section_row, mode in section_modes
        if section_row <= row_number
    ]
    if not preceding_modes:
        return fallback
    return max(preceding_modes, key=lambda item: item[0])[1]


def apply_completion_icon_thresholds(
    target_range,
    thresholds: CompletionThresholds,
    *,
    create_if_missing: bool = False,
) -> int:
    """Update every traffic-light icon rule applied to a percentage range."""
    updated_rules = 0
    format_conditions = target_range.FormatConditions
    for condition_index in range(1, int(format_conditions.Count) + 1):
        condition = format_conditions(condition_index)
        try:
            if int(condition.Type) != 6 or int(condition.IconSet.ID) != 4:
                continue
        except Exception:
            continue
        _set_completion_icon_criteria(condition, thresholds)
        updated_rules += 1
    if updated_rules == 0 and create_if_missing:
        # A moved or empty PivotTable can lose copied conditional formatting.
        # Recreate only this block's icon rule; retain other conditional rules.
        condition = format_conditions.AddIconSetCondition()
        condition.IconSet = target_range.Worksheet.Parent.IconSets(4)
        condition.ReverseOrder = False
        condition.ShowIconOnly = False
        _set_completion_icon_criteria(condition, thresholds)
        updated_rules = 1
    return updated_rules


def _set_completion_icon_criteria(condition, thresholds: CompletionThresholds) -> None:
    criteria = condition.IconCriteria
    criteria(2).Type = 0  # xlConditionValueNumber
    criteria(2).Value = thresholds.yellow_percent / 100
    criteria(2).Operator = 7  # xlGreaterEqual
    criteria(3).Type = 0  # xlConditionValueNumber
    criteria(3).Value = thresholds.green_percent / 100
    criteria(3).Operator = 5  # xlGreater


def completion_percentage_headers(pivot_sheet) -> list[tuple[int, int]]:
    """Find completion columns in source order, independent of template row shifts."""
    used = pivot_sheet.UsedRange
    values = used.Value2
    if values is None:
        return []
    if not isinstance(values, tuple):
        values = ((values,),)
    elif values and not isinstance(values[0], tuple):
        values = (values,)
    return [
        (row, column)
        for row, cells in enumerate(values, start=int(used.Row))
        for column, value in enumerate(cells, start=int(used.Column))
        if normalize_layout_text(value).startswith("percentageofcompletion")
    ]


def percentage_format_row_spans(header_row: int, end_row: int) -> list[tuple[str, int, int]]:
    """Map table roles without formatting a header or blank row as detail."""
    spans = [("header", header_row, header_row)]
    if end_row <= header_row:
        return spans
    first, last = header_row + 1, end_row - 1
    if first <= last:
        spans.append(("first_detail", first, first))
        if first + 1 <= last - 1:
            spans.append(("middle_detail", first + 1, last - 1))
        if last > first:
            spans.append(("last_detail", last, last))
    spans.append(("grand_total", end_row, end_row))
    return spans


def percentage_block_end_row(
    pivot_sheet,
    header_row: int,
    header_col: int,
    maximum_row: int,
) -> int:
    end_row = header_row
    for row_number in range(header_row + 1, maximum_row + 1):
        cell = pivot_sheet.Cells(row_number, header_col)
        formula = str(cell.Formula or "")
        value = cell.Value2
        if formula.startswith("=") or value not in (None, ""):
            end_row = row_number
            continue
        break
    return end_row


def align_completion_threshold_legends(
    pivot_sheet,
    report_date: date,
    layout_specs: Iterable[tuple[int, int, str]],
    profile: str = STANDARD_PROFILE,
) -> None:
    """Align legends and icon rules with the active weekly business thresholds."""
    try:
        used_range = pivot_sheet.UsedRange
        used_values = used_range.Value2
    except Exception:
        return
    if not used_values:
        return

    if not isinstance(used_values, tuple):
        value_rows = ((used_values,),)
    elif used_values and not isinstance(used_values[0], tuple):
        value_rows = (used_values,)
    else:
        value_rows = used_values

    percentage_headers: list[tuple[int, int]] = []
    legend_starts: list[tuple[int, int]] = []
    percentage_key = normalize_layout_text(PERCENTAGE_COMPLETION_HEADER)
    for row_index, row_values in enumerate(value_rows, start=used_range.Row):
        if not isinstance(row_values, tuple):
            row_values = (row_values,)
        for column_index, value in enumerate(row_values, start=used_range.Column):
            normalized = normalize_layout_text(value)
            if normalized == percentage_key or normalized.startswith("percentageof"):
                percentage_headers.append((row_index, column_index))
            elif str(value or "").strip().lower().startswith("green :"):
                legend_starts.append((row_index, column_index))

    percentage_headers.sort()
    legend_starts.sort()
    section_modes = completion_section_modes(value_rows, int(used_range.Row))
    maximum_row = int(used_range.Row + used_range.Rows.Count - 1)
    specs = list(layout_specs)
    for index, ((header_row, header_col), (minimum_col, row_offset, mode)) in enumerate(
        zip(percentage_headers, specs)
    ):
        resolved_mode = completion_mode_for_row(
            header_row,
            section_modes,
            mode,
        )
        thresholds = completion_thresholds_for(
            report_date,
            profile=profile,
            mode=resolved_mode,
        )
        target_row = header_row + row_offset
        target_col = max(minimum_col, header_col + 2)
        target_range = pivot_sheet.Range(
            pivot_sheet.Cells(target_row, target_col),
            pivot_sheet.Cells(target_row + 2, target_col),
        )

        if index < len(legend_starts):
            source_row, source_col = legend_starts[index]
            source_range = pivot_sheet.Range(
                pivot_sheet.Cells(source_row, source_col),
                pivot_sheet.Cells(source_row + 2, source_col),
            )
            if (source_row, source_col) != (target_row, target_col):
                try:
                    source_range.Copy()
                    target_range.PasteSpecial(Paste=-4104)  # xlPasteAll
                    pivot_sheet.Application.CutCopyMode = False
                    source_range.Clear()
                except Exception:
                    pass

        target_range.Value = tuple(
            (value,)
            for value in completion_legend_values(
                report_date,
                resolved_mode,
                profile,
            )
        )

        end_row = percentage_block_end_row(
            pivot_sheet,
            header_row,
            header_col,
            maximum_row,
        )
        if end_row <= header_row:
            continue
        percentage_range = pivot_sheet.Range(
            pivot_sheet.Cells(header_row + 1, header_col),
            pivot_sheet.Cells(end_row, header_col),
        )
        if apply_completion_icon_thresholds(percentage_range, thresholds, create_if_missing=True) == 0:
            raise RuntimeError(
                "No traffic-light icon rule found for percentage block "
                f"{pivot_sheet.Cells(header_row, header_col).Address}."
            )

    for row_index, column_index in legend_starts[len(specs) :]:
        try:
            pivot_sheet.Range(
                pivot_sheet.Cells(row_index, column_index),
                pivot_sheet.Cells(row_index + 2, column_index),
            ).Clear()
        except Exception:
            pass


def section_banner_end_columns(
    titles: list[tuple[int, int]],
    pivot_ranges: list[tuple[int, int, int]],
    side_headers: list[tuple[int, int]],
) -> dict[tuple[int, int], int]:
    """Return each title's right edge from the tables in its row/column section."""
    boundaries: dict[tuple[int, int], int] = {}
    distinct_title_rows = sorted({row for row, _ in titles})

    for title_row, title_col in sorted(titles):
        next_title_row = next(
            (row for row in distinct_title_rows if row > title_row),
            EXCEL_MAX_ROW + 1,
        )
        next_title_col = min(
            (
                column
                for row, column in titles
                if row == title_row and column > title_col
            ),
            default=None,
        )

        def belongs_to_horizontal_section(column: int) -> bool:
            return column >= title_col and (
                next_title_col is None or column < next_title_col
            )

        candidate_columns = [
            end_col
            for pivot_row, start_col, end_col in pivot_ranges
            if title_row < pivot_row < next_title_row
            and belongs_to_horizontal_section(start_col)
        ]
        candidate_columns.extend(
            column
            for header_row, column in side_headers
            if title_row < header_row < next_title_row
            and belongs_to_horizontal_section(column)
        )
        if candidate_columns:
            boundaries[(title_row, title_col)] = max(candidate_columns)

    return boundaries


def update_dynamic_pivot_section_titles(
    pivot_sheet,
    replacements: dict[str, str],
) -> int:
    """Replace month-bearing section titles without relying on fixed rows."""
    normalized_replacements = {
        normalize_layout_text(marker): value
        for marker, value in replacements.items()
    }
    try:
        used_range = pivot_sheet.UsedRange
        used_values = used_range.Value
    except Exception:
        return 0
    if not used_values:
        return 0

    if not isinstance(used_values, tuple):
        value_rows = ((used_values,),)
    elif used_values and not isinstance(used_values[0], tuple):
        value_rows = (used_values,)
    else:
        value_rows = used_values

    updated = 0
    for row_offset, row_values in enumerate(value_rows, start=used_range.Row):
        if not isinstance(row_values, tuple):
            row_values = (row_values,)
        for column_offset, raw_value in enumerate(row_values, start=used_range.Column):
            normalized_value = normalize_layout_text(raw_value)
            if not normalized_value:
                continue
            replacement = next(
                (
                    value
                    for marker, value in normalized_replacements.items()
                    if marker in normalized_value
                ),
                None,
            )
            if replacement is None:
                continue
            try:
                pivot_sheet.Cells(row_offset, column_offset).Value = replacement
                updated += 1
            except Exception:
                pass
    return updated


def resize_dynamic_pivot_section_banners(
    pivot_sheet,
    title_markers: Iterable[str],
    side_header_markers: Iterable[str] = (),
) -> None:
    """Resize merged Pivot section titles to the current table width."""
    normalized_title_markers = tuple(normalize_layout_text(value) for value in title_markers)
    normalized_side_markers = tuple(normalize_layout_text(value) for value in side_header_markers)

    try:
        used_range = pivot_sheet.UsedRange
        used_values = used_range.Value
    except Exception:
        return
    if not used_values:
        return

    if not isinstance(used_values, tuple):
        value_rows = ((used_values,),)
    elif used_values and not isinstance(used_values[0], tuple):
        value_rows = (used_values,)
    else:
        value_rows = used_values

    titles: list[tuple[int, int]] = []
    side_headers: list[tuple[int, int]] = []
    for row_offset, row_values in enumerate(value_rows, start=used_range.Row):
        if not isinstance(row_values, tuple):
            row_values = (row_values,)
        for column_offset, raw_value in enumerate(row_values, start=used_range.Column):
            normalized_value = normalize_layout_text(raw_value)
            if normalized_value and any(marker in normalized_value for marker in normalized_title_markers):
                titles.append((row_offset, column_offset))
            if normalized_value and any(marker in normalized_value for marker in normalized_side_markers):
                side_headers.append((row_offset, column_offset))

    if not titles:
        return

    pivot_ranges: list[tuple[int, int, int]] = []
    try:
        pivot_tables = pivot_sheet.PivotTables()
        for pivot_index in range(1, pivot_tables.Count + 1):
            table_range = pivot_tables(pivot_index).TableRange2
            pivot_ranges.append(
                (
                    int(table_range.Row),
                    int(table_range.Column),
                    int(table_range.Column + table_range.Columns.Count - 1),
                )
            )
    except Exception:
        return

    desired_boundaries = section_banner_end_columns(titles, pivot_ranges, side_headers)
    for (title_row, title_col), desired_end_col in desired_boundaries.items():
        try:
            title_cell = pivot_sheet.Cells(title_row, title_col)
            merge_area = title_cell.MergeArea if bool(title_cell.MergeCells) else title_cell
            merge_end_row = int(merge_area.Row + merge_area.Rows.Count - 1)
            old_end_col = int(merge_area.Column + merge_area.Columns.Count - 1)
            if old_end_col == desired_end_col:
                continue

            title_text = title_cell.Value
            fill_pattern = merge_area.Interior.Pattern
            fill_color = merge_area.Interior.Color
            border_styles = {}
            for border_index in (7, 8, 9, 10):  # left, top, bottom, right
                border = merge_area.Borders(border_index)
                border_styles[border_index] = (
                    border.LineStyle,
                    border.Weight,
                    border.Color,
                )

            if bool(title_cell.MergeCells):
                merge_area.UnMerge()

            if old_end_col > desired_end_col:
                trailing_range = pivot_sheet.Range(
                    pivot_sheet.Cells(title_row, desired_end_col + 1),
                    pivot_sheet.Cells(merge_end_row, old_end_col),
                )
                trailing_range.ClearContents()
                trailing_range.Interior.Pattern = -4142  # xlPatternNone
                trailing_range.Borders.LineStyle = -4142  # xlLineStyleNone

            new_merge_area = pivot_sheet.Range(
                pivot_sheet.Cells(title_row, title_col),
                pivot_sheet.Cells(merge_end_row, desired_end_col),
            )
            new_merge_area.Merge()
            new_merge_area.Value = title_text
            new_merge_area.Interior.Pattern = fill_pattern
            new_merge_area.Interior.Color = fill_color
            for border_index, (line_style, weight, color) in border_styles.items():
                border = new_merge_area.Borders(border_index)
                border.LineStyle = line_style
                if line_style != -4142:
                    border.Weight = weight
                    border.Color = color
        except Exception:
            continue
