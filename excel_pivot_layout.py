"""Shared Excel COM helpers for dynamic PivotTable section layouts."""

from __future__ import annotations

import re
from collections.abc import Iterable


EXCEL_MAX_ROW = 1_048_576


def normalize_layout_text(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


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
