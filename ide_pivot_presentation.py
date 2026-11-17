"""Preserve native IDE pivot presentation across a schema/cache replacement."""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field, replace
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from excel_pivot_layout import normalize_layout_text
from ide_workbook_package import S, worksheet_part


@dataclass(frozen=True)
class CellPresentation:
    font_name: str
    font_size: float
    number_format: str
    alignment: tuple


@dataclass
class PivotPresentation:
    headers: dict = field(default_factory=dict)
    populated_headers: set = field(default_factory=set)
    header_rows: dict = field(default_factory=dict)
    cells: dict = field(default_factory=dict)
    defaults: dict = field(default_factory=lambda: defaultdict(Counter))
    heights: dict = field(default_factory=dict)
    height_defaults: dict = field(default_factory=lambda: defaultdict(Counter))
    row_fields: tuple = ()


@dataclass
class SheetPresentation:
    pivots: dict
    rows: dict
    peers: dict
    peer_heights: dict


def read_presentation_grid(path: Path) -> tuple[dict, dict]:
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book["PIVOT"]
        grid = {}
        for row in sheet:
            for cell in row:
                if not getattr(cell, "has_style", False) and cell.value is None:
                    continue
                a = cell.alignment
                grid[cell.row, cell.column] = CellPresentation(
                    cell.font.name or "Calibri", float(cell.font.sz or 11), cell.number_format,
                    (a.horizontal, a.vertical, bool(a.wrapText), bool(a.shrinkToFit), a.indent, a.textRotation),
                )
        with ZipFile(path) as archive:
            root = ET.fromstring(archive.read(worksheet_part(archive, "PIVOT")))
        sheet_format = root.find(f"{S}sheetFormatPr")
        default_height = float(sheet_format.get("defaultRowHeight", "15")) if sheet_format is not None else 15
        rows = {r: (default_height, False) for r in range(1, sheet.max_row + 1)}
        for row in root.findall(f"{S}sheetData/{S}row"):
            index = int(row.get("r"))
            rows[index] = (float(row.get("ht", default_height)) or default_height,
                           row.get("hidden", "0") in {"1", "true"})
        return grid, rows
    finally:
        book.close()


def row_role(identity: tuple, fields: tuple) -> str:
    names = {name for name, _ in identity}
    return next((name for name in reversed(fields) if name in names), "__total__")


def measure_keys(pivot, body) -> list[tuple]:
    fields = list(pivot.DataFields)
    keys = []
    for column in range(1, int(body.Columns.Count) + 1):
        key = None
        for row in range(1, int(body.Rows.Count) + 1):
            cell = body.Cells(row, column)
            try:
                item = cell.PivotCell
                measure = item.DataField
                categories = item.ColumnItems
                key = (str(measure.SourceName).casefold(), int(measure.Function), int(measure.Calculation),
                       tuple((str(categories(i).Parent.SourceName).casefold(), str(categories(i).Name) or "(blank)")
                             for i in range(1, int(categories.Count) + 1)))
                break
            except Exception:
                if cell.Value2 not in (None, ""):
                    raise
        if key is None:
            measure = fields[(column - 1) % len(fields)]
            key = (str(measure.SourceName).casefold(), int(measure.Function), int(measure.Calculation), ())
        keys.append(key)
    return keys


def _body_layout(pivot):
    region = pivot.TableRange1
    body = pivot.DataBodyRange
    if body is None:
        bottom = int(region.Row + region.Rows.Count - 1)
        return bottom, bottom, int(region.Column + region.Columns.Count - 1), []
    return (int(body.Row), int(body.Row + body.Rows.Count - 1), int(body.Column), measure_keys(pivot, body))


def pivot_row_identity(cell, *, field_names: dict | None = None) -> tuple:
    items = cell.PivotCell.RowItems
    count = int(items.Count)
    sources = field_names.get(count) if field_names is not None else None
    names, identity = [], []
    for index in range(1, count + 1):
        item = items(index)
        source = sources[index - 1] if sources is not None else normalize_layout_text(item.Parent.SourceName)
        names.append(source)
        identity.append((source, str(item.Name)))
    if field_names is not None and sources is None:
        # Within a single pivot pass, RowItems at a given depth share the same
        # row-field hierarchy. Never share this cache across pivots or refreshes.
        field_names[count] = tuple(names)
    return tuple(sorted(identity))


def _row_identity(sheet, pivot, row: int, data_column: int, *, field_names=None,
                   data_columns=None, data_rows=None, cell_at=None) -> tuple:
    if data_columns is None or data_rows is None:
        body = pivot.DataBodyRange
        if body is None:
            return ()
        data_columns, data_rows = int(body.Columns.Count), int(body.Rows.Count)
    if cell_at is None:
        cell_at = sheet.Cells
    error = None
    values = []
    for offset in range(data_columns):
        cell = cell_at(row, data_column+offset)
        try:
            return (pivot_row_identity(cell, field_names=field_names) if field_names is not None
                    else pivot_row_identity(cell))
        except Exception as exc:
            values.append(cell.Value2)
            error = exc
    if data_rows == 1 and all(value in (None, "") for value in values):
        return ()  # Excel exposes no RowItems for an empty grand-total cell.
    raise RuntimeError(f"Cannot identify presentation row {row} in pivot {pivot.Name}.") from error


def capture_presentation(sheet, grid: dict, rows: dict) -> SheetPresentation:
    models = {}
    peers, peer_heights = defaultdict(Counter), defaultdict(Counter)
    cell_at = sheet.Cells
    for pivot in sheet.PivotTables():
        region, outer = pivot.TableRange1, pivot.TableRange2
        start, end, data_column, keys = _body_layout(pivot)
        fields = tuple(str(f.SourceName).casefold() for f in pivot.RowFields)
        kind = "summary" if len(list(pivot.DataFields)) > 1 else "phase"
        model = PivotPresentation(row_fields=fields)
        field_names = {}
        for r in range(int(outer.Row), start):
            model.header_rows[r-int(outer.Row)] = rows.get(r, (19.5, False))
            for c in range(int(outer.Column), int(outer.Column+outer.Columns.Count)):
                if (r, c) in grid:
                    model.headers[r-int(outer.Row), c-int(outer.Column)] = grid[r, c]
                    if sheet.Cells(r, c).Value2 not in (None, ""):
                        model.populated_headers.add((r-int(outer.Row), c-int(outer.Column)))
        for r in range(start, end+1):
            identity = _row_identity(sheet, pivot, r, data_column, field_names=field_names,
                                     data_columns=len(keys), data_rows=end-start+1, cell_at=cell_at) if keys else ()
            role = row_role(identity, fields)
            height = rows.get(r, (19.5, False))[0]
            model.heights[identity] = height
            model.height_defaults[role][height] += 1
            peer_heights[kind, role][height] += 1
            for c in range(int(region.Column), int(region.Column+region.Columns.Count)):
                style = grid.get((r, c))
                if style is None:
                    continue
                key = ("label", c-int(region.Column)) if c < data_column else ("measure", keys[c-data_column] if keys else ())
                model.cells[identity, key] = style
                fallback = (role, key) if key[0] == "label" else key[:1]+key[1][:3]
                model.defaults[fallback][style] += 1
                peers[kind, fallback][style] += 1
        models[str(pivot.Name)] = model
    return SheetPresentation(models, rows, peers, peer_heights)


def select_cell_style(model: PivotPresentation, peers: dict, kind: str, identity: tuple, key: tuple):
    exact = model.cells.get((identity, key))
    if exact is not None:
        return exact
    role = row_role(identity, model.row_fields)
    if key[0] == "label" and role != "__total__" and model.cells and all(not row for row, _ in model.cells):
        total = model.cells.get(((), key))
        if total is not None:
            alignment = list(total.alignment)
            alignment[4] = model.row_fields.index(role)
            return replace(total, alignment=tuple(alignment))
    fallback = (role, key) if key[0] == "label" else key[:1]+key[1][:3]
    counter = model.defaults.get(fallback) or peers.get((kind, fallback))
    if counter:
        return counter.most_common(1)[0][0]
    # An empty reference pivot has only a styled total row.
    candidates = [counter for candidate, counter in model.defaults.items()
                  if candidate[0] == "__total__" and key[0] == "label"]
    return candidates[0].most_common(1)[0][0] if candidates else None


def display_style(style: CellPresentation | None, kind: str, key: tuple):
    if style is None or kind != "summary" or key[0] != "measure":
        return style
    alignment = list(style.alignment)
    alignment[:2] = ["center", "center"]
    return replace(style, alignment=tuple(alignment))


def header_style(style: CellPresentation | None, *, is_new_measure_heading: bool):
    if style is None or not is_new_measure_heading:
        return style
    alignment = list(style.alignment)
    alignment[:2] = ["center", "center"]
    return replace(style, alignment=tuple(alignment))


def validate_centered_header(cell) -> None:
    if cell.value is not None and (
        cell.alignment.horizontal != "center" or cell.alignment.vertical != "center"
    ):
        raise RuntimeError(f"Column heading {cell.coordinate} is not centered: {cell.value!r}.")


def _apply(target, style: CellPresentation) -> None:
    from ide_list_format import apply_column_alignment

    font = target.Font
    if font.Name != style.font_name:
        font.Name = style.font_name
    if font.Size != style.font_size:
        font.Size = style.font_size
    if target.NumberFormat != style.number_format:
        target.NumberFormat = style.number_format
    apply_column_alignment(target, style.alignment)


def style_ranges(cells: list[tuple[int, int]], *, limit: int = 240):
    """Build bounded multi-area ranges so COM formats matching cells in batches."""
    columns = defaultdict(set)
    for row, column in cells:
        columns[column].add(row)
    areas = []
    for column, rows in sorted(columns.items()):
        start = previous = None
        for row in sorted(rows):
            if start is not None and row != previous+1:
                letter = get_column_letter(column)
                areas.append(f"{letter}{start}:{letter}{previous}")
                start = None
            if start is None:
                start = row
            previous = row
        if start is not None:
            letter = get_column_letter(column)
            areas.append(f"{letter}{start}:{letter}{previous}")
    chunk = []
    for area in areas:
        if chunk and len(",".join([*chunk, area])) > limit:
            yield ",".join(chunk)
            chunk = []
        chunk.append(area)
    if chunk:
        yield ",".join(chunk)


def restore_presentation(sheet, snapshot: SheetPresentation) -> None:
    # Restore spacer visibility before expanding the new data rows. Never let
    # an obsolete hidden spacer conceal a newly grown pivot record.
    row_heights = {row: height for row, (height, _) in snapshot.rows.items()}
    hidden = {row: flag for row, (_, flag) in snapshot.rows.items()}
    assignments = defaultdict(list)
    cell_at = sheet.Cells
    for pivot in sheet.PivotTables():
        model = snapshot.pivots[str(pivot.Name)]
        outer, region = pivot.TableRange2, pivot.TableRange1
        start, end, data_column, keys = _body_layout(pivot)
        kind = "summary" if len(list(pivot.DataFields)) > 1 else "phase"
        field_names = {}
        for measure in pivot.DataFields:
            fallback = ("measure", str(measure.SourceName).casefold(), int(measure.Function), int(measure.Calculation))
            styles = model.defaults.get(fallback) or snapshot.peers.get((kind, fallback))
            if styles:
                number_format = styles.most_common(1)[0][0].number_format
                if measure.NumberFormat != number_format:
                    measure.NumberFormat = number_format
        for r in range(int(outer.Row), start):
            offset = r-int(outer.Row)
            height, flag = model.header_rows.get(offset, (19.5, False))
            row_heights[r], hidden[r] = height, flag
            for c in range(int(outer.Column), int(outer.Column+outer.Columns.Count)):
                style = model.headers.get((offset, c-int(outer.Column)))
                if style is None:
                    candidates = [(column, value) for (row, column), value in model.headers.items() if row == offset]
                    style = max(candidates, key=lambda item: item[0])[1] if candidates else None
                if style is not None:
                    position = (offset, c-int(outer.Column))
                    is_new_measure_heading = (
                        position not in model.populated_headers and c >= data_column
                        and sheet.Cells(r, c).Value2 not in (None, "")
                    )
                    style = header_style(style, is_new_measure_heading=is_new_measure_heading)
                    assignments[style].append((r, c))
        for r in range(start, end+1):
            identity = _row_identity(sheet, pivot, r, data_column, field_names=field_names,
                                     data_columns=len(keys), data_rows=end-start+1, cell_at=cell_at) if keys else ()
            role = row_role(identity, model.row_fields)
            heights = model.height_defaults.get(role) or snapshot.peer_heights.get((kind, role))
            row_heights[r] = model.heights.get(identity, heights.most_common(1)[0][0] if heights else 19.5)
            hidden[r] = False
            for c in range(int(region.Column), int(region.Column+region.Columns.Count)):
                key = ("label", c-int(region.Column)) if c < data_column else ("measure", keys[c-data_column] if keys else ())
                style = select_cell_style(model, snapshot.peers, kind, identity, key)
                style = display_style(style, kind, key)
                if style is not None:
                    assignments[style].append((r, c))
    for style, cells in assignments.items():
        for address in style_ranges(cells):
            _apply(sheet.Range(address), style)
    # Batch equal adjacent geometry instead of setting every worksheet row.
    first = previous = None
    state = None
    for row in sorted(row_heights):
        current = (row_heights[row], hidden.get(row, False))
        if first is not None and (row != previous+1 or current != state):
            target = sheet.Rows(f"{first}:{previous}")
            target.RowHeight, target.Hidden = state
            first = None
        if first is None:
            first, state = row, current
        previous = row
    if first is not None:
        target = sheet.Rows(f"{first}:{previous}")
        target.RowHeight, target.Hidden = state
