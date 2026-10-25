"""Update editable weekly PowerPoints from refreshed service-specific pivot workbooks."""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .dates import parse_report_month
from .presentation_review import load_presentation_review


CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PRES_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"c": CHART_NS, "a": DRAW_NS, "p": PRES_NS}
OVERFLOW_NOTE_URI = "urn:secm:weekly:chart-overflow"
MONTHS = {name: number for number, name in enumerate(
    ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"), 1
)}
MONTHS.update({"Mei": 5, "Agu": 8, "Okt": 10, "Des": 12})
EXPECTED_SERIES = {
    1: ("LEADTIME FAB-CASH", "LEADTIME CASH-COMPLETE"),
    2: ("INCOMING FAB", "MRC+OTC"),
    3: ("INCOMING FAB", "MRC+OTC"),
    4: ("LEADTIME FAB-CASH", "LEADTIME CASH-COMPLETE"),
    5: ("BULK SERVICE", "MOBILE SERVICE"),
    6: ("COMPLETION ORDER", "MRC+OTC"),
}


@dataclass(frozen=True)
class ChartData:
    labels: tuple[str, ...]
    first: tuple[int | float | None, ...]
    second: tuple[int | float | None, ...]


@dataclass(frozen=True)
class ChartOverflow:
    series: str
    label: str
    value: float
    limit: float
    bound: str


def _rounded(value: object) -> int:
    if not isinstance(value, (float, int, Decimal)):
        raise ValueError(f"Pivot metric is not numeric: {value!r}")
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _metric_value(value: object) -> int | float:
    if isinstance(value, bool) or not isinstance(value, (float, int, Decimal)):
        raise ValueError(f"Pivot metric is not numeric: {value!r}")
    number = Decimal(str(value))
    if not number.is_finite():
        raise ValueError(f"Pivot metric is not finite: {value!r}")
    return int(number) if number == number.to_integral_value() else float(number)


def _section(sheet, column: int, title: str) -> int:
    needle = " ".join(title.upper().split())
    for row in range(1, sheet.max_row + 1):
        value = sheet.cell(row, column).value
        if isinstance(value, str) and needle in " ".join(value.upper().split()):
            return row
    raise ValueError(f"PIVOT PERFORMANCE is missing section {title!r}.")


def _pivot_rows(sheet, start: int, label_column: int, value_columns: tuple[int, ...]) -> tuple[dict[str, tuple[object, ...]], tuple[object, ...]]:
    header = next((row for row in range(start, min(start + 24, sheet.max_row) + 1)
                   if str(sheet.cell(row, label_column).value or "").strip() == "Row Labels"), None)
    if header is None:
        raise ValueError(f"Pivot row header missing near row {start}, column {label_column}.")
    values: dict[str, tuple[object, ...]] = {}
    for row in range(header + 1, min(header + 80, sheet.max_row) + 1):
        label = sheet.cell(row, label_column).value
        if label == "Grand Total":
            return values, tuple(sheet.cell(row, column).value for column in value_columns)
        if isinstance(label, str) and label.strip():
            values[label.strip()] = tuple(sheet.cell(row, column).value for column in value_columns)
    raise ValueError(f"Pivot grand total missing near row {start}, column {label_column}.")


def _month_key(label: str) -> tuple[int, int] | None:
    match = re.fullmatch(r"([A-Za-z]{3})-(\d{2})", label)
    if not match or match.group(1).title() not in MONTHS:
        return None
    return (2000 + int(match.group(2)), MONTHS[match.group(1).title()])


def extract_mobile_charts(workbook: str | Path, report_date: date, report_week: int) -> tuple[dict[int, ChartData], dict[str, int]]:
    source = load_workbook(workbook, read_only=True, data_only=True)
    try:
        if "PIVOT PERFORMANCE" not in source:
            raise ValueError("Mobile workbook has no PIVOT PERFORMANCE sheet.")
        sheet = source["PIVOT PERFORMANCE"]
        revenue = _section(sheet, 1, "INCOMING FAB (MRC)")
        quantity = _section(sheet, 1, "INCOMING FAB (QTY)")
        fab_cash = _section(sheet, 1, "LEAD TIME FAB - CASH")
        cash_complete = _section(sheet, 10, "LEAD TIME CASH-COMPLETE")

        rev_month, _ = _pivot_rows(sheet, revenue, 1, (4,))
        qty_month, _ = _pivot_rows(sheet, quantity, 1, (4,))
        cash_month, cash_total = _pivot_rows(sheet, cash_complete, 10, (11, 12))
        rev_bulk, _ = _pivot_rows(sheet, revenue, 7, (8,))
        rev_mobile, _ = _pivot_rows(sheet, revenue, 10, (11,))
        qty_bulk, _ = _pivot_rows(sheet, quantity, 7, (8,))
        qty_mobile, _ = _pivot_rows(sheet, quantity, 10, (11,))
        fab_bulk, fab_bulk_total = _pivot_rows(sheet, fab_cash, 7, (8,))
        fab_mobile, fab_mobile_total = _pivot_rows(sheet, fab_cash, 10, (11,))
        cash_bulk, cash_bulk_total = _pivot_rows(sheet, cash_complete, 1, (2,))
        cash_mobile, cash_mobile_total = _pivot_rows(sheet, cash_complete, 7, (8,))

        weekly_sources = (rev_bulk, rev_mobile, qty_bulk, qty_mobile,
                          fab_bulk, fab_mobile, cash_bulk, cash_mobile)
        weeks = [key for key in qty_bulk if re.fullmatch(r"W\d{1,2}", key)]
        labels = tuple(weeks[-4:])
        if not labels or labels[-1] != f"W{report_week}":
            raise ValueError(f"Mobile pivot does not end at W{report_week}.")
        for rows in weekly_sources:
            if any(label not in rows for label in labels):
                raise ValueError("Weekly Mobile pivot sections do not share the same four weeks.")

        def weekly(rows_a, rows_b):
            return ChartData(labels,
                             tuple(_metric_value(rows_a[label][0]) for label in labels),
                             tuple(_metric_value(rows_b[label][0]) for label in labels))

        months = sorted((key, label) for label in qty_month
                        if (key := _month_key(label)) and key[0] == report_date.year
                        and key[1] <= report_date.month)
        if not months or months[-1][0][1] != report_date.month:
            raise ValueError("Monthly Mobile pivot does not include the report month.")
        month_labels = tuple(f"{label[:3].upper()}-{report_date.year % 100:02d}" for _, label in months)
        for rows in (rev_month, cash_month):
            if any(label not in rows for _, label in months):
                raise ValueError("Monthly Mobile pivot sections have different month coverage.")
        charts = {
            1: weekly(fab_bulk, cash_bulk),
            2: weekly(qty_bulk, rev_bulk),
            3: weekly(qty_mobile, rev_mobile),
            4: weekly(fab_mobile, cash_mobile),
            5: ChartData(month_labels,
                         tuple(_metric_value(cash_month[label][0]) for _, label in months),
                         tuple(_metric_value(cash_month[label][1]) for _, label in months)),
            6: ChartData(month_labels,
                         tuple(_metric_value(qty_month[label][0]) for _, label in months),
                         tuple(_metric_value(rev_month[label][0]) for _, label in months)),
        }
        averages = {
            "bulk_fab": _rounded(fab_bulk_total[0]),
            "bulk_cash": _rounded(cash_bulk_total[0]),
            "mobile_fab": _rounded(fab_mobile_total[0]),
            "mobile_cash": _rounded(cash_mobile_total[0]),
            "monthly_bulk_cash": _rounded(cash_total[0]),
            "monthly_mobile_cash": _rounded(cash_total[1]),
        }
        return charts, averages
    finally:
        source.close()


def _xml(data: bytes):
    return etree.fromstring(data)


def _bytes(root) -> bytes:
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def _replace_cache(cache, values: tuple[str, ...] | tuple[int | float | None, ...], *, format_code: str = "#,##0") -> None:
    point_count = cache.find(f"{{{CHART_NS}}}ptCount")
    if point_count is None:
        raise ValueError("PowerPoint chart cache has no point count.")
    if cache.tag == f"{{{CHART_NS}}}numCache":
        number_format = cache.find(f"{{{CHART_NS}}}formatCode")
        if number_format is None:
            number_format = etree.Element(f"{{{CHART_NS}}}formatCode")
            cache.insert(0, number_format)
        number_format.text = format_code
    point_count.set("val", str(len(values)))
    for point in cache.findall(f"{{{CHART_NS}}}pt"):
        cache.remove(point)
    for index, value in enumerate(values):
        if value is None:
            continue
        point = etree.SubElement(cache, f"{{{CHART_NS}}}pt", idx=str(index))
        etree.SubElement(point, f"{{{CHART_NS}}}v").text = str(value)


def _chart_axis_overflows(root) -> list[ChartOverflow]:
    outliers: list[ChartOverflow] = []
    for axis in root.xpath("//c:valAx", namespaces=NS):
        axis_id = axis.find("c:axId", namespaces=NS)
        scaling = axis.find("c:scaling", namespaces=NS)
        if axis_id is None or scaling is None:
            continue
        for chart in root.xpath("//c:plotArea/*[c:ser]", namespaces=NS):
            if axis_id.get("val") not in chart.xpath("./c:axId/@val", namespaces=NS):
                continue
            for series in chart.findall("c:ser", namespaces=NS):
                name = "".join(series.xpath("./c:tx//c:v/text()", namespaces=NS)).strip()
                labels = {point.get("idx"): point.findtext("c:v", namespaces=NS)
                          for point in series.xpath("./c:cat//c:pt", namespaces=NS)}
                for point in series.xpath("./c:val//c:pt", namespaces=NS):
                    value = float(point.findtext("c:v", namespaces=NS))
                    for bound in ("min", "max"):
                        limit = scaling.find(f"c:{bound}", namespaces=NS)
                        if limit is None:
                            continue
                        threshold = float(limit.get("val"))
                        outside = value < threshold if bound == "min" else value > threshold
                        if outside:
                            outliers.append(ChartOverflow(
                                name, labels.get(point.get("idx")) or f"Point {int(point.get('idx')) + 1}",
                                value, threshold, bound,
                            ))
    return outliers


def _set_chart_overflow_note(root, outliers: list[ChartOverflow], *, weekly: bool) -> None:
    chart = root.find("c:chart", namespaces=NS)
    if chart is None:
        return
    title = chart.find("c:title", namespaces=NS)
    if title is not None and title.xpath(
        "./c:extLst/c:ext[@uri=$uri]", uri=OVERFLOW_NOTE_URI, namespaces=NS
    ):
        chart.remove(title)
        title = None
        auto = chart.find("c:autoTitleDeleted", namespaces=NS)
        if auto is not None:
            auto.set("val", "1")
    if not outliers:
        return
    if title is not None:
        raise ValueError("Weekly chart has an existing title; cannot place the off-scale note safely.")

    def number(value: float) -> str:
        return f"{value:,.0f}" if value.is_integer() else f"{value:,g}"

    ranked = sorted(outliers, key=lambda item: abs(item.value - item.limit) / max(abs(item.limit), 1), reverse=True)
    if weekly:
        point = ranked[0]
        name = re.sub(r"^LEADTIME\s+", "", point.series, flags=re.IGNORECASE)
        name = re.sub(r"\s*\((?:BRA|SAD)\)", "", name, flags=re.IGNORECASE).strip()
        operator = "<" if point.bound == "min" else ">"
        detail = f"{number(point.value)} ({operator}{number(point.limit)}, off-scale)"
        if len(ranked) > 1:
            detail = f"{number(point.value)} {operator}{number(point.limit)}; +{len(ranked) - 1} off-scale"
        text = f"{name} {point.label}\n{detail}"
    else:
        items = [f"{point.label} {point.series} = {number(point.value)} "
                 f"({'<' if point.bound == 'min' else '>'}{number(point.limit)})" for point in ranked[:2]]
        text = "Off-scale: " + "; ".join(items)
        if len(ranked) > 2:
            remaining = len(ranked) - 2
            text += f"; +{remaining} {'point' if remaining == 1 else 'points'}"

    # An overlay title discloses clipped values without rescaling or replacing data.
    title = etree.Element(f"{{{CHART_NS}}}title")
    rich = etree.SubElement(etree.SubElement(title, f"{{{CHART_NS}}}tx"), f"{{{DRAW_NS}}}rich")
    etree.SubElement(rich, f"{{{DRAW_NS}}}bodyPr", wrap="square", lIns="0", rIns="0", tIns="0", bIns="0")
    etree.SubElement(rich, f"{{{DRAW_NS}}}lstStyle")
    for line in text.split("\n"):
        paragraph = etree.SubElement(rich, f"{{{DRAW_NS}}}p")
        etree.SubElement(paragraph, f"{{{DRAW_NS}}}pPr", algn="l")
        run = etree.SubElement(paragraph, f"{{{DRAW_NS}}}r")
        properties = etree.SubElement(run, f"{{{DRAW_NS}}}rPr", lang="en-US", sz="800", b="1")
        etree.SubElement(etree.SubElement(properties, f"{{{DRAW_NS}}}solidFill"),
                         f"{{{DRAW_NS}}}srgbClr", val="333333")
        etree.SubElement(run, f"{{{DRAW_NS}}}t").text = line
    layout = etree.SubElement(etree.SubElement(title, f"{{{CHART_NS}}}layout"), f"{{{CHART_NS}}}manualLayout")
    for name in ("xMode", "yMode", "wMode", "hMode"):
        etree.SubElement(layout, f"{{{CHART_NS}}}{name}", val="edge" if name in {"xMode", "yMode"} else "factor")
    coordinates = (0.68, 0.22, 0.31, 0.14) if weekly else (0.02, 0.90, 0.95, 0.09)
    for name, value in zip(("x", "y", "w", "h"), coordinates):
        etree.SubElement(layout, f"{{{CHART_NS}}}{name}", val=str(value))
    etree.SubElement(title, f"{{{CHART_NS}}}overlay", val="1")
    extension = etree.SubElement(etree.SubElement(title, f"{{{CHART_NS}}}extLst"),
                                 f"{{{CHART_NS}}}ext", uri=OVERFLOW_NOTE_URI)
    etree.SubElement(extension, f"{{{OVERFLOW_NOTE_URI}}}note")
    chart.insert(0, title)
    auto = chart.find("c:autoTitleDeleted", namespaces=NS)
    if auto is None:
        auto = etree.Element(f"{{{CHART_NS}}}autoTitleDeleted")
        chart.insert(1, auto)
    auto.set("val", "0")


def _update_chart_xml(data: bytes, values: ChartData, first_column: int, chart_number: int,
                      *, expected_series: tuple[str, ...] | None = None,
                      weekly: bool | None = None, number_formats: tuple[str, ...] | None = None,
                      label_position: str | None = None) -> bytes:
    root = _xml(data)
    series = root.xpath("//c:ser", namespaces=NS)
    expected = expected_series or EXPECTED_SERIES[chart_number]
    if len(series) != len(expected):
        raise ValueError(f"Weekly PPT template chart must have exactly {len(expected)} series.")
    titles = [" ".join("".join(node.xpath("./c:tx//c:v/text()", namespaces=NS)).upper().split())
              for node in series]
    if any(name not in title for name, title in zip(expected, titles)):
        raise ValueError(f"Weekly PPT chart {chart_number} series do not match the selected template.")
    last_column = first_column + len(values.labels) - 1
    start = get_column_letter(first_column)
    end = get_column_letter(last_column)
    for index, series_node in enumerate(series):
        for part, row, cache_name, points in (
            ("cat", 1, "strCache", values.labels),
            ("val", index + 2, "numCache", values.first if index == 0 else values.second),
        ):
            ref = series_node.find(f"c:{part}/*", namespaces=NS)
            if ref is None:
                raise ValueError("Weekly PPT template chart reference is missing.")
            formula = ref.find("c:f", namespaces=NS)
            cache = ref.find(f"c:{cache_name}", namespaces=NS)
            if formula is None or cache is None:
                raise ValueError("Weekly PPT template chart cache is missing.")
            formula.text = f"Sheet1!${start}${row}:${end}${row}"
            _replace_cache(cache, points, format_code=number_formats[index] if number_formats else "#,##0")
    if label_position is not None:
        for position in root.xpath("//c:dLbls/c:dLblPos | //c:dLbl/c:dLblPos", namespaces=NS):
            position.set("val", label_position)
    _set_chart_overflow_note(root, _chart_axis_overflows(root), weekly=chart_number <= 4 if weekly is None else weekly)
    return _bytes(root)


def _apply_presentation_review(
    chart_number: int, data: ChartData, decisions: dict[tuple[int, int, str], int]
) -> ChartData:
    values = [list(data.first), list(data.second)]
    for (chart, series, label), value in decisions.items():
        if chart != chart_number:
            continue
        if label not in data.labels:
            raise ValueError(f"Presentation review label {label!r} is absent from chart {chart}.")
        values[series - 1][data.labels.index(label)] = value
    return ChartData(data.labels, tuple(values[0]), tuple(values[1]))


def _update_embedded(data: bytes, values: ChartData, *, weekly: bool, report_week: int,
                     number_formats: tuple[str, ...] | None = None) -> tuple[bytes, int]:
    workbook = load_workbook(BytesIO(data))
    try:
        sheet = workbook.active
        if weekly:
            populated = [column for column in range(2, sheet.max_column + 1)
                         if sheet.cell(1, column).value is not None]
            if not populated:
                raise ValueError("Embedded weekly chart data is empty.")
            last = populated[-1]
            target = last if sheet.cell(1, last).value == f"W{report_week}" else last + 1
            first = target - len(values.labels) + 1
            if first < 2:
                raise ValueError("Embedded weekly chart data has too few weeks.")
            for column in range(2, first):
                label = sheet.cell(1, column).value
                if not isinstance(label, str) or not re.match(r"W\d", label):
                    continue
                if not any(dimension.hidden and dimension.min <= column <= dimension.max
                           for dimension in sheet.column_dimensions.values()
                           if dimension.min is not None and dimension.max is not None):
                    sheet.column_dimensions[get_column_letter(column)].hidden = True
            for column in range(first, target + 1):
                if any(dimension.hidden and dimension.min <= column <= dimension.max
                       for dimension in sheet.column_dimensions.values()
                       if dimension.min is not None and dimension.max is not None):
                    raise ValueError("Embedded chart template hides an active report week.")
        else:
            first = 2
            for column in range(2, max(sheet.max_column, 13) + 1):
                for row in (1, 2, 3):
                    sheet.cell(row, column).value = None
        for offset, label in enumerate(values.labels):
            column = first + offset
            sheet.cell(1, column).value = label
            sheet.cell(2, column).value = values.first[offset]
            sheet.cell(3, column).value = values.second[offset] if values.second else None
            sheet.cell(2, column).number_format = number_formats[0] if number_formats else "#,##0"
            sheet.cell(3, column).number_format = number_formats[1] if number_formats and len(number_formats) > 1 else "#,##0"
        output = BytesIO()
        workbook.save(output)
        return output.getvalue(), first
    finally:
        workbook.close()


def _shape(root, shape_id: str):
    for shape in root.xpath("//p:sp", namespaces=NS):
        properties = shape.find("p:nvSpPr/p:cNvPr", namespaces=NS)
        if properties is not None and properties.get("id") == shape_id:
            return shape
    raise ValueError(f"Weekly PPT template shape {shape_id} is missing.")


def _set_shape_text(root, shape_id: str, value: str, expected: str) -> None:
    runs = _shape(root, shape_id).xpath(".//a:t", namespaces=NS)
    if not runs or expected not in "".join(run.text or "" for run in runs):
        raise ValueError(f"Weekly PPT template shape {shape_id} has unexpected text.")
    runs[0].text = value
    for run in runs[1:]:
        run.text = ""


def _set_annotation(root, shape_id: str, values: tuple[int | None, ...]) -> None:
    runs = _shape(root, shape_id).xpath(".//a:t", namespaces=NS)
    candidates = [run for run in runs if re.search(r"(?::\s*|^\s*)(?:\d+|N/A)\s*$", run.text or "")]
    if len(candidates) != len(values):
        raise ValueError(f"Weekly PPT template annotation {shape_id} has unexpected layout.")
    for run, number in zip(candidates, values):
        run.text = re.sub(r"(?:\d+|N/A)(\s*)$", lambda match: f"{number if number is not None else 'N/A'}{match.group(1)}", run.text)


def _update_slide(data: bytes, slide_number: int, report_date: date, report_week: int,
                  averages: dict[str, int | None], *, report_type: str = "mobile") -> bytes:
    root = _xml(data)
    if slide_number == 1:
        _set_shape_text(root, "8", str(report_date.year), "20")
    else:
        _set_shape_text(root, "25", f"{report_date.month}/{report_date.day}/{report_date.year}", "/")
        date_shape = _shape(root, "25")
        for field in date_shape.xpath(".//a:fld", namespaces=NS):
            parent = field.getparent()
            text_run = etree.Element(f"{{{DRAW_NS}}}r")
            run_properties = field.find(f"{{{DRAW_NS}}}rPr")
            text = field.find(f"{{{DRAW_NS}}}t")
            if run_properties is not None:
                text_run.append(run_properties)
            if text is not None:
                text_run.append(text)
            parent.replace(field, text_run)
        _set_shape_text(root, "27" if slide_number == 2 else "17",
                        f"Weekly Report Group Customer Operations W{report_week}", "Weekly Report Group Customer")
        if slide_number == 2:
            if report_type == "m2m":
                _set_annotation(root, "46", (averages["m2m_fab"], averages["m2m_cash"]))
            else:
                _set_annotation(root, "45", (averages["bulk_fab"], averages["bulk_cash"]))
                _set_annotation(root, "47", (averages["mobile_fab"], averages["mobile_cash"]))
        elif slide_number == 3:
            _set_annotation(root, "15", (averages["monthly_m2m_cash"],) if report_type == "m2m"
                            else (averages["monthly_bulk_cash"], averages["monthly_mobile_cash"]))
    return _bytes(root)


def _embedded_path(archive: ZipFile, chart_number: int) -> str:
    path = f"ppt/charts/_rels/chart{chart_number}.xml.rels"
    root = _xml(archive.read(path))
    targets = [relation.get("Target") for relation in root.findall(f"{{{REL_NS}}}Relationship")
               if relation.get("Type", "").endswith("/package")]
    if len(targets) != 1 or not targets[0].startswith("../embeddings/"):
        raise ValueError(f"Chart {chart_number} has no editable embedded workbook.")
    return "ppt/" + targets[0][3:]


def _validate_chart_placement(archive: ZipFile, *, report_type: str = "mobile") -> None:
    placements: dict[int, tuple[int, int, int]] = {}
    relation_id = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    for slide_number in (2, 3):
        root = _xml(archive.read(f"ppt/slides/slide{slide_number}.xml"))
        relationships = _xml(archive.read(f"ppt/slides/_rels/slide{slide_number}.xml.rels"))
        targets = {relation.get("Id"): relation.get("Target")
                   for relation in relationships.findall(f"{{{REL_NS}}}Relationship")}
        for frame in root.xpath("//p:graphicFrame", namespaces=NS):
            chart = frame.find(".//c:chart", namespaces=NS)
            if chart is None:
                continue
            target = targets.get(chart.get(relation_id), "")
            match = re.search(r"chart(\d+)\.xml$", target)
            offset = frame.find(f"p:xfrm/{{{DRAW_NS}}}off", namespaces=NS)
            if match is None or offset is None:
                raise ValueError("Weekly PPT chart placement is invalid.")
            placements[int(match.group(1))] = (slide_number, int(offset.get("x")), int(offset.get("y")))
    if report_type == "m2m":
        if set(placements) != set(range(1, 5)) or not (
            placements[1][0] == placements[2][0] == 2 and placements[2][2] < placements[1][2]
            and placements[3][0] == placements[4][0] == 3 and placements[4][2] < placements[3][2]
        ):
            raise ValueError("Weekly PPT chart positions do not match the M2M template.")
        return
    if set(placements) != set(range(1, 7)) or not (
        placements[2][0] == placements[3][0] == placements[1][0] == placements[4][0] == 2
        and placements[2][1] < placements[3][1]
        and placements[1][1] < placements[4][1]
        and placements[2][2] < placements[1][2]
        and placements[3][2] < placements[4][2]
        and placements[5][0] == placements[6][0] == 3
        and placements[6][2] < placements[5][2]
    ):
        raise ValueError("Weekly PPT chart positions do not match the Mobile template.")


def write_weekly_presentation(*, mobile_workbook: str | Path, template_pptx: str | Path,
                              output_pptx: str | Path, report_date: date, report_week: int,
                              report_month: date | None = None,
                              presentation_review: str | Path | None = None,
                              report_type: str = "mobile") -> Path:
    """Make a new PPTX, preserving the supplied deck's layout and non-chart content."""
    source_path = Path(template_pptx)
    destination = Path(output_pptx)
    if source_path.resolve() == destination.resolve():
        raise ValueError("Output PPTX must not overwrite the template.")
    if report_type not in {"mobile", "m2m"}:
        raise ValueError("Presentation report_type must be mobile or m2m.")
    chart_series = EXPECTED_SERIES
    number_formats = {}
    if report_type == "m2m":
        from .m2m_presentation import extract_m2m_charts, EXPECTED_SERIES as chart_series, NUMBER_FORMATS as number_formats
        charts, averages = extract_m2m_charts(mobile_workbook, report_month or report_date, report_week)
    else:
        charts, averages = extract_mobile_charts(mobile_workbook, report_month or report_date, report_week)
    review_values = load_presentation_review(presentation_review)
    replacements: dict[str, bytes] = {}
    with ZipFile(source_path) as source:
        if source.testzip() is not None:
            raise ValueError("PowerPoint template is corrupt.")
        slide_parts = sorted(name for name in source.namelist()
                             if re.fullmatch(r"ppt/slides/slide\d+\.xml", name))
        chart_parts = sorted(name for name in source.namelist()
                             if re.fullmatch(r"ppt/charts/chart\d+\.xml", name))
        if len(slide_parts) != 4 or len(chart_parts) != len(chart_series):
            raise ValueError(f"Weekly {report_type} PPT template must have four slides and {len(chart_series)} editable charts.")
        _validate_chart_placement(source, report_type=report_type)
        for number, data in charts.items():
            chart_part = f"ppt/charts/chart{number}.xml"
            chart_xml = source.read(chart_part)
            data = _apply_presentation_review(number, data, review_values)
            embedded = _embedded_path(source, number)
            edited_workbook, first_column = _update_embedded(
                source.read(embedded), data, weekly=number <= (2 if report_type == "m2m" else 4),
                report_week=report_week, number_formats=number_formats.get(number),
            )
            replacements[embedded] = edited_workbook
            replacements[chart_part] = _update_chart_xml(chart_xml, data, first_column, number,
                expected_series=chart_series[number], weekly=number <= (2 if report_type == "m2m" else 4),
                number_formats=number_formats.get(number),
                label_position="t" if report_type == "m2m" and number == 3 else None)
        for slide_number in range(1, 5):
            path = f"ppt/slides/slide{slide_number}.xml"
            replacements[path] = _update_slide(source.read(path), slide_number, report_date,
                                               report_week, averages, report_type=report_type)
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_suffix(".pptx.tmp")
        try:
            with ZipFile(temporary, "w") as output:
                for member in source.infolist():
                    output.writestr(member, replacements.get(member.filename, source.read(member.filename)))
            with ZipFile(temporary) as check:
                if check.testzip() is not None:
                    raise ValueError("Generated PowerPoint ZIP is corrupt.")
            temporary.replace(destination)
        finally:
            temporary.unlink(missing_ok=True)
    return destination


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mobile-workbook", required=True, type=Path)
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--report-date", required=True, type=date.fromisoformat)
    parser.add_argument("--report-month", type=parse_report_month)
    parser.add_argument("--report-week", required=True, type=int)
    parser.add_argument("--presentation-review", type=Path)
    args = parser.parse_args()
    path = write_weekly_presentation(mobile_workbook=args.mobile_workbook,
                                     template_pptx=args.template, output_pptx=args.output,
                                     report_date=args.report_date, report_week=args.report_week,
                                     report_month=args.report_month,
                                     presentation_review=args.presentation_review)
    print(json.dumps({"output": str(path.resolve())}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
