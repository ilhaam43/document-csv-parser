"""Read IDE package metadata without deserializing full native PivotCaches."""

from __future__ import annotations

import posixpath
from xml.etree import ElementTree as ET

from openpyxl.formatting.formatting import ConditionalFormatting, ConditionalFormattingList


S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def relationship_parts(archive, part: str, *, kind: str | None = None) -> dict[str, str]:
    rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    return {item.get("Id"): (item.get("Target").lstrip("/") if item.get("Target").startswith("/")
                            else posixpath.normpath(posixpath.join(posixpath.dirname(part), item.get("Target"))))
            for item in ET.fromstring(archive.read(rels)) if item.get("TargetMode") != "External"
            and (kind is None or item.get("Type") == kind)}


def worksheet_part(archive, name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = relationship_parts(archive, "xl/workbook.xml")
    for sheet in workbook.findall(f"{S}sheets/{S}sheet"):
        if sheet.get("name") == name:
            return relationships[sheet.get(f"{R}id")]
    raise ValueError(f"IDE workbook has no worksheet {name!r}.")


def conditional_formats(root) -> ConditionalFormattingList:
    result = ConditionalFormattingList()
    for node in root.findall(f"{S}conditionalFormatting"):
        area = ConditionalFormatting.from_tree(node)
        for rule in area.rules:
            result.add(area, rule)
    return result


def validate_table_and_caches(archive, data_root, data_part: str, expected_range: str,
                              record_count: int, pivot_count: int) -> None:
    data_relationships = relationship_parts(archive, data_part)
    tables = [ET.fromstring(archive.read(data_relationships[node.get(f"{R}id")]))
              for node in data_root.findall(f"{S}tableParts/{S}tablePart")]
    table = next((table for table in tables if table.get("name") == "Table1"), None)
    if table is None or table.get("ref") != expected_range:
        raise RuntimeError("Saved IDE List table range differs from records.")

    pivot_part = worksheet_part(archive, "PIVOT")
    pivots = list(relationship_parts(archive, pivot_part, kind=R[1:-1] + "/pivotTable").values())
    if len(pivots) != pivot_count or not pivots:
        raise RuntimeError("IDE List pivots were lost during generation.")
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    workbook_relationships = relationship_parts(archive, "xl/workbook.xml")
    caches = {cache.get("cacheId"): workbook_relationships[cache.get(f"{R}id")]
              for cache in workbook.findall(f"{S}pivotCaches/{S}pivotCache")}
    checked = set()
    for part in pivots:
        pivot = ET.fromstring(archive.read(part))
        cache_part = caches.get(pivot.get("cacheId"))
        if cache_part is None:
            raise RuntimeError(f"Pivot {pivot.get('name')} has no registered cache.")
        if cache_part in checked:
            continue
        cache = ET.fromstring(archive.read(cache_part))
        source = cache.find(f"{S}cacheSource/{S}worksheetSource")
        if source is None or source.get("name") != "Table1" or int(cache.get("recordCount", "-1")) != record_count:
            raise RuntimeError(f"Pivot {pivot.get('name')} is not bound to current Table1 records.")
        checked.add(cache_part)
