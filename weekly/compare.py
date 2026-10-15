"""Compare generated weekly workbooks with a reviewed reference by cell value."""

from __future__ import annotations

import argparse
import json
import posixpath
from collections import Counter
from datetime import datetime
from itertools import zip_longest
from pathlib import Path
from typing import Any
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook

from .constants import MASTER_COLUMNS


_SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NUMBER_COLUMN = MASTER_COLUMNS.index("No.") + 1
_FAB_UPLOAD_COLUMN = MASTER_COLUMNS.index("fab_upload") + 1


def _equal(expected: Any, actual: Any) -> bool:
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        if isinstance(expected, bool) or isinstance(actual, bool):
            return expected is actual
        return abs(expected - actual) <= 1e-8
    return expected == actual


def _business_equal(sheet_name: str, row_number: int, column_number: int, expected: Any, actual: Any) -> bool:
    if sheet_name == "ALL DATA" and row_number > 1 and column_number == _NUMBER_COLUMN:
        return True
    if sheet_name == "ALL DATA" and row_number > 1 and column_number == _FAB_UPLOAD_COLUMN and isinstance(expected, datetime) and isinstance(actual, datetime):
        return expected.replace(second=0, microsecond=0) == actual.replace(second=0, microsecond=0)
    return _equal(expected, actual)


def _chart_references(path: Path) -> dict[str, list[str]]:
    namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    with ZipFile(path) as archive:
        return {
            name: [
                element.text or ""
                for element in ElementTree.fromstring(archive.read(name)).iter(f"{{{namespace}}}f")
            ]
            for name in sorted(archive.namelist())
            if name.startswith("xl/charts/chart") and name.endswith(".xml")
        }


def _hidden_content_rows(sheet_xml) -> set[int]:
    rows: set[int] = set()
    content_tags = {f"{{{_SHEET_NS}}}{name}" for name in ("v", "f", "is")}
    for _, element in ElementTree.iterparse(sheet_xml, events=("end",)):
        if element.tag != f"{{{_SHEET_NS}}}row":
            continue
        if element.get("hidden") == "1" and any(
            child.tag == f"{{{_SHEET_NS}}}c"
            and any(value.tag in content_tags for value in child)
            for child in element
        ):
            rows.add(int(element.attrib["r"]))
        element.clear()
    return rows


def _workbook_hidden_content_rows(path: Path) -> dict[str, set[int]]:
    with ZipFile(path) as archive:
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            item.attrib["Id"]: item.attrib["Target"]
            for item in relationships
            if item.attrib.get("Type", "").endswith("/worksheet")
        }
        sheets = workbook.find(f"{{{_SHEET_NS}}}sheets")
        if sheets is None:
            raise ValueError(f"Workbook has no sheets: {path}")
        result: dict[str, set[int]] = {}
        for sheet in sheets:
            target = targets[sheet.attrib[f"{{{_OFFICE_REL_NS}}}id"]].lstrip("/")
            xml_path = (
                target if target.startswith("xl/")
                else posixpath.normpath(posixpath.join("xl", target))
            )
            with archive.open(xml_path) as content:
                result[sheet.attrib["name"]] = _hidden_content_rows(content)
        return result


def compare_weekly_workbooks(
    reference: str | Path,
    generated: str | Path,
    *,
    example_limit: int = 20,
    business_only: bool = False,
) -> dict[str, Any]:
    reference_path = Path(reference)
    generated_path = Path(generated)
    expected_book = load_workbook(reference_path, read_only=True, data_only=True)
    actual_book = load_workbook(generated_path, read_only=True, data_only=True)
    try:
        sheets: dict[str, Any] = {}
        for name in dict.fromkeys((*expected_book.sheetnames, *actual_book.sheetnames)):
            if name not in expected_book or name not in actual_book:
                sheets[name] = {"difference_count": 1, "examples": ["Sheet missing"]}
                continue
            expected_sheet = expected_book[name]
            actual_sheet = actual_book[name]
            max_column = max(expected_sheet.max_column, actual_sheet.max_column)
            expected_rows = expected_sheet.iter_rows(max_col=max_column, values_only=True)
            actual_rows = actual_sheet.iter_rows(max_col=max_column, values_only=True)
            difference_count = 0
            column_differences: Counter[int] = Counter()
            examples: list[dict[str, Any]] = []
            for row_number, (expected_row, actual_row) in enumerate(
                zip_longest(expected_rows, actual_rows), start=1
            ):
                empty_row = (None,) * max_column
                for column_number, (expected_value, actual_value) in enumerate(
                    zip(expected_row or empty_row, actual_row or empty_row), start=1
                ):
                    if (
                        _business_equal(name, row_number, column_number, expected_value, actual_value)
                        if business_only else _equal(expected_value, actual_value)
                    ):
                        continue
                    difference_count += 1
                    column_differences[column_number] += 1
                    if len(examples) < example_limit:
                        examples.append({
                            "row": row_number,
                            "column": column_number,
                            "reference": str(expected_value),
                            "generated": str(actual_value),
                        })
            sheets[name] = {
                "difference_count": difference_count,
                "differences_by_column": dict(sorted(column_differences.items())),
                "examples": examples,
            }
    finally:
        expected_book.close()
        actual_book.close()
    expected_charts = _chart_references(reference_path)
    actual_charts = _chart_references(generated_path)
    expected_hidden = _workbook_hidden_content_rows(reference_path)
    actual_hidden = _workbook_hidden_content_rows(generated_path)
    visibility_differences = {
        name: {
            "hidden_only_in_reference": sorted(expected_hidden.get(name, set()) - actual_hidden.get(name, set()))[:20],
            "hidden_only_in_generated": sorted(actual_hidden.get(name, set()) - expected_hidden.get(name, set()))[:20],
            "difference_count": len(expected_hidden.get(name, set()) ^ actual_hidden.get(name, set())),
        }
        for name in dict.fromkeys((*expected_hidden, *actual_hidden))
        if expected_hidden.get(name, set()) != actual_hidden.get(name, set())
    }
    return {
        "mode": "business" if business_only else "strict",
        "matches": all(entry["difference_count"] == 0 for entry in sheets.values())
        and expected_charts == actual_charts
        and (business_only or not visibility_differences),
        "sheets": sheets,
        "chart_references_match": expected_charts == actual_charts,
        "content_visibility_match": not visibility_differences,
        "content_visibility_differences": visibility_differences,
        "reference_chart_references": expected_charts if expected_charts != actual_charts else None,
        "generated_chart_references": actual_charts if expected_charts != actual_charts else None,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--reference", required=True, type=Path)
    parser.add_argument("--generated", required=True, type=Path)
    parser.add_argument("--business-only", action="store_true", help="Ignore No. and FAB_Upload seconds; compare every other cell and chart reference.")
    args = parser.parse_args()
    result = compare_weekly_workbooks(args.reference, args.generated, business_only=args.business_only)
    print(json.dumps(result, indent=2))
    return 0 if result["matches"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
