"""Audit IDE ALL ORDER values against a reference without changing either file."""

from __future__ import annotations

import argparse
import json
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import load_workbook

from generate_ide_tracking import (
    ALL_ORDER_SHEET_NAME,
    DATE_HEADERS,
    QUOTE_ID_HEADER,
    is_excel_zero_date,
    normalize_header,
    normalize_quote_id,
    parse_dashboard_field_datetime,
)


def _read_all_order(path: Path) -> tuple[tuple[str, ...], list[tuple[object, ...]]]:
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        if ALL_ORDER_SHEET_NAME not in workbook.sheetnames:
            raise ValueError(f"{path.name} is missing {ALL_ORDER_SHEET_NAME}.")
        rows = workbook[ALL_ORDER_SHEET_NAME].iter_rows(values_only=True)
        first = next(rows, ())
        populated = [i for i, value in enumerate(first) if value is not None]
        if not populated:
            raise ValueError(f"{path.name} has no ALL ORDER headers.")
        width = max(populated) + 1
        headers = tuple(normalize_header(value) if value is not None else "" for value in first[:width])
        if len(set(headers)) != len(headers) or "" in headers:
            raise ValueError(f"{path.name} has blank or duplicate ALL ORDER headers.")
        if QUOTE_ID_HEADER not in headers:
            raise ValueError(f"{path.name} is missing {QUOTE_ID_HEADER}.")
        quote_column = headers.index(QUOTE_ID_HEADER)
        values = [tuple(row[:width]) for row in rows if normalize_quote_id(row[quote_column])]
        return headers, values
    finally:
        workbook.close()


def _empty(value: object) -> object:
    return None if isinstance(value, str) and not value.strip() else value


def _equal(left: object, right: object, header: str) -> bool:
    left, right = _empty(left), _empty(right)
    if header in DATE_HEADERS and is_excel_zero_date(left) and is_excel_zero_date(right):
        return True
    if isinstance(left, bool) != isinstance(right, bool):
        return False
    if isinstance(left, (datetime, date, time)) != isinstance(right, (datetime, date, time)):
        return False
    return left == right


def _difference_kind(left: object, right: object, header: str) -> str:
    if header in DATE_HEADERS and isinstance(left, str) != isinstance(right, str):
        left_date = parse_dashboard_field_datetime(left, header)
        right_date = parse_dashboard_field_datetime(right, header)
        if left_date is not None and left_date == right_date:
            return "representation"
    return "value"


def compare_all_order(
    manual_path: Path,
    generated_path: Path,
    *,
    max_examples: int = 10,
) -> dict[str, object]:
    manual_headers, manual = _read_all_order(manual_path)
    generated_headers, generated = _read_all_order(generated_path)
    mh = {header: i for i, header in enumerate(manual_headers)}
    gh = {header: i for i, header in enumerate(generated_headers)}
    missing = [header for header in manual_headers if header not in gh]
    extra = [header for header in generated_headers if header not in mh]
    column_order_matches = manual_headers == generated_headers
    quote_order_matches = (
        [normalize_quote_id(row[mh[QUOTE_ID_HEADER]]) for row in manual]
        == [normalize_quote_id(row[gh[QUOTE_ID_HEADER]]) for row in generated]
    )
    differences = {}
    changed_rows: set[int] = set()
    changed_cells = representation_cells = 0
    for header in manual_headers:
        if header not in gh:
            continue
        examples = []
        count = representations = 0
        for row_number, (left, right) in enumerate(zip(manual, generated), 2):
            a, b = left[mh[header]], right[gh[header]]
            if _equal(a, b, header):
                continue
            kind = _difference_kind(a, b, header)
            count += 1
            representations += int(kind == "representation")
            changed_rows.add(row_number)
            if len(examples) < max_examples:
                examples.append({
                    "row": row_number, "quote": left[mh[QUOTE_ID_HEADER]],
                    "manual": a, "generated": b, "kind": kind,
                })
        if count:
            differences[header] = {
                "count": count, "value_differences": count - representations,
                "representation_differences": representations, "examples": examples,
            }
            changed_cells += count
            representation_cells += representations
    schema_matches = not missing and not extra and column_order_matches
    same_rows = len(manual) == len(generated) and quote_order_matches
    return {
        "manual": str(manual_path.resolve()), "generated": str(generated_path.resolve()),
        "manual_rows": len(manual), "generated_rows": len(generated),
        "manual_columns": len(manual_headers), "generated_columns": len(generated_headers),
        "quote_order_matches": quote_order_matches,
        "schema": {"matches": schema_matches, "column_order_matches": column_order_matches,
                   "missing_columns": missing, "extra_columns": extra},
        "changed_cells": changed_cells, "changed_rows": len(changed_rows),
        "value_difference_cells": changed_cells - representation_cells,
        "representation_difference_cells": representation_cells,
        "column_differences": differences,
        "exact_value_matches": schema_matches and same_rows and changed_cells == 0,
        "semantic_value_matches": schema_matches and same_rows and changed_cells == representation_cells,
        "comparison_scope": "ALL ORDER only; native PivotTables must be audited separately",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manual", type=Path, required=True)
    parser.add_argument("--generated", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    report = compare_all_order(args.manual, args.generated)
    payload = json.dumps(report, default=str, ensure_ascii=True, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(payload + "\n", encoding="utf-8")
    summary = {key: value for key, value in report.items() if key != "column_differences"}
    print(json.dumps(summary, default=str, indent=2))
    return 0 if report["exact_value_matches"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
