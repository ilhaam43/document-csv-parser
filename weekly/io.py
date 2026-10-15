"""Read-only input adapters for weekly report workbooks."""

from __future__ import annotations

import csv
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterator

from openpyxl import load_workbook

from .constants import (
    DATE_COLUMNS,
    DECIMAL_COLUMNS,
    INTEGER_COLUMNS,
    MASTER_COLUMNS,
    MASTER_SHEET_NAME,
)
from .dates import as_date, parse_datetime
from .models import ValidationIssue
from .rules import to_decimal


def iter_master_rows(path: str | Path) -> Iterator[dict[str, Any]]:
    """Stream the stable 34-column ALL DATA contract from a workbook."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if MASTER_SHEET_NAME not in workbook.sheetnames:
            raise ValueError(f"Workbook has no {MASTER_SHEET_NAME!r} sheet: {path}")
        sheet = workbook[MASTER_SHEET_NAME]
        rows = sheet.iter_rows(min_col=1, max_col=len(MASTER_COLUMNS), values_only=True)
        try:
            actual_headers = tuple(next(rows))
        except StopIteration as exc:
            raise ValueError(f"Workbook has an empty {MASTER_SHEET_NAME!r} sheet: {path}") from exc
        if actual_headers != MASTER_COLUMNS:
            raise ValueError(f"Unexpected ALL DATA schema in {path}")
        for values in rows:
            if not any(value is not None and value != "" for value in values):
                continue
            yield dict(zip(MASTER_COLUMNS, values))
    finally:
        workbook.close()


def load_holidays(path: str | Path, *, sheet_name: str | None = None) -> set[date]:
    """Load holiday dates from column B of the supplied holiday workbook."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        resolved_sheet_name = sheet_name or workbook.sheetnames[0]
        if resolved_sheet_name not in workbook.sheetnames:
            raise ValueError(f"Holiday workbook has no sheet {resolved_sheet_name!r}.")
        sheet = workbook[resolved_sheet_name]
        holidays: set[date] = set()
        for (value,) in sheet.iter_rows(min_col=2, max_col=2, values_only=True):
            parsed = as_date(value)
            if parsed:
                holidays.add(parsed)
        if not holidays:
            raise ValueError("Holiday workbook column B contains no valid dates.")
        return holidays
    finally:
        workbook.close()


def _csv_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, Decimal):
        return format(value, "f")
    return value


def write_master_csv(path: str | Path, records: list[dict[str, Any]]) -> Path:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=MASTER_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for record in records:
            writer.writerow({column: _csv_value(record.get(column)) for column in MASTER_COLUMNS})
    return output_path


def iter_master_csv(path: str | Path) -> Iterator[dict[str, Any]]:
    """Read a generated master CSV back into Excel-compatible typed values."""
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        actual_headers = tuple(reader.fieldnames or ())
        if actual_headers != MASTER_COLUMNS:
            raise ValueError(f"Unexpected generated master CSV schema in {path}")
        for csv_row in reader:
            row: dict[str, Any] = {}
            for column in MASTER_COLUMNS:
                value = csv_row.get(column, "")
                if value == "":
                    row[column] = None
                elif column in DATE_COLUMNS:
                    row[column] = parse_datetime(value)
                elif column in INTEGER_COLUMNS:
                    row[column] = int(value)
                elif column in DECIMAL_COLUMNS:
                    row[column] = to_decimal(value)
                else:
                    row[column] = value
            yield row


def write_issues_csv(path: str | Path, issues: list[ValidationIssue]) -> Path:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    columns = ("code", "message", "quote_num", "order_number", "row_number")
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        for issue in issues:
            writer.writerow(
                {
                    "code": issue.code,
                    "message": issue.message,
                    "quote_num": issue.key.quote_num if issue.key else "",
                    "order_number": issue.key.order_number if issue.key else "",
                    "row_number": issue.row_number or "",
                }
            )
    return output_path
