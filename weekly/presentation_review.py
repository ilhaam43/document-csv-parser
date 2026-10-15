"""Attributed overrides for manually approved weekly PowerPoint chart points."""

from __future__ import annotations

import csv
from decimal import Decimal, InvalidOperation
from pathlib import Path


def load_presentation_review(path: str | Path | None) -> dict[tuple[int, int, str], int]:
    if path is None:
        return {}
    review_path = Path(path)
    if review_path.suffix.casefold() != ".csv":
        raise ValueError("Weekly presentation review must be a CSV file.")
    required = {"chart", "series", "label", "value", "reviewed_by", "review_reason"}
    result: dict[tuple[int, int, str], int] = {}
    with review_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not required.issubset(reader.fieldnames or ()):
            raise ValueError("Weekly presentation review is missing required columns.")
        for row_number, row in enumerate(reader, start=2):
            chart = (row.get("chart") or "").strip()
            series = (row.get("series") or "").strip()
            label = (row.get("label") or "").strip()
            reviewer = (row.get("reviewed_by") or "").strip()
            reason = (row.get("review_reason") or "").strip()
            try:
                value = Decimal((row.get("value") or "").strip())
            except InvalidOperation as exc:
                raise ValueError(f"Presentation review row {row_number}: invalid value.") from exc
            if not chart.isdecimal() or not 1 <= int(chart) <= 6 or series not in {"1", "2"}:
                raise ValueError(f"Presentation review row {row_number}: invalid chart or series.")
            if not label or not reviewer or not reason or not value.is_finite():
                raise ValueError(f"Presentation review row {row_number}: value and attribution are required.")
            if value != value.to_integral_value():
                raise ValueError(f"Presentation review row {row_number}: chart value must be integral.")
            key = (int(chart), int(series), label)
            if key in result:
                raise ValueError(f"Presentation review row {row_number}: duplicate chart point.")
            result[key] = int(value)
    return result
