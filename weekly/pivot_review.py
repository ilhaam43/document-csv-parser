"""Attributed review decisions for weekly PivotTable adjustment filters."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class PivotReviewDecision:
    pivot: str
    field: str
    item: str
    visible: bool | None = None
    position: int | None = None


def load_pivot_review(path: str | Path | None) -> dict[str, list[PivotReviewDecision]]:
    result: dict[str, list[PivotReviewDecision]] = {"master": [], "mobile": []}
    if path is None:
        return result
    review_path = Path(path)
    if review_path.suffix.casefold() != ".csv":
        raise ValueError("Weekly pivot review must be a CSV file.")
    required = {"workbook", "pivot", "item", "visible", "reviewed_by", "review_reason"}
    seen: set[tuple[str, str, str, str]] = set()
    with review_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not required.issubset(reader.fieldnames or ()):
            raise ValueError("Weekly pivot review is missing required columns.")
        for row_number, row in enumerate(reader, start=2):
            workbook = (row.get("workbook") or "").strip().casefold()
            pivot = (row.get("pivot") or "").strip()
            item = (row.get("item") or "").strip()
            field = (row.get("field") or "ADJUSTMENT").strip()
            visible = (row.get("visible") or "").strip().casefold()
            raw_position = (row.get("position") or "").strip()
            reviewer = (row.get("reviewed_by") or "").strip()
            reason = (row.get("review_reason") or "").strip()
            if workbook not in result:
                raise ValueError(f"Pivot review row {row_number}: invalid workbook.")
            if not pivot or not item or not reviewer or not reason:
                raise ValueError(f"Pivot review row {row_number}: decision and attribution are required.")
            if field == "ADJUSTMENT":
                if visible not in {"true", "false"} or raw_position:
                    raise ValueError(f"Pivot review row {row_number}: adjustment requires visible.")
                decision = PivotReviewDecision(pivot, field, item, visible == "true")
            elif field in {"Week FAB", "Week Completion"}:
                if visible or not raw_position.isdecimal() or int(raw_position) < 1:
                    raise ValueError(f"Pivot review row {row_number}: week requires position.")
                decision = PivotReviewDecision(pivot, field, item, position=int(raw_position))
            else:
                raise ValueError(f"Pivot review row {row_number}: unsupported field.")
            key = (workbook, pivot, field, item)
            if key in seen:
                raise ValueError(f"Pivot review row {row_number}: duplicate decision.")
            seen.add(key)
            result[workbook].append(decision)
    return result
