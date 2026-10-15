"""Reviewed overrides for the manually maintained weekly chart source."""

from __future__ import annotations

import csv
import re
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path


def week_label(year: int, week: int) -> str:
    """Label a Thursday-start Excel week by its starting month."""
    week_start = reporting_week_start(year, week)
    return f"W{(week_start.day - 1) // 7 + 1}-{week_start:%b}"


def reporting_week_start(year: int, week: int) -> date:
    if not 1 <= week <= 53:
        raise ValueError("Report week must be between 1 and 53.")
    january_first = date(year, 1, 1)
    return january_first - timedelta(days=(january_first.weekday() - 3) % 7) + timedelta(weeks=week - 1)


def completion_month(value: str) -> date | None:
    for pattern in ("%b-%y", "%b%y"):
        try:
            return datetime.strptime(value, pattern).date().replace(day=1)
        except ValueError:
            continue
    return None


def load_slide_review(path: str | Path | None) -> dict[str, dict[str, str | float]]:
    if path is None:
        return {"master": {}, "mobile": {}}
    review_path = Path(path)
    if review_path.suffix.casefold() != ".csv":
        raise ValueError("Weekly slide review must be a CSV file.")
    result: dict[str, dict[str, str | float]] = {"master": {}, "mobile": {}}
    with review_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"workbook", "cell", "value", "reviewed_by", "review_reason"}
        if not required.issubset(reader.fieldnames or ()):
            raise ValueError("Weekly slide review is missing required columns.")
        for row_number, row in enumerate(reader, start=2):
            workbook = (row.get("workbook") or "").strip().casefold()
            cell = (row.get("cell") or "").strip().upper()
            reviewer = (row.get("reviewed_by") or "").strip()
            reason = (row.get("review_reason") or "").strip()
            if workbook not in result:
                raise ValueError(f"Slide review row {row_number}: workbook must be master or mobile.")
            match = re.fullmatch(r"([S-V])([1-9]\d*)", cell)
            if match is None or not 10 <= int(match.group(2)) <= 100:
                raise ValueError(f"Slide review row {row_number}: cell must be in S10:V100.")
            if cell in result[workbook]:
                raise ValueError(f"Slide review row {row_number}: duplicate cell {cell}.")
            if not reviewer or not reason:
                raise ValueError(f"Slide review row {row_number}: reviewer and reason are required.")
            raw_value = (row.get("value") or "").strip()
            if match.group(1) == "S":
                if not raw_value or raw_value.startswith("="):
                    raise ValueError(f"Slide review row {row_number}: invalid label.")
                value: str | float = raw_value
            else:
                try:
                    number = Decimal(raw_value)
                except InvalidOperation as exc:
                    raise ValueError(f"Slide review row {row_number}: value must be numeric.") from exc
                if not number.is_finite():
                    raise ValueError(f"Slide review row {row_number}: value must be finite.")
                value = float(number)
            result[workbook][cell] = value
    return result


def write_slide_review_template(path: str | Path) -> Path:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        csv.writer(handle).writerow(("workbook", "cell", "value", "reviewed_by", "review_reason"))
    return output_path
