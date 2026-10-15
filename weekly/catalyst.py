"""Catalyst review import and controlled status overrides."""

from __future__ import annotations

import csv
from collections.abc import Iterable, Mapping
from datetime import datetime
from pathlib import Path
from typing import Any

from .io import iter_master_rows
from .matching import record_key
from .models import CatalystOverride, RecordKey, ValidationIssue
from .rules import clean_text, normalize_order_status


def collect_catalyst_overrides(
    rows: Iterable[Mapping[str, Any]],
) -> tuple[dict[RecordKey, CatalystOverride], list[ValidationIssue]]:
    """Read only quote, order, and status; reject conflicting duplicate reviews."""
    overrides: dict[RecordKey, CatalystOverride] = {}
    issues: list[ValidationIssue] = []
    conflicted: set[RecordKey] = set()

    for row_number, row in enumerate(rows, start=2):
        key = record_key(row.get("quote_num"), row.get("order number", row.get("order_number")))
        raw_status = clean_text(
            row.get("reviewed_status")
            if "reviewed_status" in row
            else row.get("order_status")
        )
        if not key.quote_num or not raw_status:
            continue
        candidate = CatalystOverride(key=key, order_status=raw_status)
        existing = overrides.get(key)
        if existing and normalize_order_status(existing.order_status) != normalize_order_status(raw_status):
            conflicted.add(key)
            issues.append(
                ValidationIssue(
                    code="conflicting_catalyst_status",
                    message="Catalyst review contains conflicting statuses for one record key.",
                    key=key,
                    row_number=row_number,
                )
            )
            continue
        overrides[key] = candidate

    for key in conflicted:
        overrides.pop(key, None)
    return overrides, issues


def iter_catalyst_review(path: str | Path):
    """Read either the legacy reviewed workbook or the new checkpoint CSV."""
    review_path = Path(path)
    if review_path.suffix.casefold() == ".csv":
        with review_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            required = {"quote_num", "order_number", "reviewed_status"}
            missing = required.difference(reader.fieldnames or ())
            if missing:
                raise ValueError(
                    "Catalyst review CSV is missing required columns: "
                    + ", ".join(sorted(missing))
                )
            for row in reader:
                yield dict(row)
        return
    yield from iter_master_rows(review_path)


def write_catalyst_review_csv(
    path: str | Path,
    rows: Iterable[Mapping[str, Any]],
    *,
    row_numbers: set[int] | None = None,
) -> Path:
    """Create a stable review checkpoint without modifying source workbooks."""
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    columns = (
        "master_row",
        "quote_num",
        "order_number",
        "ca_name",
        "service_group",
        "fab_upload",
        "current_status",
        "reviewed_status",
    )
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        for master_row, row in enumerate(rows, start=2):
            if row_numbers is not None and master_row not in row_numbers:
                continue
            fab_upload = row.get("fab_upload")
            writer.writerow(
                {
                    "master_row": master_row,
                    "quote_num": clean_text(row.get("quote_num")),
                    "order_number": clean_text(row.get("order number")),
                    "ca_name": clean_text(row.get("ca_name")),
                    "service_group": clean_text(row.get("service_group")),
                    "fab_upload": fab_upload.isoformat(sep=" ")
                    if isinstance(fab_upload, datetime)
                    else clean_text(fab_upload),
                    "current_status": clean_text(row.get("order_status")),
                    "reviewed_status": "",
                }
            )
    return output_path


def apply_catalyst_status(
    master_row: Mapping[str, Any],
    override: CatalystOverride,
) -> dict[str, Any]:
    """Apply the Catalyst-owned status without copying unrelated workbook cells."""
    updated = dict(master_row)
    updated["order_status"] = override.order_status
    normalized = normalize_order_status(override.order_status)
    updated["order_status2"] = normalized
    updated["order_status (Confirmed by MIDO Dept.)"] = normalized
    return updated
