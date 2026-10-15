"""Streaming source selection for weekly BQ exports."""

from __future__ import annotations

import csv
from collections import Counter
from collections.abc import Iterable, Iterator, Mapping
from datetime import date
from pathlib import Path
from typing import Any

from .constants import MASTER_SERVICE_GROUPS, SOURCE_COLUMNS
from .dates import parse_datetime, same_reporting_month
from .models import SourceSelectionResult, ValidationIssue
from .rules import clean_text, normalize_identifier, normalize_service_group


_SOURCE_COLUMN_ALIASES = {
    "quote_created": "quote_crea",
    "order_created": "order_crea",
}


def iter_bq_rows(path: str | Path, *, encoding: str = "utf-8-sig") -> Iterator[dict[str, str]]:
    """Stream a large BQ CSV instead of loading it into memory."""
    with Path(path).open("r", encoding=encoding, newline="") as handle:
        reader = csv.DictReader(handle)
        actual_columns = tuple(
            _SOURCE_COLUMN_ALIASES.get(column, column)
            for column in (reader.fieldnames or ())
        )
        duplicates = sorted(
            column for column, count in Counter(actual_columns).items()
            if count > 1 and column in SOURCE_COLUMNS
        )
        if duplicates:
            raise ValueError(f"BQ CSV has duplicate source columns: {', '.join(duplicates)}")
        missing = [column for column in SOURCE_COLUMNS if column not in actual_columns]
        if missing:
            raise ValueError(f"BQ CSV is missing required columns: {', '.join(missing)}")
        reader.fieldnames = list(actual_columns)
        for row in reader:
            yield dict(row)


def select_new_source_rows(
    rows: Iterable[Mapping[str, Any]],
    *,
    existing_quote_nums: set[str],
    report_date: date,
    report_month: date | None = None,
) -> SourceSelectionResult:
    """Apply the reviewed quote-level, service, month, and cutoff filters."""
    result = SourceSelectionResult()
    known_quotes = {normalize_identifier(value) for value in existing_quote_nums}

    for source_row_number, raw_row in enumerate(rows, start=2):
        row = dict(raw_row)
        quote_num = clean_text(row.get("quote_num"))
        if not quote_num:
            result.rejected_reasons["missing_quote_num"] += 1
            result.issues.append(
                ValidationIssue(
                    code="missing_quote_num",
                    message="Source row has no quote_num.",
                    row_number=source_row_number,
                )
            )
            continue
        if normalize_identifier(quote_num) in known_quotes:
            result.rejected_reasons["existing_quote_num"] += 1
            continue

        service_group = normalize_service_group(row.get("service_group"))
        if service_group not in MASTER_SERVICE_GROUPS:
            result.rejected_reasons["service_group_not_in_scope"] += 1
            continue

        try:
            fab_upload = parse_datetime(row.get("FAB_Upload"))
        except ValueError as exc:
            result.rejected_reasons["invalid_fab_upload"] += 1
            result.issues.append(
                ValidationIssue(
                    code="invalid_fab_upload",
                    message=str(exc),
                    row_number=source_row_number,
                )
            )
            continue
        if fab_upload is None:
            result.rejected_reasons["missing_fab_upload"] += 1
            continue
        if not same_reporting_month(fab_upload, report_month or report_date):
            result.rejected_reasons["outside_reporting_month"] += 1
            continue
        if fab_upload.date() > report_date:
            result.rejected_reasons["after_report_cutoff"] += 1
            continue

        row["service_group"] = service_group
        result.selected_rows.append(row)

    return result
