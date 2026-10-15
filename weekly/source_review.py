"""Audit BQ quote-first disagreements and load optional lifecycle overrides."""

from __future__ import annotations

import csv
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

from .refresh import ExistingSourceIndex
from .rules import clean_text, normalize_order_status


SOURCE_REVIEW_COLUMNS = (
    "master_row", "quote_num", "order_number", "ca_name", "service_group",
    "order_status", "order_completion_date", "quote_status",
    "quote_completion_date", "reviewed_source",
)


def automatic_source_choice(status: Any) -> str:
    """Freeze closed orders; refresh active lifecycle using the quote lookup."""
    return "order" if normalize_order_status(status) in {"COMPLETE", "CANCEL"} else "quote"


def collect_source_conflicts(
    previous_rows: Sequence[Mapping[str, Any]],
    source_index: ExistingSourceIndex,
    selected_rows: Sequence[Mapping[str, Any]],
) -> list[dict[str, str]]:
    """Record disagreements; the lifecycle rule chooses a source unless overridden."""
    conflicts: list[dict[str, str]] = []
    first_selected: dict[str, Mapping[str, Any]] = {}
    for source in selected_rows:
        first_selected.setdefault(clean_text(source.get("quote_num")), source)

    def add(row_number: int, row: Mapping[str, Any], order: Mapping[str, Any] | None,
            quote: Mapping[str, Any] | None) -> None:
        if order is None or quote is None:
            return
        order_status = normalize_order_status(order.get("order_status"))
        quote_status = normalize_order_status(quote.get("order_status"))
        if order_status == quote_status:
            return
        conflicts.append({
            "master_row": str(row_number),
            "quote_num": clean_text(row.get("quote_num")),
            "order_number": clean_text(row.get("order number", row.get("order_number"))),
            "ca_name": clean_text(row.get("ca_name")),
            "service_group": clean_text(row.get("service_group")),
            "order_status": clean_text(order.get("order_status")),
            "order_completion_date": clean_text(order.get("order_completion_date")),
            "quote_status": clean_text(quote.get("order_status")),
            "quote_completion_date": clean_text(quote.get("order_completion_date")),
            "reviewed_source": "",
        })

    for row_number, row in enumerate(previous_rows, start=2):
        if normalize_order_status(row.get("order_status")) in {"COMPLETE", "CANCEL"}:
            continue
        order, issue = source_index.resolve(row)
        if issue is not None or order is None:
            continue
        add(row_number, row, order, source_index.first_for_quote(row.get("quote_num")))
    for row_number, row in enumerate(selected_rows, start=len(previous_rows) + 2):
        add(row_number, row, row, first_selected.get(clean_text(row.get("quote_num"))))
    return conflicts


def source_with_reviewed_lifecycle(
    order: Mapping[str, Any], quote: Mapping[str, Any], choice: str,
) -> dict[str, Any]:
    source = dict(order)
    if choice == "quote":
        source["order_status"] = quote.get("order_status")
        source["order_completion_date"] = quote.get("order_completion_date")
    return source


def write_source_review_csv(path: str | Path, rows: Sequence[Mapping[str, str]]) -> Path:
    output = Path(path)
    with output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=SOURCE_REVIEW_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)
    return output


def load_source_review_csv(
    path: str | Path, expected_rows: Sequence[Mapping[str, str]],
) -> dict[int, str]:
    expected = {int(row["master_row"]): row for row in expected_rows}
    decisions: dict[int, str] = {}
    seen: set[int] = set()
    with Path(path).open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not {"master_row", "quote_num", "order_number", "reviewed_source"}.issubset(reader.fieldnames or ()):
            raise ValueError("Source review CSV is missing required columns.")
        for row in reader:
            try:
                number = int(row["master_row"])
            except (TypeError, ValueError) as exc:
                raise ValueError("Source review contains an invalid master_row.") from exc
            target = expected.get(number)
            if target is None or number in seen:
                raise ValueError(f"Source review contains an unexpected or duplicate row {number}.")
            seen.add(number)
            if (row["quote_num"], row["order_number"]) != (target["quote_num"], target["order_number"]):
                raise ValueError(f"Source review identity changed for Master row {number}.")
            choice = clean_text(row["reviewed_source"]).casefold()
            if not choice:
                continue
            if choice not in {"order", "quote"}:
                raise ValueError(f"Source review row {number} must choose order or quote.")
            decisions[number] = choice
    return decisions
