"""Deterministic matching for refreshes of existing weekly master rows."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Mapping
from typing import Any

from .models import MatchResult, RecordKey, ValidationIssue
from .rules import clean_text, normalize_identifier


def record_key(quote_num: Any, order_number: Any) -> RecordKey:
    return RecordKey(
        quote_num=normalize_identifier(quote_num),
        order_number=normalize_identifier(order_number),
    )


class MasterIndex:
    """Index master rows without silently selecting ambiguous quote matches."""

    def __init__(self, rows: Iterable[Mapping[str, Any]]) -> None:
        self._by_pair: dict[RecordKey, list[int]] = defaultdict(list)
        self._by_quote: dict[str, list[int]] = defaultdict(list)
        for row_number, row in enumerate(rows, start=2):
            quote_num = row.get("quote_num")
            order_number = row.get("order number")
            key = record_key(quote_num, order_number)
            if key.quote_num:
                self._by_quote[key.quote_num].append(row_number)
                self._by_pair[key].append(row_number)

    def resolve(self, source: Mapping[str, Any]) -> MatchResult:
        key = record_key(source.get("quote_num"), source.get("order_number"))
        pair_matches = self._by_pair.get(key, [])
        if len(pair_matches) == 1:
            return MatchResult(row_number=pair_matches[0], strategy="quote_and_order")
        if len(pair_matches) > 1:
            return MatchResult(
                row_number=None,
                strategy="ambiguous_quote_and_order",
                issue=ValidationIssue(
                    code="ambiguous_quote_and_order",
                    message="Multiple master rows match quote_num and order_number.",
                    key=key,
                ),
            )

        quote_matches = self._by_quote.get(key.quote_num, [])
        if len(quote_matches) == 1:
            return MatchResult(row_number=quote_matches[0], strategy="unique_quote")
        if len(quote_matches) > 1:
            return MatchResult(
                row_number=None,
                strategy="ambiguous_quote",
                issue=ValidationIssue(
                    code="ambiguous_quote",
                    message=(
                        "quote_num exists on multiple master rows and order_number did not "
                        "identify one row. Manual review is required."
                    ),
                    key=key,
                ),
            )
        return MatchResult(
            row_number=None,
            strategy="not_found",
            issue=ValidationIssue(
                code="existing_row_not_found",
                message=f"No master row matches quote_num {clean_text(source.get('quote_num'))!r}.",
                key=key,
            ),
        )
