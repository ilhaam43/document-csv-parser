"""Shared Excel COM helpers for deterministic monthly PivotTable filters."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from typing import Callable, Iterable


XL_ROW_FIELD = 1
XL_PAGE_FIELD = 3
XL_MISSING_ITEMS_NONE = 0


def normalize_pivot_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def normalize_pivot_key(value: object) -> str:
    return re.sub(r"[^a-z0-9#]+", "", normalize_pivot_text(value))


def is_dynamic_target_header(value: object) -> bool:
    normalized = normalize_pivot_key(value)
    return normalized.startswith("target") and (
        "deteminedas1" in normalized or "determinedas1" in normalized
    )


@dataclass(frozen=True)
class MonthPivotLabels:
    target: str
    before: str
    after: str
    not_yet_inputted: str = "Target Not Yet Inputted"

    @property
    def all_buckets(self) -> tuple[str, ...]:
        return (self.after, self.before, self.not_yet_inputted, self.target)


def month_pivot_labels(report_date: date) -> MonthPivotLabels:
    month_name = report_date.strftime("%B")
    return MonthPivotLabels(
        target=f"Target {month_name}",
        before=f"Before {month_name}",
        after=f"After {month_name}",
    )


def find_pivot_field(pivot_table, predicate: Callable[[object], bool]):
    try:
        fields = pivot_table.PivotFields()
        for field_index in range(1, fields.Count + 1):
            field = fields(field_index)
            if predicate(field.Name):
                return field
    except Exception:
        return None
    return None


def find_dynamic_target_field(pivot_table):
    return find_pivot_field(pivot_table, is_dynamic_target_header)


def find_named_pivot_field(pivot_table, accepted_names: Iterable[str]):
    accepted_keys = {normalize_pivot_key(value) for value in accepted_names}
    return find_pivot_field(
        pivot_table,
        lambda value: normalize_pivot_key(value) in accepted_keys,
    )


def set_page_field_to_all(field) -> bool:
    """Put a page field in Excel's true `(All)` mode."""
    if field is None:
        return False
    try:
        field.ClearAllFilters()
    except Exception:
        pass
    try:
        field.Orientation = XL_PAGE_FIELD
    except Exception:
        pass
    try:
        field.EnableMultiplePageItems = False
    except Exception:
        pass
    try:
        field.CurrentPage = "(All)"
    except Exception:
        pass
    try:
        field.IncludeNewItemsInFilter = False
    except Exception:
        pass
    return True


def configure_pivot_cache_for_current_source(pivot_cache) -> bool:
    """Prevent removed source values from leaking into refreshed PivotTables."""
    if pivot_cache is None:
        return False
    try:
        pivot_cache.MissingItemsLimit = XL_MISSING_ITEMS_NONE
    except Exception:
        return False
    try:
        pivot_cache.RefreshOnFileOpen = False
    except Exception:
        pass
    return True


def set_pivot_item_order(field, ordered_values: Iterable[str]) -> bool:
    """Apply a stable display order to matching PivotItems."""
    if field is None:
        return False

    try:
        items = list(field.PivotItems())
    except Exception:
        return False

    items_by_key = {}
    for item in items:
        item_key = normalize_pivot_key(getattr(item, "Value", item.Name))
        items_by_key.setdefault(item_key, item)
        items_by_key.setdefault(normalize_pivot_key(item.Name), item)

    changed = False
    for position, value in enumerate(ordered_values, start=1):
        item = items_by_key.get(normalize_pivot_key(value))
        if item is None:
            continue
        try:
            item.Position = position
            changed = True
        except Exception:
            pass
    return changed


def set_visible_items_exact(field, allowed_values: Iterable[str]) -> bool:
    """Show only the requested PivotItems, without ever hiding every item."""
    if field is None:
        return False

    allowed_keys = {normalize_pivot_key(value) for value in allowed_values}
    try:
        items = list(field.PivotItems())
    except Exception:
        return False

    def is_allowed(item) -> bool:
        return (
            normalize_pivot_key(getattr(item, "Value", item.Name)) in allowed_keys
            or normalize_pivot_key(item.Name) in allowed_keys
        )

    matching_items = [item for item in items if is_allowed(item)]
    if not matching_items:
        return False

    orientation = None
    try:
        orientation = int(field.Orientation)
    except Exception:
        pass

    if orientation == XL_PAGE_FIELD and len(matching_items) == 1:
        try:
            field.ClearAllFilters()
        except Exception:
            pass
        try:
            field.EnableMultiplePageItems = False
            field.CurrentPage = str(matching_items[0].Name)
            field.IncludeNewItemsInFilter = False
            return True
        except Exception:
            pass

    try:
        if orientation == XL_PAGE_FIELD:
            field.EnableMultiplePageItems = True
    except Exception:
        pass
    try:
        field.ClearAllFilters()
    except Exception:
        pass
    try:
        field.IncludeNewItemsInFilter = False
    except Exception:
        pass

    for item in matching_items:
        try:
            item.Visible = True
        except Exception:
            pass
    for item in items:
        if is_allowed(item):
            continue
        try:
            item.Visible = False
        except Exception:
            pass
    return True


def apply_month_target_filter(
    pivot_table,
    report_date: date,
    mode: str,
    *,
    orientation: int | None = None,
    position: int | None = None,
) -> bool:
    field = find_dynamic_target_field(pivot_table)
    if field is None:
        return False

    if orientation is not None:
        try:
            field.Orientation = orientation
        except Exception:
            return False
    if position is not None:
        try:
            field.Position = position
        except Exception:
            pass

    labels = month_pivot_labels(report_date)
    filters = {
        "target": (labels.target,),
        "before": (labels.before,),
        "after": (labels.after,),
        "not_yet": (labels.not_yet_inputted,),
        "target_after": (labels.target, labels.after),
        "before_not_yet": (labels.before, labels.not_yet_inputted),
        "all_buckets": labels.all_buckets,
    }
    if mode == "all":
        return set_page_field_to_all(field)
    if mode not in filters:
        raise ValueError(f"Unknown monthly PivotTable filter mode: {mode}")
    selected_items = filters[mode]
    changed = set_visible_items_exact(field, selected_items)
    if changed and len(selected_items) > 1:
        set_pivot_item_order(field, selected_items)
    return changed
