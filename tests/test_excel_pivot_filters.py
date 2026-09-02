import unittest
from datetime import date

from excel_pivot_filters import (
    apply_month_target_filter,
    configure_pivot_cache_for_current_source,
    is_dynamic_target_header,
    month_pivot_labels,
    set_page_field_to_all,
    set_visible_items_exact,
)


class FakeItem:
    def __init__(self, name):
        self.Name = name
        self.Value = name
        self.Visible = True
        self.Position = None


class FakeCache:
    def __init__(self):
        self.MissingItemsLimit = 1
        self.RefreshOnFileOpen = True


class FakeItems:
    def __init__(self, values):
        self.values = values

    def __call__(self):
        return self.values


class FakeField:
    def __init__(self, name, values, orientation=3):
        self.Name = name
        self.Orientation = orientation
        self.EnableMultiplePageItems = False
        self.IncludeNewItemsInFilter = True
        self._items = [FakeItem(value) for value in values]
        self._current_page = None

    @property
    def CurrentPage(self):
        return self._current_page

    @CurrentPage.setter
    def CurrentPage(self, value):
        self._current_page = value
        if value == "(All)":
            return
        for item in self._items:
            item.Visible = item.Name == value

    def PivotItems(self):
        return self._items

    def ClearAllFilters(self):
        for item in self._items:
            item.Visible = True


class FakeFields:
    def __init__(self, fields):
        self.fields = fields
        self.Count = len(fields)

    def __call__(self, index):
        return self.fields[index - 1]


class FakePivot:
    def __init__(self, fields):
        self.fields = FakeFields(fields)

    def PivotFields(self):
        return self.fields


class PivotFilterTests(unittest.TestCase):
    def test_recognizes_only_dynamic_target_header(self):
        self.assertTrue(is_dynamic_target_header("TARGET  Detemined as 1 Sep 26"))
        self.assertFalse(is_dynamic_target_header("Target SO Complete Date"))

    def test_builds_month_labels(self):
        labels = month_pivot_labels(date(2026, 9, 1))
        self.assertEqual(labels.target, "Target September")
        self.assertEqual(labels.before, "Before September")
        self.assertEqual(labels.after, "After September")

    def test_applies_exact_visible_items(self):
        field = FakeField("Phase", ["04-Pre Installation", "Cancel", "SO Complete"])

        self.assertTrue(set_visible_items_exact(field, {"04-Pre Installation"}))

        self.assertEqual(
            [item.Name for item in field.PivotItems() if item.Visible],
            ["04-Pre Installation"],
        )

    def test_applies_month_filter_to_dynamic_field(self):
        field = FakeField(
            "TARGET  Detemined as 1 Sep 26",
            ["After September", "Before September", "Target Not Yet Inputted", "Target September"],
            orientation=1,
        )
        pivot = FakePivot([field])

        self.assertTrue(apply_month_target_filter(pivot, date(2026, 9, 1), "target_after"))
        self.assertEqual(
            {item.Name for item in field.PivotItems() if item.Visible},
            {"Target September", "After September"},
        )
        positions = {item.Name: item.Position for item in field.PivotItems()}
        self.assertEqual(positions["Target September"], 1)
        self.assertEqual(positions["After September"], 2)

    def test_discards_items_missing_from_current_source(self):
        cache = FakeCache()

        self.assertTrue(configure_pivot_cache_for_current_source(cache))

        self.assertEqual(cache.MissingItemsLimit, 0)
        self.assertFalse(cache.RefreshOnFileOpen)

    def test_page_field_all_disables_multi_select(self):
        field = FakeField("TARGET  Detemined as 1 Sep 26", ["Target September"])

        self.assertTrue(set_page_field_to_all(field))

        self.assertFalse(field.EnableMultiplePageItems)
        self.assertEqual(field.CurrentPage, "(All)")
        self.assertFalse(field.IncludeNewItemsInFilter)


if __name__ == "__main__":
    unittest.main()
