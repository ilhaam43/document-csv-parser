import unittest
from types import SimpleNamespace
from unittest.mock import Mock

from ide_list_format import (
    PROGRESS_BAR_COLOR, capture_progress_bar, restore_progress_bar, validate_progress_bar,
)


class Columns:
    def __init__(self, columns):
        self.columns = columns
        self.Count = len(columns)

    def __call__(self, name):
        if isinstance(name, int):
            return self.columns[name - 1]
        return next(column for column in self.columns if column.Name == name)


def bar(column=1, color=PROGRESS_BAR_COLOR, minimum=0, maximum=1):
    return SimpleNamespace(Type=4, BarColor=SimpleNamespace(Color=color), BarFillType=0,
                           MinPoint=SimpleNamespace(Type=0, Value=minimum, Modify=Mock()),
                           MaxPoint=SimpleNamespace(Type=0, Value=maximum, Modify=Mock()),
                           AppliesTo=SimpleNamespace(Column=column, Columns=SimpleNamespace(Count=1)),
                           ModifyAppliesToRange=Mock())


def table(progress_rules=(), other_rules=()):
    progress = SimpleNamespace(FormatConditions=list(progress_rules))
    columns = Columns([SimpleNamespace(Name="SO Completion Date", DataBodyRange=SimpleNamespace()),
                       SimpleNamespace(Name="Progress", DataBodyRange=progress),
                       SimpleNamespace(Name="Revenue", DataBodyRange=SimpleNamespace())])
    return SimpleNamespace(ListColumns=columns, Range=SimpleNamespace(Column=1),
                           DataBodyRange=SimpleNamespace(FormatConditions=[*progress_rules, *other_rules]))


class IdeProgressBarTests(unittest.TestCase):
    def test_capture_and_rebind_named_progress_after_column_reordering(self):
        original = bar(column=54)
        before = table([original])
        captured = capture_progress_bar(before)
        after = table()
        restore_progress_bar(after, captured)
        self.assertIs(captured, original)
        original.ModifyAppliesToRange.assert_called_once_with(after.ListColumns("Progress").DataBodyRange)
        original.MinPoint.Modify.assert_called_once_with(0, 0)
        original.MaxPoint.Modify.assert_called_once_with(0, 1)
        self.assertEqual(original.BarColor.Color, PROGRESS_BAR_COLOR)
        self.assertEqual(original.BarFillType, 0)
        self.assertTrue(original.ShowValue)
        self.assertEqual((original.PercentMin, original.PercentMax), (0, 100))

    def test_recover_orphan_from_previously_generated_reference(self):
        orphan = bar(column=1)
        reference = table(other_rules=[orphan])
        self.assertIs(capture_progress_bar(reference), orphan)

    def test_prefer_progress_rule_over_another_columns_rule(self):
        correct, other = bar(column=2), bar(column=1)
        self.assertIs(capture_progress_bar(table([correct], [other])), correct)

    def test_do_not_capture_other_conditional_formats_or_unrelated_data_bars(self):
        unrelated = [SimpleNamespace(Type=3), bar(column=3), bar(column=1, color=1),
                     bar(column=1, maximum=365)]
        self.assertIsNone(capture_progress_bar(table(other_rules=unrelated)))

    def test_create_manual_style_only_when_no_existing_bar_is_available(self):
        reference = table()
        created = bar()
        conditions = SimpleNamespace(AddDatabar=Mock(return_value=created))
        reference.ListColumns("Progress").DataBodyRange.FormatConditions = conditions
        restore_progress_bar(reference, None)
        conditions.AddDatabar.assert_called_once_with()
        self.assertEqual(created.BarColor.Color, PROGRESS_BAR_COLOR)
        self.assertEqual(created.BarFillType, 0)
        self.assertTrue(created.ShowValue)
        created.ModifyAppliesToRange.assert_not_called()

    def test_preserve_reference_bar_appearance_when_rebinding(self):
        original = bar(color=12345)
        original.BarFillType = 1
        restore_progress_bar(table(), original)
        self.assertEqual(original.BarColor.Color, 12345)
        self.assertEqual(original.BarFillType, 1)

    def saved_sheet(self, ranges=("J2:J2076",), points=(("num", 0), ("num", 1))):
        from openpyxl.formatting.formatting import ConditionalFormatting, ConditionalFormattingList
        from openpyxl.formatting.rule import Rule, DataBar, FormatObject

        conditions = ConditionalFormattingList()
        for area in ranges:
            conditions.add(ConditionalFormatting(sqref=area), Rule(type="dataBar", dataBar=DataBar(
                cfvo=[FormatObject(type=kind, val=value) for kind, value in points], color="7FD3CA")))
        return SimpleNamespace(conditional_formatting=conditions)

    def test_saved_rule_covers_all_current_rows_not_the_header(self):
        validate_progress_bar(self.saved_sheet(), 10, 2075)
        validate_progress_bar(self.saved_sheet(("D2:D6",)), 4, 5)

    def test_reject_missing_stale_partial_or_duplicate_progress_rules(self):
        for ranges in ((), ("AT2:AT2076",), ("J2:J2070",), ("J1:J2076",),
                       ("J2:J2076", "J2000:J2076")):
            with self.subTest(ranges=ranges), self.assertRaisesRegex(RuntimeError, "must cover"):
                validate_progress_bar(self.saved_sheet(ranges), 10, 2075)

    def test_reject_auto_scaled_progress_rule(self):
        with self.assertRaisesRegex(RuntimeError, "fixed 0-100%"):
            validate_progress_bar(self.saved_sheet(points=(("min", None), ("max", None))), 10, 2075)


if __name__ == "__main__":
    unittest.main()
