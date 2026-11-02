import unittest
from types import SimpleNamespace
from unittest.mock import Mock

from excel_pivot_layout import align_pivot_row_order


class Item:
    def __init__(self, field, name, visible=True):
        self.field, self.Name, self.Visible = field, name, visible

    @property
    def Position(self):
        return [item for item in self.field.items if item.Visible].index(self) + 1 if self.Visible else -1

    @Position.setter
    def Position(self, position):
        visible = [item for item in self.field.items if item.Visible and item is not self]
        visible.insert(position - 1, self)
        self.field.items = visible + [item for item in self.field.items if not item.Visible]


def field(names, source='Division', hidden=()):
    result = SimpleNamespace(SourceName=source, Name=source, AutoSort=Mock())
    result.items = [Item(result, name) for name in names] + [Item(result, name, False) for name in hidden]
    result.PivotItems = lambda: result.items
    return result


class PivotRowOrderTests(unittest.TestCase):
    def apply(self, source, target):
        phase = SimpleNamespace(Name='Phases', RowFields=[target], Update=Mock())
        count = align_pivot_row_order(SimpleNamespace(RowFields=[source]), phase)
        return count, phase

    def test_reversed_categories_follow_summary_without_fixed_names(self):
        for order in (('Division B', 'Division A'), ('Division A', 'Division B')):
            source, target = field(order), field(reversed(order))
            count, phase = self.apply(source, target)
            self.assertEqual(count, 1)
            self.assertEqual([item.Name for item in target.items], list(order))
            target.AutoSort.assert_called_once_with(-4135, 'Division')
            phase.Update.assert_called_once()
            self.assertEqual([item.Name for item in source.items], list(order))

    def test_hidden_items_remain_hidden_and_new_items_keep_relative_order(self):
        source = field(('A', 'B'), hidden=('Hidden',))
        target = field(('B', 'New Z', 'New Y', 'A'), hidden=('Hidden',))
        self.apply(source, target)
        self.assertEqual([item.Name for item in target.items], ['A', 'B', 'New Z', 'New Y', 'Hidden'])
        self.assertFalse(target.items[-1].Visible)

    def test_matching_or_empty_order_does_not_change_sort_or_refresh(self):
        for names in (('A', 'B'), ()):
            source, target = field(names), field(names)
            count, phase = self.apply(source, target)
            self.assertEqual(count, 0)
            target.AutoSort.assert_not_called()
            phase.Update.assert_not_called()

    def test_field_captions_are_matched_without_changing_row_hierarchy(self):
        source, target = field(('A', 'B'), source='Service Delivery Div. '), field(('B', 'A'), source='Service Delivery Div.')
        self.assertEqual(self.apply(source, target)[0], 1)
        other = field(('B', 'A'), source='Different field')
        self.assertEqual(self.apply(source, other)[0], 0)

    def test_distinct_item_whitespace_is_not_merged(self):
        source, target = field(('A', 'A ')), field(('A ', 'A'))
        self.apply(source, target)
        self.assertEqual([item.Name for item in target.items], ['A', 'A '])

    def test_failed_position_update_is_not_silently_accepted(self):
        class StuckItem(Item):
            @Item.Position.setter
            def Position(self, position):
                pass
        source, target = field(('A', 'B')), field(())
        target.items = [StuckItem(target, 'B'), StuckItem(target, 'A')]
        with self.assertRaisesRegex(RuntimeError, 'Could not align row order'):
            self.apply(source, target)


if __name__ == '__main__':
    unittest.main()
