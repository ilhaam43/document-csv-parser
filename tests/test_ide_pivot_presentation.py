import unittest
from types import SimpleNamespace
from unittest.mock import Mock

from generate_ide_tracking import (
    capture_pivot_alignments,
    restore_pivot_alignments,
    separate_delivery_pivot_sections,
)


class Collection:
    def __init__(self, values):
        self.values = values
        self.Count = len(values)

    def __call__(self, index):
        return self.values[index - 1]


class IdePivotAlignmentTests(unittest.TestCase):
    def worksheet(self):
        empty_items = Collection([])
        cells = {(9, column): SimpleNamespace(Value2=caption, HorizontalAlignment=-4108, VerticalAlignment=-4108)
                 for column, caption in enumerate(('Div./Dept.', 'MRC', 'Count'), 1)}
        for column, field, function, horizontal in ((2, 'MRC', -4106, 1), (3, 'QUOTE ID', -4112, -4108)):
            cells[10, column] = SimpleNamespace(Value2=12, HorizontalAlignment=horizontal, VerticalAlignment=-4108,
                PivotCell=SimpleNamespace(DataField=SimpleNamespace(SourceName=field, Function=function), ColumnItems=empty_items))
        columns = {i: SimpleNamespace(HorizontalAlignment=1, VerticalAlignment=-4107) for i in (1, 2)}
        body = SimpleNamespace(Row=10, Column=2, Columns=Collection(list(columns.values())),
                               Cells=lambda row, column: cells[9 + row, 1 + column])
        pivot = SimpleNamespace(Name='Example', DataBodyRange=body,
                                TableRange1=SimpleNamespace(Row=9, Column=1, Columns=SimpleNamespace(Count=3)))
        sheet = SimpleNamespace(PivotTables=lambda: [pivot], Cells=lambda row, column: cells[row, column])
        return sheet, cells, columns, pivot

    def test_caption_and_data_alignment_restore_without_touching_row_labels_or_values(self):
        sheet, cells, columns, pivot = self.worksheet()
        models = capture_pivot_alignments(sheet)
        for cell in cells.values():
            cell.HorizontalAlignment = 1
            cell.VerticalAlignment = -4107
        restore_pivot_alignments(sheet, models)
        self.assertEqual(cells[9, 1].HorizontalAlignment, -4108)
        self.assertEqual(cells[9, 1].VerticalAlignment, -4108)
        self.assertEqual(columns[1].HorizontalAlignment, 1)
        self.assertEqual(columns[2].HorizontalAlignment, -4108)
        self.assertEqual(cells[10, 3].Value2, 12)

    def test_new_phase_caption_uses_the_same_data_field_alignment(self):
        sheet, cells, columns, pivot = self.worksheet()
        models = capture_pivot_alignments(sheet)
        cells[9, 3].Value2 = 'New phase'
        cells[9, 3].HorizontalAlignment = 1
        cells[10, 3].PivotCell.ColumnItems = Collection([
            SimpleNamespace(Name='New phase', Parent=SimpleNamespace(SourceName='STATUS DELIVERY'))])
        restore_pivot_alignments(sheet, models)
        self.assertEqual(cells[9, 3].HorizontalAlignment, -4108)
        self.assertEqual(columns[2].HorizontalAlignment, -4108)

    def test_empty_pivot_has_no_alignment_model(self):
        sheet, cells, columns, pivot = self.worksheet()
        pivot.DataBodyRange = None
        self.assertEqual(capture_pivot_alignments(sheet), {})
        restore_pivot_alignments(sheet, {})


class IdeDeliverySpacingTests(unittest.TestCase):
    def worksheet(self, *, right_column=8, left_width=7, occupied=False):
        rows = [[None] * right_column for _ in range(40)]
        rows[0][0] = 'ORDER AGING START FROM PRE-INSTALLATION STATUS MARCH 2029'
        rows[0][right_column - 1] = 'ONTIME DELIVERY'
        rows[-1][0] = 'STATUS ORDER WITH AGING FROM RFS COMMIT DATE MARCH 2029'
        ranges = [SimpleNamespace(Row=44, Column=1, Rows=SimpleNamespace(Count=12), Columns=SimpleNamespace(Count=left_width)),
                  SimpleNamespace(Row=43, Column=right_column, Rows=SimpleNamespace(Count=19), Columns=SimpleNamespace(Count=4))]
        extension, block = Mock(), Mock()
        sheet = SimpleNamespace(UsedRange=SimpleNamespace(Row=41, Column=1, Value2=tuple(tuple(row) for row in rows)),
            PivotTables=lambda: [SimpleNamespace(TableRange2=r) for r in ranges],
            Cells=lambda row, column: (row, column), Range=Mock(side_effect=[extension, block]),
            Application=SimpleNamespace(WorksheetFunction=SimpleNamespace(CountA=lambda _: int(occupied)), CutCopyMode=True))
        return sheet, block

    def test_only_independent_otd_block_moves_with_its_title_and_filters(self):
        sheet, block = self.worksheet()
        self.assertEqual(separate_delivery_pivot_sections(sheet), 2)
        sheet.Range.assert_any_call((41, 8), (61, 11))
        block.Cut.assert_called_once_with(Destination=(41, 10))
        self.assertFalse(sheet.Application.CutCopyMode)

    def test_existing_gap_is_not_changed(self):
        sheet, block = self.worksheet(right_column=10)
        self.assertEqual(separate_delivery_pivot_sections(sheet), 0)
        sheet.Range.assert_not_called()

    def test_gutter_tracks_changed_aging_pivot_width(self):
        sheet, block = self.worksheet(right_column=9, left_width=8)
        self.assertEqual(separate_delivery_pivot_sections(sheet), 2)
        block.Cut.assert_called_once_with(Destination=(41, 11))

    def test_other_destination_content_is_never_overwritten(self):
        sheet, block = self.worksheet(occupied=True)
        with self.assertRaisesRegex(RuntimeError, 'destination contains other content'):
            separate_delivery_pivot_sections(sheet)
        block.Cut.assert_not_called()


if __name__ == '__main__':
    unittest.main()
