import unittest
from types import SimpleNamespace

from excel_pivot_layout import align_table_body_columns, capture_pivot_alignments, restore_pivot_alignments


class Collection:
    def __init__(self, values):
        self.values = values
        self.Count = len(values)

    def __iter__(self):
        return iter(self.values)

    def __call__(self, index):
        return self.values[index - 1]


class EmptyPivotAlignmentTests(unittest.TestCase):
    def worksheet(self, data_alignment=-4108):
        cells = {
            (8, 5): SimpleNamespace(Value2='Row Labels', HorizontalAlignment=-4108, VerticalAlignment=-4108),
            (8, 6): SimpleNamespace(Value2='(blank)', HorizontalAlignment=-4108, VerticalAlignment=-4108),
            (9, 5): SimpleNamespace(Value2='Grand Total', HorizontalAlignment=1, VerticalAlignment=-4107),
            (9, 6): SimpleNamespace(Value2=None, HorizontalAlignment=data_alignment, VerticalAlignment=-4108),
        }
        field = SimpleNamespace(SourceName='quo', Function=-4112)
        pivot = SimpleNamespace(Name='Empty', DataBodyRange=None, DataFields=Collection([field]),
            TableRange1=SimpleNamespace(Row=8, Column=5, Rows=SimpleNamespace(Count=2), Columns=SimpleNamespace(Count=2)))
        sheet = SimpleNamespace(PivotTables=lambda: [pivot], Cells=lambda row, col: cells[row, col])
        return sheet, pivot, cells, field

    def test_empty_phase_template_supplies_alignment_for_new_categories(self):
        sheet, pivot, cells, field = self.worksheet()
        models = capture_pivot_alignments(sheet)
        cells[8, 6].Value2 = 'New Phase'
        cells[8, 6].HorizontalAlignment = 1
        column = SimpleNamespace(HorizontalAlignment=1, VerticalAlignment=-4107)
        cell = SimpleNamespace(PivotCell=SimpleNamespace(DataField=field, ColumnItems=Collection([
            SimpleNamespace(Name='New Phase', Parent=SimpleNamespace(SourceName='Phase'))])))
        pivot.DataBodyRange = SimpleNamespace(Row=9, Column=6, Columns=Collection([column]), Cells=lambda *_: cell)
        restore_pivot_alignments(sheet, models)
        self.assertEqual(cells[8, 6].HorizontalAlignment, -4108)
        self.assertEqual(column.HorizontalAlignment, -4108)

    def test_empty_summary_retains_general_alignment_not_forced_center(self):
        sheet, pivot, cells, field = self.worksheet(data_alignment=1)
        model = capture_pivot_alignments(sheet)['Empty']
        self.assertEqual(model.data_defaults[('quo', -4112)][0], 1)

    def test_empty_total_label_is_not_treated_as_a_column_header(self):
        sheet, pivot, cells, field = self.worksheet()
        model = capture_pivot_alignments(sheet)
        restore_pivot_alignments(sheet, model)
        self.assertEqual(cells[9, 5].HorizontalAlignment, 1)
        self.assertEqual(cells[9, 5].VerticalAlignment, -4107)

    def test_excel_blank_total_without_datafield_is_a_template_not_a_data_error(self):
        sheet, pivot, cells, field = self.worksheet()
        class EmptyPivotCell:
            @property
            def DataField(self):
                raise RuntimeError('Excel blank pivot cell has no DataField')
        blank = SimpleNamespace(Value2=None, PivotCell=EmptyPivotCell(), HorizontalAlignment=-4108, VerticalAlignment=-4108)
        column = SimpleNamespace(HorizontalAlignment=1, VerticalAlignment=-4107)
        pivot.DataBodyRange = SimpleNamespace(Row=9, Column=6, Columns=Collection([column]), Cells=lambda *_: blank)
        models = capture_pivot_alignments(sheet)
        restore_pivot_alignments(sheet, models)
        self.assertEqual(column.HorizontalAlignment, -4108)
        blank.Value2 = 12
        with self.assertRaisesRegex(RuntimeError, 'no DataField'):
            capture_pivot_alignments(sheet)

    def test_new_pipeline_header_preserves_native_horizontal_and_template_vertical_alignment(self):
        sheet, pivot, cells, field = self.worksheet()
        model = capture_pivot_alignments(sheet)['Empty']
        model.empty_header_positions.clear()
        cells[8, 6].Value2 = 'New Phase'
        cells[8, 6].HorizontalAlignment = 1
        cells[8, 6].VerticalAlignment = -4107
        restore_pivot_alignments(sheet, {'Empty': model}, preserve_new_header_horizontal=True)
        self.assertEqual(cells[8, 6].HorizontalAlignment, 1)
        self.assertEqual(cells[8, 6].VerticalAlignment, -4108)

    def test_preformatted_empty_header_slot_is_restored_when_a_category_appears(self):
        sheet, pivot, cells, field = self.worksheet()
        model = capture_pivot_alignments(sheet)['Empty']
        cells[8, 6].Value2 = 'New Phase'
        cells[8, 6].HorizontalAlignment = 1
        restore_pivot_alignments(sheet, {'Empty': model}, preserve_new_header_horizontal=True)
        self.assertEqual(cells[8, 6].HorizontalAlignment, -4108)

    def test_rebuilt_row_header_keeps_native_alignment_without_changing_numeric_header(self):
        sheet, pivot, cells, field = self.worksheet()
        model = capture_pivot_alignments(sheet)
        pivot.RowRange = SimpleNamespace(Row=8, Column=5, Columns=SimpleNamespace(Count=1))
        cells[8, 5].HorizontalAlignment = 1
        cells[8, 5].VerticalAlignment = -4107
        cells[8, 6].HorizontalAlignment = 1
        restore_pivot_alignments(sheet, model, native_row_headers={'Empty'})
        self.assertEqual(cells[8, 5].HorizontalAlignment, 1)
        self.assertEqual(cells[8, 5].VerticalAlignment, -4107)
        self.assertEqual(cells[8, 6].HorizontalAlignment, -4108)

    def test_categories_beyond_an_empty_pivot_inherit_measure_style_not_unused_cells(self):
        sheet, pivot, cells, field = self.worksheet()
        model = capture_pivot_alignments(sheet)
        model['Empty'].empty_header_positions[1] = (1, -4107)
        cells[8, 6].Value2 = 'First phase'
        cells[8, 7] = SimpleNamespace(Value2='Second phase', HorizontalAlignment=1, VerticalAlignment=-4107)
        cells[9, 7] = SimpleNamespace(Value2=None, HorizontalAlignment=1, VerticalAlignment=-4107)
        pivot.TableRange1.Columns.Count = 3
        restore_pivot_alignments(sheet, model, preserve_new_header_horizontal=True)
        self.assertEqual(cells[8, 6].HorizontalAlignment, 1)
        self.assertEqual(cells[8, 7].HorizontalAlignment, -4108)
        self.assertEqual(cells[8, 7].VerticalAlignment, -4108)

    def test_reference_header_policy_handles_empty_slots_and_new_categories(self):
        sheet, pivot, cells, field = self.worksheet()
        model = capture_pivot_alignments(sheet)
        model['Empty'].empty_header_positions[1] = (1, -4107)
        cells[8, 6].Value2 = 'First phase'
        cells[8, 7] = SimpleNamespace(Value2='Second phase', HorizontalAlignment=1, VerticalAlignment=-4107)
        cells[9, 7] = SimpleNamespace(Value2=None, HorizontalAlignment=1, VerticalAlignment=-4107)
        pivot.TableRange1.Columns.Count = 3
        restore_pivot_alignments(sheet, model)
        self.assertEqual(cells[8, 6].HorizontalAlignment, 1)
        self.assertEqual(cells[8, 7].HorizontalAlignment, -4108)

    def test_populated_reference_header_style_applies_to_new_categories(self):
        sheet, pivot, cells, field = self.worksheet()
        model = capture_pivot_alignments(sheet)['Empty']
        model.empty_header_positions.clear()
        cells[8, 6].Value2 = 'New Phase'
        cells[8, 6].HorizontalAlignment = 1
        restore_pivot_alignments(sheet, {'Empty': model})
        self.assertEqual(cells[8, 6].HorizontalAlignment, -4108)


class TableBodyAlignmentTests(unittest.TestCase):
    def worksheet(self, headers):
        columns = [SimpleNamespace(HorizontalAlignment=1, VerticalAlignment=-4107, Value2='unchanged') for _ in headers]
        table = SimpleNamespace(DataBodyRange=SimpleNamespace(Columns=Collection(columns)),
                                HeaderRowRange=SimpleNamespace(Value2=(tuple(headers),)))
        sheet = SimpleNamespace(ListObjects=Collection([table]))
        return sheet, columns, table

    def test_body_rules_follow_headers_and_preserve_other_columns_and_values(self):
        sheet, columns, table = self.worksheet(('Quo', 'Survey Date', 'Other', 'RFS Commmit', 'STATUS ORDER', 'Aging Of RFS'))
        self.assertEqual(align_table_body_columns(sheet), 4)
        for column in (1, 3, 4, 5):
            self.assertEqual(columns[column].HorizontalAlignment, -4108)
            self.assertEqual(columns[column].VerticalAlignment, -4108)
        self.assertEqual(columns[0].HorizontalAlignment, 1)
        self.assertEqual(columns[2].HorizontalAlignment, 1)
        self.assertTrue(all(column.Value2 == 'unchanged' for column in columns))

    def test_explicit_general_rule_and_normalized_header_names(self):
        sheet, columns, table = self.worksheet((' PRE-INSTALLATION START DATE ', 'survey_date'))
        self.assertEqual(align_table_body_columns(sheet, {'Pre-Installation Start Date': (1, -4107), 'Survey Date': (-4108, -4108)}), 2)
        self.assertEqual(columns[0].HorizontalAlignment, 1)
        self.assertEqual(columns[1].HorizontalAlignment, -4108)

    def test_missing_table_empty_body_and_empty_policy_are_noops(self):
        self.assertEqual(align_table_body_columns(SimpleNamespace(ListObjects=Collection([]))), 0)
        sheet, columns, table = self.worksheet(('Survey Date',))
        self.assertEqual(align_table_body_columns(sheet, {}), 0)
        table.DataBodyRange = None
        self.assertEqual(align_table_body_columns(sheet), 0)


if __name__ == '__main__':
    unittest.main()
