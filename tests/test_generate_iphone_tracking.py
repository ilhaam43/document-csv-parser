import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from completion_threshold_rules import TARGET_AFTER_MODE, WEEKLY_MODE
from generate_iphone_tracking import (
    IPHONE_COMPLETION_LAYOUT_SPECS, YearPivotSelection, capture_year_pivot_selection,
    local_pivot_source_table, normalize_iphone_completion_pivot_names, refresh_pivots,
    shift_iphone_lower_sections_for_month_rollover,
)


class IphoneThresholdLayoutTests(unittest.TestCase):
    def test_all_target_block_uses_target_after_thresholds(self) -> None:
        self.assertEqual(
            tuple(mode for _, _, mode in IPHONE_COMPLETION_LAYOUT_SPECS),
            (WEEKLY_MODE, WEEKLY_MODE, TARGET_AFTER_MODE),
        )


class FakeCollection:
    def __init__(self, values):
        self.values = values
        self.Count = len(values)

    def __iter__(self):
        return iter(self.values)

    def __call__(self, index):
        if isinstance(index, str):
            return next(value for value in self.values if value.Name == index)
        return self.values[index - 1]


class IphoneLocalPivotSourceTests(unittest.TestCase):
    def workbook(self, *, expected=7, actual=None, source="'D:/old/file.xlsx'!Table27"):
        table = SimpleNamespace(Name='Table27', ListRows=SimpleNamespace(Count=expected))
        all_table = SimpleNamespace(Name='Table57', ListRows=SimpleNamespace(Count=100))
        cache = Mock(SourceData=source, RecordCount=1, Index=3)
        cache.Refresh.side_effect = lambda: setattr(cache, 'RecordCount', expected if actual is None else actual)
        pivots = [Mock(Name=f'Pivot{i}', CacheIndex=1) for i in (1, 2)]
        for pivot in pivots:
            pivot.PivotCache.return_value = cache
            pivot.RefreshTable.return_value = True
        sheets = [SimpleNamespace(Name='ALL ORDER', ListObjects=lambda _: all_table, PivotTables=lambda: FakeCollection([])),
                  SimpleNamespace(Name='ALL ORDER IPHONE', ListObjects=lambda _: table, PivotTables=lambda: FakeCollection([])),
                  SimpleNamespace(Name='PIVOT', PivotTables=lambda: FakeCollection(pivots))]
        def create(**kwargs):
            cache.SourceData = kwargs['SourceData']
            return cache
        for pivot in pivots:
            pivot.PivotFields.return_value = FakeCollection([])
        return SimpleNamespace(Worksheets=sheets, PivotCaches=lambda: SimpleNamespace(Create=create)), cache, pivots

    def test_each_pivot_is_rebound_to_a_fresh_local_cache_with_complete_coverage(self):
        workbook, cache, pivots = self.workbook(expected=13)
        with patch('generate_iphone_tracking.capture_year_pivot_selection', return_value=None):
            self.assertEqual(refresh_pivots(workbook), 2)
        self.assertEqual(cache.SourceData, 'Table27')
        self.assertEqual(cache.RecordCount, 13)
        self.assertEqual(cache.Refresh.call_count, 2)
        for pivot in pivots:
            pivot.ChangePivotCache.assert_called_once_with(cache)
            pivot.RefreshTable.assert_called_once()
            self.assertTrue(pivot.PreserveFormatting)
            self.assertFalse(pivot.HasAutoFormat)

    def test_new_cache_is_attached_before_excel_initializes_it(self):
        workbook, cache, pivots = self.workbook()
        def initialize():
            pivots[0].ChangePivotCache.assert_called_once_with(cache)
            cache.RecordCount = 7
        cache.Refresh.side_effect = initialize
        with patch('generate_iphone_tracking.capture_year_pivot_selection', return_value=None):
            refresh_pivots(workbook)

    def test_period_field_is_detached_and_its_role_restored_after_cache_change(self):
        workbook, cache, pivots = self.workbook()
        old_fields = [SimpleNamespace(Orientation=1, Position=1), SimpleNamespace(Orientation=3, Position=2)]
        new_fields = [SimpleNamespace(Orientation=0, Position=0), SimpleNamespace(Orientation=0, Position=0)]
        for pivot, old in zip(pivots, old_fields):
            pivot.ChangePivotCache.side_effect = lambda _, field=old: self.assertEqual(field.Orientation, 0)
        with patch('generate_iphone_tracking.capture_year_pivot_selection', return_value=None), \
             patch('generate_iphone_tracking.find_dynamic_target_field', side_effect=old_fields + new_fields):
            refresh_pivots(workbook)
        self.assertEqual([(f.Orientation, f.Position) for f in new_fields], [(1, 1), (3, 2)])

    def test_stale_cache_refresh_error_and_false_table_refresh_fail_closed(self):
        for cause in ('count', 'exception', 'false'):
            workbook, cache, pivots = self.workbook(expected=13, actual=12 if cause == 'count' else None)
            if cause == 'exception':
                cache.Refresh.side_effect = RuntimeError('unavailable')
            if cause == 'false':
                pivots[0].RefreshTable.return_value = False
            with self.subTest(cause=cause), patch('generate_iphone_tracking.capture_year_pivot_selection', return_value=None):
                with self.assertRaises(RuntimeError):
                    refresh_pivots(workbook)

    def test_unknown_table_is_not_silently_redirected_to_iphone_data(self):
        table = object()
        for source in ('Table27', "'D:/archive.xlsx'!Table27", '[Table27]'):
            self.assertIs(local_pivot_source_table(source, {'table27': table}), table)
        with self.assertRaisesRegex(RuntimeError, 'Unsupported'):
            local_pivot_source_table('Table270', {'table27': table})

    def test_all_years_and_explicit_selections_are_restored_separately(self):
        workbook, cache, pivots = self.workbook()
        states = [YearPivotSelection(True, ()), YearPivotSelection(False, ('2027',))]
        field = object()
        with patch('generate_iphone_tracking.capture_year_pivot_selection', side_effect=states), \
             patch('generate_iphone_tracking.find_named_pivot_field', return_value=field), \
             patch('generate_iphone_tracking.set_page_field_to_all') as all_years, \
             patch('generate_iphone_tracking.set_visible_items_exact', return_value=True) as selected:
            refresh_pivots(workbook)
        all_years.assert_called_once_with(field)
        selected.assert_called_once_with(field, ('2027',))

    def test_new_period_resets_year_filter_without_dropping_missing_fab_rows(self):
        workbook, cache, pivots = self.workbook()
        with patch('generate_iphone_tracking.capture_year_pivot_selection', return_value=YearPivotSelection(False, ('2027',))), \
             patch('generate_iphone_tracking.find_named_pivot_field', return_value=object()), \
             patch('generate_iphone_tracking.set_page_field_to_all') as all_years, \
             patch('generate_iphone_tracking.set_visible_items_exact') as selected:
            refresh_pivots(workbook, reset_year_filter=True)
        self.assertEqual(all_years.call_count, 2)
        selected.assert_not_called()

    def test_single_all_page_ignores_unreliable_hidden_item_flags(self):
        field = SimpleNamespace(Orientation=3, EnableMultiplePageItems=False, CurrentPage='(All)',
                                PivotItems=lambda: [SimpleNamespace(Name='2027', Visible=False)])
        with patch('generate_iphone_tracking.find_named_pivot_field', return_value=field):
            self.assertEqual(capture_year_pivot_selection(object()), YearPivotSelection(True, ()))
        field.EnableMultiplePageItems = True
        field.PivotItems = lambda: [SimpleNamespace(Name='2027', Visible=True), SimpleNamespace(Name='1900', Visible=False)]
        with patch('generate_iphone_tracking.find_named_pivot_field', return_value=field):
            self.assertEqual(capture_year_pivot_selection(object()), YearPivotSelection(False, ('2027',)))

    def test_completion_roles_are_identified_by_fields_not_original_pivot_names(self):
        pivots = [SimpleNamespace(Name=name, bucket=bucket,
                                  ColumnFields=[SimpleNamespace(SourceName='Phase')] if phase else [])
                  for name, bucket, phase in (('OldSummary', 'New Reg', False), ('OldPhases', 'New Reg', True),
                                              ('OldOtherSummary', 'Non New Reg', False), ('OldOtherPhases', 'Non New Reg', True))]
        sheet = SimpleNamespace(PivotTables=lambda: FakeCollection(pivots))
        workbook = SimpleNamespace(Worksheets=[SimpleNamespace(Name='PIVOT', PivotTables=sheet.PivotTables)])
        with patch('generate_iphone_tracking.find_named_pivot_field', side_effect=lambda pivot, _: SimpleNamespace(
            Orientation=3, CurrentPage=pivot.bucket,
        )):
            normalize_iphone_completion_pivot_names(workbook)
        self.assertEqual([p.Name for p in pivots], ['PivotTable1', 'PivotTable3', 'PivotTable2', 'PivotTable4'])

    def test_section_growth_uses_existing_spacing_before_inserting_rows(self):
        pivots = {name: SimpleNamespace(TableRange1=SimpleNamespace(Row=50, Rows=SimpleNamespace(Count=10)))
                  for name in ('PivotTable1', 'PivotTable2', 'PivotTable3', 'PivotTable4')}
        sheet = Mock()
        sheet.PivotTables.side_effect = pivots.__getitem__
        with patch('generate_iphone_tracking.find_pivot_title_row', return_value=100):
            self.assertEqual(shift_iphone_lower_sections_for_month_rollover(sheet, 1), (0, None))
            sheet.Rows.assert_not_called()
            pivots['PivotTable2'].TableRange1.Rows.Count = 52
            self.assertEqual(shift_iphone_lower_sections_for_month_rollover(sheet, 1), (3, 100))
            sheet.Rows.assert_called_once_with('100:102')


if __name__ == "__main__":
    unittest.main()
