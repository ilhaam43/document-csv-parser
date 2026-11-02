import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd

from csv_to_excel import (
    PHASE_HEADER_PREFIX,
    REPORT_ONE_PIVOT_TARGET_FILTERS,
    PIVOT_TEMPLATE_SERVICE_DELIVERY_DIV_HEADER,
    QUO_HEADER,
    RFS_COMMMIT_HEADER,
    SERVICE_DELIVERY_DIV_HEADER,
    TARGET_SO_COMPLETE_DATE_HEADERS,
    add_pivot_template_compatibility_columns,
    add_aging_of_rfs_column,
    apply_previous_target_date_fallback,
    apply_current_status_phase_overrides,
    apply_current_month_process_adjustment_overrides,
    clean_dataframe,
    completion_denominator_references,
    completion_percentage_formula,
    completion_summary_pivot,
    ensure_percentage_completion_column_widths,
    fill_missing_with_quo_lookup,
    pivot_row_identity,
    prepare_all_order_dataframe_for_template,
    restore_dashboard_pivot_presentation,
    target_header_month_name,
    target_header_period,
    validated_pivot_report_date,
)


class ReportOneBusinessRuleTests(unittest.TestCase):
    def test_current_export_phase_wins_over_inconsistent_completed_history(self) -> None:
        data = pd.DataFrame({'Status': ['SO Complete'], 'Phase': ['07-UAT On Hold']})
        apply_current_status_phase_overrides(data, 'Phase', pd.Series(['SO Complete']),
                                             source_phase=pd.Series(['SO complete']))
        self.assertEqual(data['Phase'].tolist(), ['SO Complete'])

    def test_explicit_current_phase_is_not_replaced_by_a_conflicting_status(self) -> None:
        data = pd.DataFrame({'Status': ['SO Complete', 'Installation'], 'Phase': ['SO Complete', '06-Installation']})
        apply_current_status_phase_overrides(data, 'Phase', source_phase=pd.Series(['New', 'Pre Installation']))
        self.assertEqual(data['Phase'].tolist(), ['00-New', '04-Pre Installation'])

    def test_status_fallback_only_applies_when_source_phase_is_unrecognized_or_missing(self) -> None:
        data = pd.DataFrame({'Status': ['SO Complete', 'Survey', 'Unknown'], 'Phase': ['00-New'] * 3})
        apply_current_status_phase_overrides(data, 'Phase', source_phase=pd.Series([None, 'Unknown', None]))
        self.assertEqual(data['Phase'].tolist(), ['SO Complete', '02-Survey', '00-New'])

    def test_process_adjustment_uses_original_export_not_display_history_or_old_classification(self) -> None:
        data = pd.DataFrame({'Process': ['New Registration'] * 5, 'Process Adjustment': ['New Reg'] * 5})
        original = pd.Series(['New Registration', 'UPGRADE', 'Reconfiguration', None, ''])
        apply_current_month_process_adjustment_overrides(data, date(2029, 3, 2), source_process=original, rebuild_period=True)
        self.assertEqual(data['Process Adjustment'].tolist(), ['New Reg'] + ['Non New Reg'] * 4)
        self.assertEqual(data['Process'].tolist(), ['New Registration'] * 5)

    def test_process_classification_does_not_depend_on_specific_months_or_fab_population(self) -> None:
        for report_date in (date(2027, 9, 2), date(2028, 1, 2)):
            data = pd.DataFrame({'Process': [' NEW REGISTRATION ', 'Relocation'],
                                 'Process Adjustment': ['Non New Reg', 'New Reg']})
            apply_current_month_process_adjustment_overrides(data, report_date, rebuild_period=True)
            self.assertEqual(data['Process Adjustment'].tolist(), ['New Reg', 'Non New Reg'])

    def test_daily_continuation_retains_historical_classifications_and_only_fills_missing(self) -> None:
        data = pd.DataFrame({'Process': ['New Registration', 'Upgrade', 'Upgrade'],
                             'Process Adjustment': ['Non New Reg', 'New Reg', None]})
        apply_current_month_process_adjustment_overrides(data, date(2029, 3, 3))
        self.assertEqual(data['Process Adjustment'].tolist(), ['Non New Reg', 'New Reg', 'Non New Reg'])

    def test_template_period_selects_rebuild_or_daily_continuation_including_year_rollover(self) -> None:
        source = pd.DataFrame({'Process': ['Upgrade']})
        enriched = pd.DataFrame({'Process': ['New Registration'], 'Process Adjustment': ['New Reg'],
                                 'TARGET Determined as 1 January 2029': ['Target January']})
        for previous_period, expected in (((2028, 12), 'Non New Reg'), ((2029, 1), 'New Reg'), (None, 'New Reg')):
            mappings = SimpleNamespace(reference_target_period=previous_period, target_so_complete_date={}, phase_header='Phase')
            with self.subTest(previous_period=previous_period), \
                 patch('csv_to_excel.apply_lookup_values', return_value=enriched.copy()):
                result, _ = prepare_all_order_dataframe_for_template(source, Path('DataOrderSD-20290102-060000.csv'), mappings)
                self.assertEqual(result.loc[0, 'Process Adjustment'], expected)

    def test_weekly_pivot_rules_preserve_the_actual_report_day(self) -> None:
        report_date = date(2026, 8, 28)

        self.assertEqual(
            validated_pivot_report_date(report_date, "TARGET  Detemined as 1 Aug 26"),
            report_date,
        )

    def test_cleaning_preserves_pm_display_key_whitespace(self) -> None:
        options = SimpleNamespace(
            normalize_headers=True,
            keep_empty=True,
            drop_empty_columns=False,
            dedupe=False,
            infer_types=False,
        )
        source = pd.DataFrame(
            {
                "#Quo": ["2-674178064625  (2)"],
                "PM": ["Muhammad Mahdi Ramadhan "],
                "Phase": [" 00-New "],
            }
        )

        result = clean_dataframe(source, options, Path("DataOrderSD-20260901.csv"))

        self.assertEqual(result.loc[0, "PM"], "Muhammad Mahdi Ramadhan ")
        self.assertEqual(result.loc[0, "Quo"], "2-674178064625  (2)")
        self.assertEqual(result.loc[0, "Phase"], "00-New")

    def test_rfs_aging_ignores_pre_2000_sentinel_dates(self) -> None:
        source = pd.DataFrame({RFS_COMMMIT_HEADER: ["1999-11-30", "2026-08-31"]})

        result = add_aging_of_rfs_column(source, date(2026, 9, 1))

        self.assertTrue(pd.isna(result.loc[0, "Aging Of RFS"]))
        self.assertEqual(result.loc[1, "Aging Of RFS"], 1)

    def test_current_pm_wins_and_previous_pm_is_only_a_blank_fallback(self) -> None:
        current = pd.Series(["Current PM", "", pd.NA], dtype="string")
        quote_ids = pd.Series(["Q-1", "Q-2", "Q-3"], dtype="string")

        result = fill_missing_with_quo_lookup(
            current,
            quote_ids,
            {"Q-1": "Old PM", "Q-2": "Fallback PM", "Q-3": "Fallback PM 2"},
        )

        self.assertEqual(result.tolist(), ["Current PM", "Fallback PM", "Fallback PM 2"])

    def test_percentage_column_width_is_enforced_after_pivot_refresh(self) -> None:
        percentage_header = "Percentage of Completion (SO Complete, Cancel & Change Target)"

        class Worksheet:
            def __init__(self) -> None:
                self.widths = {column: SimpleNamespace(ColumnWidth=10.0) for column in range(4, 20)}
                self.UsedRange = SimpleNamespace(Value2=((percentage_header,),), Row=115, Column=10)

            def Cells(self, row: int, column: int):
                value = percentage_header if (row, column) == (114, 10) else None
                return SimpleNamespace(Value=value)

            def Columns(self, column: int):
                return self.widths[column]

        worksheet = Worksheet()
        ensure_percentage_completion_column_widths(worksheet)

        self.assertEqual(worksheet.widths[10].ColumnWidth, 22.0)
        self.assertEqual(worksheet.widths[9].ColumnWidth, 10.0)

    def test_both_new_registration_summary_names_apply_the_target_month(self) -> None:
        self.assertEqual(REPORT_ONE_PIVOT_TARGET_FILTERS['PivotTable1'], 'target')
        self.assertEqual(REPORT_ONE_PIVOT_TARGET_FILTERS['PivotTable2'], 'target')
        self.assertEqual(REPORT_ONE_PIVOT_TARGET_FILTERS['PivotTable3'], 'target')

    def test_completion_formula_groups_the_numerator_and_only_guards_zero_counts(self) -> None:
        self.assertEqual(completion_percentage_formula(['E118', 'G118'], 'C119'),
                         '=IF(C119=0,0,(E118+G118)/C119)')
        self.assertEqual(completion_percentage_formula([], 'C381'), '=IF(C381=0,0,(0)/C381)')

    def test_pivot_row_identity_distinguishes_subtotals_and_uses_field_names(self) -> None:
        class Items:
            def __init__(self, values):
                self.values = values
                self.Count = len(values)

            def __call__(self, index):
                field, label = self.values[index - 1]
                return SimpleNamespace(Name=label, Parent=SimpleNamespace(SourceName=field))

        def identity(items):
            return pivot_row_identity(SimpleNamespace(PivotCell=SimpleNamespace(RowItems=Items(items))))

        detail = [('Service Delivery Div. ', 'Delivery A'), ('Dept SD', 'Team A')]
        self.assertEqual(identity(detail), identity(list(reversed(detail))))
        self.assertNotEqual(identity(detail), identity(detail[:1]))
        self.assertNotEqual(identity(detail), identity([('Service Delivery Div. ', 'Delivery B'), ('Dept SD', 'Team A')]))
        self.assertEqual(identity([]), ())

    def test_denominator_lookup_follows_count_data_field_not_mrc_or_row_offset(self) -> None:
        from unittest.mock import patch
        fields = {1: SimpleNamespace(SourceName='MRC', Function=-4157),
                  2: SimpleNamespace(SourceName='quo', Function=-4112)}
        cells = {(r, c): SimpleNamespace(
            PivotCell=SimpleNamespace(DataField=fields[c]), Row=115 + r, Column=1 + c,
            identity=('detail',) if r == 1 else (),
        ) for r in (1, 2) for c in (1, 2)}
        pivot = SimpleNamespace(Name='Summary', DataBodyRange=SimpleNamespace(
            Cells=lambda r, c: cells[r, c], Rows=SimpleNamespace(Count=2), Columns=SimpleNamespace(Count=2),
        ))
        with patch('excel_pivot_layout.pivot_row_identity', side_effect=lambda cell: cell.identity):
            self.assertEqual(completion_denominator_references(pivot), {('detail',): 'C116', (): 'C117'})
            cells[2, 2].identity = ('detail',)
            with self.assertRaisesRegex(RuntimeError, 'duplicate summary row'):
                completion_denominator_references(pivot)
        fields[2].Function = -4157
        with self.assertRaisesRegex(RuntimeError, 'no quote-count'):
            completion_denominator_references(pivot)

    def test_adjacent_summary_uses_nearest_matching_row_fields(self) -> None:
        def pivot(name, row, col, row_names):
            return SimpleNamespace(Name=name, TableRange1=SimpleNamespace(Column=col),
                                   RowRange=SimpleNamespace(Row=row),
                                   RowFields=[SimpleNamespace(SourceName=n) for n in row_names],
                                   DataFields=[SimpleNamespace(SourceName='quo', Function=-4112)])
        names = ['Dept SD', 'Service Delivery Div. ']
        phase = pivot('phase', 114, 4, names)
        correct = pivot('summary', 115, 1, list(reversed(names)))
        other_section = pivot('old_section', 9, 1, names)
        wrong_grain = pivot('other', 114, 1, ['PM'])
        sheet = SimpleNamespace(PivotTables=lambda: [phase, other_section, wrong_grain, correct])
        self.assertIs(completion_summary_pivot(sheet, phase), correct)
        with self.assertRaisesRegex(RuntimeError, 'no adjacent'):
            completion_summary_pivot(SimpleNamespace(PivotTables=lambda: [phase, wrong_grain]), phase)

    def test_refreshed_pivots_restore_template_widths_and_reveal_body_not_page_filters(self) -> None:
        pivots = [SimpleNamespace(RowRange=SimpleNamespace(EntireRow=SimpleNamespace(Hidden=True)),
                                  PageRange=SimpleNamespace(EntireRow=SimpleNamespace(Hidden=True))) for _ in range(2)]
        columns = {1: SimpleNamespace(ColumnWidth=8), 4: SimpleNamespace(ColumnWidth=10)}
        sheet = SimpleNamespace(PivotTables=lambda: pivots, Columns=columns.__getitem__)
        restore_dashboard_pivot_presentation(sheet, {1: 45, 4: 41})
        self.assertEqual([columns[c].ColumnWidth for c in (1, 4)], [45, 41])
        for pivot in pivots:
            self.assertTrue(pivot.PreserveFormatting)
            self.assertFalse(pivot.HasAutoFormat)
            self.assertFalse(pivot.RowRange.EntireRow.Hidden)
            self.assertTrue(pivot.PageRange.EntireRow.Hidden)

    def test_target_header_month_accepts_short_and_long_year_formats(self) -> None:
        self.assertEqual(target_header_month_name("TARGET  Detemined as 1 Aug 26"), "August")
        self.assertEqual(target_header_month_name("TARGET Determined as 1 August 2026"), "August")

    def test_target_header_period_normalizes_short_and_long_year_formats(self) -> None:
        self.assertEqual(target_header_period("TARGET  Detemined as 1 Aug 26"), (2026, 8))
        self.assertEqual(target_header_period("TARGET Determined as 1 September 2026"), (2026, 9))

    def test_implausible_future_target_uses_valid_previous_date_and_bucket(self) -> None:
        target_header = "TARGET  Detemined as 1 Aug 26"
        target_date_header = TARGET_SO_COMPLETE_DATE_HEADERS[0]
        data = pd.DataFrame(
            {
                QUO_HEADER: ["2-803768807413"],
                target_header: ["After August"],
                target_date_header: ["2046-05-28"],
            }
        )
        mappings = SimpleNamespace(target_so_complete_date={"2-803768807413": "2026-05-28"})

        result = apply_previous_target_date_fallback(data, mappings, date(2026, 8, 28))

        self.assertEqual(result.loc[0, target_date_header], "2026-05-28")
        self.assertEqual(result.loc[0, target_header], "Before August")

    def test_reasonable_future_target_is_not_replaced(self) -> None:
        target_header = "TARGET  Detemined as 1 Aug 26"
        target_date_header = TARGET_SO_COMPLETE_DATE_HEADERS[0]
        data = pd.DataFrame(
            {
                QUO_HEADER: ["Q-1"],
                target_header: ["After August"],
                target_date_header: ["2027-05-28"],
            }
        )
        mappings = SimpleNamespace(target_so_complete_date={"Q-1": "2026-05-28"})

        result = apply_previous_target_date_fallback(data, mappings, date(2026, 8, 28))

        self.assertEqual(result.loc[0, target_date_header], "2027-05-28")
        self.assertEqual(result.loc[0, target_header], "After August")

    def test_service_delivery_compatibility_renames_instead_of_duplicating(self) -> None:
        data = pd.DataFrame(
            {
                QUO_HEADER: ["Q-1"],
                SERVICE_DELIVERY_DIV_HEADER: ["Seradel"],
                "Phase 27 August 2026": ["04-Pre Installation"],
            }
        )

        result, hidden_columns = add_pivot_template_compatibility_columns(
            data,
            Path("DataOrderSD-20260828-060743.csv"),
            "",
            "Phase 27 August 2026",
        )

        self.assertNotIn(SERVICE_DELIVERY_DIV_HEADER, result.columns)
        self.assertIn(PIVOT_TEMPLATE_SERVICE_DELIVERY_DIV_HEADER, result.columns)
        self.assertEqual(result.loc[0, PIVOT_TEMPLATE_SERVICE_DELIVERY_DIV_HEADER], "Seradel")
        self.assertNotIn(PIVOT_TEMPLATE_SERVICE_DELIVERY_DIV_HEADER, hidden_columns)
        self.assertEqual(result.loc[0, PHASE_HEADER_PREFIX], "04-Pre Installation")


if __name__ == "__main__":
    unittest.main()
