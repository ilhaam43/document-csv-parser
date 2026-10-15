import unittest
from datetime import date, datetime
from decimal import Decimal
from types import SimpleNamespace

from weekly.workbook import (
    _changed_adjustments,
    _com_value,
    _excel_serial,
    _flatten_range_values,
    _fit_column_to_range,
    _hidden_pivot_tail_range,
    _latest_report_week,
    _leadtime_item_visible,
    _uniform_relative_formula,
    _apply_nonnegative_leadtime_filters,
    _restore_pivot_helpers,
    PivotHelperFormula,
    _master_month_rollover,
    _apply_mobile_week_window,
    _include_new_adjustments,
    _week_number,
)


class WeeklyWorkbookValueTests(unittest.TestCase):
    def test_mobile_includes_historical_adjustments_without_granting_pm_approval(self) -> None:
        field = SimpleNamespace(AllItemsVisible=False)
        field.ClearAllFilters = lambda: setattr(field, "AllItemsVisible", True)
        pivot = SimpleNamespace(Name="PivotTable2", PivotFields=lambda _: field)

        class Pivots:
            Count = 1

            def __call__(self, _):
                return pivot

        sheet = SimpleNamespace(Name="PIVOT PERFORMANCE", PivotTables=lambda: Pivots())
        workbook = SimpleNamespace(Worksheets=[sheet])
        self.assertEqual(_include_new_adjustments(workbook,set(),mobile=True),
                         {("PIVOT PERFORMANCE", "PivotTable2")})
        self.assertTrue(field.AllItemsVisible)
        self.assertEqual(_include_new_adjustments(workbook,set(),mobile=True),set())
        field.AllItemsVisible = False

        def fail():
            raise RuntimeError("Excel refused the filter")

        field.ClearAllFilters = fail
        with self.assertRaisesRegex(RuntimeError, "Cannot reset Mobile adjustment filter"):
            _include_new_adjustments(workbook,set(),mobile=True)

    def test_four_week_window_excludes_both_older_and_future_weeks(self) -> None:
        class Items:
            def __init__(self, values):
                self.values = values
                self.Count = len(values)

            def __call__(self, number):
                return self.values[number - 1]

        items = Items([SimpleNamespace(Name=f"W{week}", Visible=True) for week in range(35, 42)])
        field = SimpleNamespace(Orientation=1, PivotItems=lambda: items)
        pivot = SimpleNamespace(Name="Weekly", PivotFields=lambda _: field)
        pivots = Items([pivot])
        sheet = SimpleNamespace(Name="PIVOT PERFORMANCE",
            PivotTables=lambda index=None: pivots if index is None else pivot)
        changed = _apply_mobile_week_window(SimpleNamespace(Worksheets=[sheet]), 39)
        self.assertEqual(changed, {("PIVOT PERFORMANCE", "Weekly")})
        self.assertEqual([i.Name for i in items.values if i.Visible], ["W36", "W37", "W38", "W39"])

    def test_master_filter_rollover_uses_reporting_week_not_bq_export_month(self) -> None:
        field = SimpleNamespace(Orientation=1)
        area = SimpleNamespace(Row=8, Column=13, Rows=SimpleNamespace(Count=4))
        pivot = SimpleNamespace(PivotFields=lambda _: field, TableRange1=area)
        sheet = SimpleNamespace(Name="SLIDE GH MEETING",
            PivotTables=lambda index=None: SimpleNamespace(Count=1) if index is None else pivot,
            Cells=lambda row, column: (row,column),
            Range=lambda *_: SimpleNamespace(Value2=(("Row Labels",),("Sep-26",),("W39",),("Grand Total",))))
        workbook = SimpleNamespace(Worksheets=[sheet])
        self.assertFalse(_master_month_rollover(workbook, 2026, 39))
        self.assertTrue(_master_month_rollover(workbook, 2026, 40))

    def test_leadtime_filter_includes_previously_hidden_valid_values(self) -> None:
        for value in (0, 1, 57, 58, 103, "102.91666"):
            self.assertTrue(_leadtime_item_visible(value))
        for value in (None, "(blank)", "", "NaN", "Infinity", -1, "-20"):
            self.assertFalse(_leadtime_item_visible(value))

    def test_helper_formula_is_relative_and_not_tied_to_footer_row(self) -> None:
        self.assertEqual(_uniform_relative_formula(("=RC[-3]/RC[-1]", None, "=RC[-3]/RC[-1]")),
                         "=RC[-3]/RC[-1]")
        self.assertIsNone(_uniform_relative_formula((None, 0.5)))
        self.assertIsNone(_uniform_relative_formula(("=RC[-3]", "=RC[-1]")))
        self.assertIsNone(_uniform_relative_formula(("=RC[-3]", "Note")))

    def test_numeric_filter_skips_unused_cache_items_and_batches_visibility_changes(self) -> None:
        class Items:
            def __init__(self, values):
                self.values = values
                self.Count = len(values)

            def __call__(self, number):
                return self.values[number - 1]

        items = Items([
            SimpleNamespace(Name="1", Visible=True, RecordCount=3),
            SimpleNamespace(Name="57", Visible=False, RecordCount=4),
            SimpleNamespace(Name="103", Visible=False, RecordCount=0),
            SimpleNamespace(Name="-2", Visible=True, RecordCount=1),
            SimpleNamespace(Name="(blank)", Visible=True, RecordCount=2),
        ])
        field = SimpleNamespace(Orientation=3, EnableMultiplePageItems=True, PivotItems=lambda: items)
        pivot = SimpleNamespace(Name="Average", ManualUpdate=False, PivotFields=lambda _: field)
        pivots = Items([pivot])
        sheet = SimpleNamespace(Name="Summary", PivotTables=lambda index=None: pivots if index is None else pivot)
        workbook = SimpleNamespace(Worksheets=[sheet])
        changed = _apply_nonnegative_leadtime_filters(workbook)
        self.assertEqual(changed, {("Summary", "Average")})
        self.assertEqual([i.Visible for i in items.values], [True, True, False, False, False])
        self.assertFalse(pivot.ManualUpdate)

    def test_helper_formula_extends_through_refreshed_grand_total(self) -> None:
        area = SimpleNamespace(Row=20, Rows=SimpleNamespace(Count=12),
                               Column=1, Columns=SimpleNamespace(Count=4))
        pivot = SimpleNamespace(TableRange1=area, DataBodyRange=SimpleNamespace(Row=22))
        result = SimpleNamespace()
        sheet = SimpleNamespace(Name="Summary", PivotTables=lambda _: pivot,
                                Cells=lambda row, column: (row, column))
        calls = []

        def cell_range(first, last):
            calls.append((first, last))
            return result

        sheet.Range = cell_range
        _restore_pivot_helpers(SimpleNamespace(Worksheets=[sheet]), [
            PivotHelperFormula("Summary", "Average", 5, "=RC[-3]/RC[-1]", "0.0%")
        ])
        self.assertEqual(calls, [((22, 5), (31, 5))])
        self.assertEqual(result.FormulaR1C1, "=RC[-3]/RC[-1]")
        self.assertEqual(result.NumberFormat, "0.0%")

    def test_helper_formula_removes_old_tail_when_pivot_shrinks_without_removing_notes(self) -> None:
        area = SimpleNamespace(Row=20, Rows=SimpleNamespace(Count=5),
                               Column=1, Columns=SimpleNamespace(Count=4))
        pivot = SimpleNamespace(TableRange1=area, DataBodyRange=SimpleNamespace(Row=22))
        cleared = []
        cells = {}
        for row in range(20, 29):
            cells[row] = SimpleNamespace(FormulaR1C1="Note" if row == 27 else "=RC[-3]/RC[-1]",
                                         ClearContents=lambda row=row: cleared.append(row))
        sheet = SimpleNamespace(Name="Summary", PivotTables=lambda _: pivot,
                                Cells=lambda row, column: cells[row], Range=lambda *_: SimpleNamespace())
        _restore_pivot_helpers(SimpleNamespace(Worksheets=[sheet]), [
            PivotHelperFormula("Summary", "Average", 5, "=RC[-3]/RC[-1]", "0.0%", 22, 28)
        ])
        self.assertEqual(cleared, [25, 26, 28])

    def test_excel_serial_preserves_date_without_timezone_shift(self) -> None:
        self.assertEqual(_excel_serial(date(2026, 9, 18)), 46283.0)
        self.assertAlmostEqual(
            _excel_serial(datetime(2026, 9, 18, 12, 46)),
            46283.53194444445,
        )

    def test_com_value_uses_serial_only_for_date_columns(self) -> None:
        value = datetime(2026, 9, 18, 12, 46)
        self.assertAlmostEqual(_com_value("fab_upload", value), 46283.53194444445)
        self.assertEqual(_com_value("Month of FAB Upload", "September-26"), "September-26")

    def test_com_value_converts_decimals_to_excel_numbers(self) -> None:
        self.assertEqual(_com_value("mrc", Decimal("1250.50")), 1250.5)
        self.assertIsNone(_com_value("cash_dt", None))

    def test_flatten_range_values_handles_com_shapes(self) -> None:
        self.assertEqual(_flatten_range_values((("A",), ("B",))), ("A", "B"))
        self.assertEqual(_flatten_range_values("A"), ("A",))
        self.assertEqual(_flatten_range_values(None), ())

    def test_changed_adjustments_detects_reviewed_existing_and_new_rows(self) -> None:
        records = [
            {"ADJUSTMENT": "Approved A", "ACHIEVEMENT": "ACHIEVED"},
            {"ADJUSTMENT": "Unchanged", "ACHIEVEMENT": "ACHIEVED"},
            {"ADJUSTMENT": "Approved B", "ACHIEVEMENT": "ACHIEVED"},
        ]
        self.assertEqual(
            _changed_adjustments(
                ("Approved A", "Unchanged"),
                ("NOT ACHIEVED", "ACHIEVED"),
                records,
            ),
            {"Approved A", "Approved B"},
        )

    def test_week_number_accepts_only_excel_week_labels(self) -> None:
        self.assertEqual(_week_number("W1"), 1)
        self.assertEqual(_week_number("w53"), 53)
        self.assertIsNone(_week_number("W54"))
        self.assertIsNone(_week_number("(blank)"))

    def test_latest_report_week_uses_latest_fab_upload(self) -> None:
        records = [
            {"fab_upload": datetime(2026, 9, 17)},
            {"fab_upload": datetime(2026, 9, 23, 23, 19)},
            {"fab_upload": None},
        ]
        self.assertEqual(_latest_report_week(records), 38)

    def test_hidden_pivot_tail_reveals_total_and_blank_spacing(self) -> None:
        hidden = set(range(48, 353))
        self.assertEqual(
            _hidden_pivot_tail_range(lambda row: row in hidden, lambda row: True, 8, 48),
            (48, 50),
        )
        hidden = set(range(413, 5355))
        self.assertEqual(
            _hidden_pivot_tail_range(lambda row: row in hidden, lambda row: True, 17, 414),
            (413, 416),
        )

    def test_hidden_pivot_tail_preserves_other_hidden_rows(self) -> None:
        hidden = set(range(3, 7)) | set(range(48, 353))
        self.assertIsNone(
            _hidden_pivot_tail_range(lambda row: row in hidden, lambda row: True, 8, 47)
        )
        self.assertEqual(
            _hidden_pivot_tail_range(
                lambda row: row in hidden, lambda row: row != 49, 8, 48
            ),
            (48, 48),
        )

    def test_column_fit_preserves_existing_width_and_caps_filter_width(self) -> None:
        class Column:
            ColumnWidth = 20.0

        class Range:
            Column = 2

            def __init__(self, fitted_width):
                self.Columns = self
                self.fitted_width = fitted_width

            def AutoFit(self):
                Column.ColumnWidth = self.fitted_width

        class Sheet:
            def Columns(self, number):
                self.last_column = number
                return Column

        sheet = Sheet()
        _fit_column_to_range(sheet, Range(10), padding=3, limit=32)
        self.assertEqual(sheet.last_column, 2)
        self.assertEqual(Column.ColumnWidth, 20)

        _fit_column_to_range(sheet, Range(30), padding=5.5, limit=32)
        self.assertEqual(Column.ColumnWidth, 32)


if __name__ == "__main__":
    unittest.main()
