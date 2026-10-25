import unittest
from datetime import date, datetime
from decimal import Decimal
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from weekly.workbook import (
    _com_value,
    _excel_serial,
    _flatten_range_values,
    _fit_column_to_range,
    _fit_pivot_headers,
    _autofit_preserving_hidden_rows,
    _hidden_pivot_tail_range,
    _latest_report_week,
    _leadtime_item_visible,
    _uniform_relative_formula,
    _apply_nonnegative_leadtime_filters,
    _restore_pivot_helpers,
    _style_pivot_helpers,
    PivotHelperFormula,
    _apply_mobile_week_window,
    _reset_summary_adjustment_filters,
    _refresh_pivots,
    _refresh_slide_summary,
    _week_number,
)


class WeeklyWorkbookValueTests(unittest.TestCase):
    def test_mobile_slide_summary_retains_pivot_precision_instead_of_rounding_values(self) -> None:
        cells = {}
        ranges = {}
        copies = []
        def cell(row, column):
            return cells.setdefault((row, column), SimpleNamespace(Value=None, Value2=None))
        def value(row, column, content):
            cell(row, column).Value = content
            cell(row, column).Value2 = content
        value(8, 13, "Row Labels")
        value(8, 14, "Bulk Service")
        value(8, 15, "Mobile")
        value(9, 20, "Bulk Service")
        value(9, 21, "Mobile")
        value(10, 19, "Jan-26")
        value(18, 19, "W1-Sep")
        value(22, 19, "AVG Total")
        def cell_range(first, last):
            key = next(key for key, item in cells.items() if item is first)
            end = next(key for key, item in cells.items() if item is last)
            if (key, end) not in ranges:
                region = MagicMock()
                region.Copy = lambda destination: copies.append((key, next(k for k, item in cells.items() if item is destination)))
                ranges[key, end] = region
            return ranges[key, end]
        for row, label, bulk, mobile in (
            (9, "Jan-26", 1.25, 1.75), (10, "Sep-26", 2.6, 1.8),
            (11, "W36", 2.5714, 1.3378), (12, "W37", 8.4285, 1.8536),
            (13, "W38", 2.0731, 1.1481), (14, "W39", 1.6666, 2.4875),
            (15, "Grand Total", 2.93809, 2.40063),
        ):
            value(row, 13, label)
            value(row, 14, bulk)
            value(row, 15, mobile)
        month = SimpleNamespace(Name="Sep-26", Visible=True, ShowDetail=True)
        class Collection:
            Count = 1
            def __init__(self, item):
                self.item = item
            def __call__(self, _):
                return self.item
        pivot = SimpleNamespace(ManualUpdate=False, RefreshTable=lambda: None,
            PivotFields=lambda name: SimpleNamespace(Orientation=1, PivotItems=lambda: Collection(month)),
            TableRange2=SimpleNamespace(Row=1, Column=13,
                Rows=SimpleNamespace(Count=15), Columns=SimpleNamespace(Count=4)))
        class Series:
            Count = 2
            def __call__(self, _):
                return SimpleNamespace()
        chart = SimpleNamespace(Chart=SimpleNamespace(SeriesCollection=lambda: Series()))
        sheet = SimpleNamespace(Cells=cell, Range=cell_range,
            PivotTables=lambda index=None: Collection(pivot) if index is None else pivot,
            ChartObjects=lambda: Collection(chart))
        with patch("weekly.workbook._sheet_by_name", return_value=sheet):
            last = _refresh_slide_summary(SimpleNamespace(), report_year=2026, report_week=39,
                                          mobile=True, reviewed_values={})
        self.assertEqual(last, 14)
        self.assertEqual(cell(last, 20).Value, 1.6666)
        self.assertEqual(cell(last, 21).Value, 2.4875)
        self.assertEqual(cell(last + 1, 20).Value, 2.93809)
        self.assertEqual(cell(last + 1, 21).Value, 2.40063)
        self.assertIn(((18, 19), (last, 19)), copies)
        self.assertTrue(ranges[((last + 1, 19), (last + 1, 22))].Font.Bold)

    @staticmethod
    def _styled_cell(row, column, *, fill=16777215, pattern=-4142, bold=False):
        borders = {edge: SimpleNamespace(LineStyle=-4142, Color=0, Weight=2) for edge in (8, 9)}
        source_borders = {edge: SimpleNamespace(LineStyle=-4142, Color=0, Weight=2) for edge in (8, 9)}
        return SimpleNamespace(
            Row=row, Column=column,
            Interior=SimpleNamespace(Color=fill, Pattern=pattern),
            Font=SimpleNamespace(Bold=bold),
            Borders=borders.__getitem__,
            DisplayFormat=SimpleNamespace(
                Interior=SimpleNamespace(Color=fill, Pattern=pattern),
                Font=SimpleNamespace(Bold=bold),
                Borders=source_borders.__getitem__,
            ),
        )

    def test_summaries_include_all_adjustments_without_granting_pm_approval(self) -> None:
        field = SimpleNamespace(AllItemsVisible=False, Orientation=3)
        field.ClearAllFilters = lambda: setattr(field, "AllItemsVisible", True)
        pivot = SimpleNamespace(Name="PivotTable2", PivotFields=lambda _: field)

        class Pivots:
            Count = 1

            def __call__(self, _):
                return pivot

        sheet = SimpleNamespace(Name="PIVOT PERFORMANCE", PivotTables=lambda: Pivots())
        workbook = SimpleNamespace(Worksheets=[sheet])
        self.assertEqual(_reset_summary_adjustment_filters(workbook),
                         {("PIVOT PERFORMANCE", "PivotTable2")})
        self.assertTrue(field.AllItemsVisible)
        self.assertEqual(_reset_summary_adjustment_filters(workbook),set())
        field.AllItemsVisible = False

        def fail():
            raise RuntimeError("Excel refused the filter")

        field.ClearAllFilters = fail
        with self.assertRaisesRegex(RuntimeError, "Cannot reset adjustment filter"):
            _reset_summary_adjustment_filters(workbook)

    def test_adjustment_policy_discovers_new_pivots_but_preserves_detail_filters(self) -> None:
        for name in ("PIVOT PERFORMANCE", "SLIDE GH MEETING", "LEADTIME NOT ACHIEVED"):
            with self.subTest(sheet=name):
                field = SimpleNamespace(AllItemsVisible=False, Orientation=3)
                field.ClearAllFilters = lambda: setattr(field, "AllItemsVisible", True)
                pivot = SimpleNamespace(Name="NewSummaryName", PivotFields=lambda _: field)
                class Collection:
                    Count = 1
                    def __call__(self, _):
                        return pivot
                sheet = SimpleNamespace(Name=name, PivotTables=lambda: Collection())
                changed = _reset_summary_adjustment_filters(SimpleNamespace(Worksheets=[sheet]))
                is_summary = name == "PIVOT PERFORMANCE"
                self.assertEqual(field.AllItemsVisible, is_summary)
                self.assertEqual(bool(changed), is_summary)

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

    def test_filter_policy_runs_for_both_roles_each_week_before_explicit_review(self) -> None:
        for mobile in (False, True):
            for week in (9, 10, 11):
                with self.subTest(mobile=mobile, week=week):
                    events = []
                    workbook = SimpleNamespace(Worksheets=[])
                    def stage(name):
                        def apply(*args):
                            events.append(name)
                            return set()
                        return apply
                    with patch("weekly.workbook._reset_summary_adjustment_filters", side_effect=stage("adjustment")), \
                         patch("weekly.workbook._apply_nonnegative_leadtime_filters", side_effect=stage("leadtime")), \
                         patch("weekly.workbook._apply_mobile_week_window", return_value=set()), \
                         patch("weekly.workbook._order_week_items", return_value=set()), \
                         patch("weekly.workbook._apply_pivot_review", side_effect=stage("review")), \
                         patch("weekly.workbook._refresh_selected_pivots"):
                        _refresh_pivots(workbook, mobile=mobile, report_week=week, pivot_review=[])
                    self.assertEqual(events, ["adjustment", "leadtime", "review"])

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
            SimpleNamespace(Name="0", Visible=False, RecordCount=2),
            SimpleNamespace(Name="57", Visible=False, RecordCount=4),
            SimpleNamespace(Name="103", Visible=False, RecordCount=0),
            SimpleNamespace(Name="-2", Visible=True, RecordCount=1),
            SimpleNamespace(Name="(blank)", Visible=True, RecordCount=2),
        ])
        field = SimpleNamespace(Orientation=3, EnableMultiplePageItems=False, PivotItems=lambda: items)
        pivot = SimpleNamespace(Name="Average", ManualUpdate=False, PivotFields=lambda _: field)
        pivots = Items([pivot])
        sheet = SimpleNamespace(Name="PIVOT PERFORMANCE", PivotTables=lambda index=None: pivots if index is None else pivot)
        workbook = SimpleNamespace(Worksheets=[sheet])
        changed = _apply_nonnegative_leadtime_filters(workbook)
        self.assertEqual(changed, {("PIVOT PERFORMANCE", "Average")})
        self.assertEqual([i.Visible for i in items.values], [True, True, True, False, False, False])
        self.assertTrue(field.EnableMultiplePageItems)
        self.assertFalse(pivot.ManualUpdate)
        self.assertEqual(_apply_nonnegative_leadtime_filters(workbook), set())
        items.values[1].Visible = False
        sheet.Name = "LEADTIME NOT ACHIEVED"
        self.assertEqual(_apply_nonnegative_leadtime_filters(workbook), set())
        self.assertFalse(items.values[1].Visible)

    def test_helper_formula_extends_through_refreshed_grand_total(self) -> None:
        area = SimpleNamespace(Row=20, Rows=SimpleNamespace(Count=12),
                               Column=1, Columns=SimpleNamespace(Count=4))
        pivot = SimpleNamespace(TableRange1=area, DataBodyRange=SimpleNamespace(Row=22))
        result = SimpleNamespace()
        sheet = SimpleNamespace(Name="Summary", PivotTables=lambda _: pivot,
                                Cells=self._styled_cell)
        calls = []

        def cell_range(first, last):
            calls.append(((first.Row, first.Column), (last.Row, last.Column)))
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
        cleared_formats = []
        cells = {}
        for row in range(20, 29):
            cells[row] = self._styled_cell(row, 5)
            cells[row].FormulaR1C1 = "Note" if row == 27 else "=RC[-3]/RC[-1]"
            cells[row].ClearContents = lambda row=row: cleared.append(row)
            cells[row].ClearFormats = lambda row=row: cleared_formats.append(row)
        sheet = SimpleNamespace(Name="Summary", PivotTables=lambda _: pivot,
                                Cells=lambda row, column: cells[row] if column == 5
                                else self._styled_cell(row, column), Range=lambda *_: SimpleNamespace())
        _restore_pivot_helpers(SimpleNamespace(Worksheets=[sheet]), [
            PivotHelperFormula("Summary", "Average", 5, "=RC[-3]/RC[-1]", "0.0%", 22, 28)
        ])
        self.assertEqual(cleared, [25, 26, 28])
        self.assertEqual(cleared_formats, cleared)

    def test_helper_styles_follow_current_pivot_rows_when_footer_moves(self) -> None:
        for new_last_row in (25, 28, 31):
            with self.subTest(last_row=new_last_row):
                area = SimpleNamespace(Row=20, Rows=SimpleNamespace(Count=new_last_row - 19),
                                       Column=1, Columns=SimpleNamespace(Count=4))
                pivot = SimpleNamespace(TableRange1=area, DataBodyRange=SimpleNamespace(Row=22))
                cells = {}
                for row in range(22, max(28, new_last_row) + 1):
                    for column in (4, 5):
                        total = row == (new_last_row if column == 4 else 28)
                        cell = self._styled_cell(
                            row, column, fill=15917529 if total else 16777215,
                            pattern=1 if total else -4142, bold=total,
                        )
                        cell.FormulaR1C1 = "=RC[-3]/RC[-1]" if column == 5 else None
                        if total:
                            border = (cell.DisplayFormat if column == 4 else cell).Borders(8)
                            border.LineStyle = 1
                            border.Color = 123456
                        cell.ClearContents = lambda: None
                        cell.ClearFormats = lambda: None
                        cells[row, column] = cell
                sheet = SimpleNamespace(Name="Summary", PivotTables=lambda _: pivot,
                    Cells=lambda row, column: cells[row, column], Range=lambda *_: SimpleNamespace())
                helpers = [
                    PivotHelperFormula("Summary", "Average", 5, "=RC[-3]/RC[-1]", "0.00%", 22, 28)
                ]
                workbook = SimpleNamespace(Worksheets=[sheet])
                _restore_pivot_helpers(workbook, helpers)
                _style_pivot_helpers(workbook, helpers)
                for row in range(22, new_last_row + 1):
                    source, target = cells[row, 4], cells[row, 5]
                    self.assertEqual(target.Interior.Color, source.DisplayFormat.Interior.Color)
                    self.assertEqual(target.Interior.Pattern, source.DisplayFormat.Interior.Pattern)
                    self.assertEqual(target.Font.Bold, row == new_last_row)
                    self.assertEqual(target.Borders(8).LineStyle, 1 if row == new_last_row else -4142)
                    if row == new_last_row:
                        self.assertEqual(target.Borders(8).Color, 123456)

    def test_helper_styles_preserve_subtotal_bands_and_percentage_format(self) -> None:
        area = SimpleNamespace(Row=20, Rows=SimpleNamespace(Count=5),
                               Column=1, Columns=SimpleNamespace(Count=4))
        pivot = SimpleNamespace(TableRange1=area, DataBodyRange=SimpleNamespace(Row=22))
        cells = {}
        for row in range(22, 25):
            cells[row, 4] = self._styled_cell(row, 4, fill=123456 if row == 23 else 16777215,
                                             pattern=1 if row == 23 else -4142, bold=row == 23)
            cells[row, 5] = self._styled_cell(row, 5)
        result = SimpleNamespace()
        sheet = SimpleNamespace(Name="Summary", PivotTables=lambda _: pivot,
            Cells=lambda row, column: cells[row, column], Range=lambda *_: result)
        helpers = [
            PivotHelperFormula("Summary", "Average", 5, "=RC[-3]/RC[-1]", "0.00%")
        ]
        workbook = SimpleNamespace(Worksheets=[sheet])
        _restore_pivot_helpers(workbook, helpers)
        _style_pivot_helpers(workbook, helpers)
        self.assertEqual(cells[23, 5].Interior.Color, 123456)
        self.assertTrue(cells[23, 5].Font.Bold)
        self.assertFalse(cells[24, 5].Font.Bold)
        self.assertEqual(result.NumberFormat, "0.00%")

    def test_helper_styles_follow_revealed_total_borders_without_rewriting_formulas(self) -> None:
        area = SimpleNamespace(Row=20, Rows=SimpleNamespace(Count=5),
                               Column=1, Columns=SimpleNamespace(Count=4))
        pivot = SimpleNamespace(TableRange1=area, DataBodyRange=SimpleNamespace(Row=22))
        cells = {(row, column): self._styled_cell(row, column)
                 for row in range(22, 25) for column in (4, 5)}
        result = SimpleNamespace()
        sheet = SimpleNamespace(Name="Summary", PivotTables=lambda _: pivot,
            Cells=lambda row, column: cells[row, column], Range=lambda *_: result)
        helpers = [PivotHelperFormula("Summary", "Average", 5, "=RC[-3]/RC[-1]", "0.00%")]
        workbook = SimpleNamespace(Worksheets=[sheet])
        _restore_pivot_helpers(workbook, helpers)

        cells[23, 4].DisplayFormat.Borders(9).LineStyle = 1
        cells[24, 4].DisplayFormat.Borders(8).LineStyle = 1
        cells[24, 4].DisplayFormat.Font.Bold = True
        cells[24, 4].DisplayFormat.Interior.Color = 15917529
        cells[24, 4].DisplayFormat.Interior.Pattern = 1
        sheet.Range = lambda *_: self.fail("Styling must not rewrite helper formulas")
        _style_pivot_helpers(workbook, helpers)

        self.assertEqual(cells[23, 5].Borders(9).LineStyle, 1)
        self.assertEqual(cells[24, 5].Borders(8).LineStyle, 1)
        self.assertTrue(cells[24, 5].Font.Bold)
        self.assertEqual(cells[24, 5].Interior.Color, 15917529)
        self.assertEqual(result.FormulaR1C1, "=RC[-3]/RC[-1]")
        self.assertEqual(result.NumberFormat, "0.00%")

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
        self.assertTrue(_fit_column_to_range(sheet, Range(70), limit=32))

    def test_long_headers_wrap_without_changing_data_or_shrinking_existing_columns(self) -> None:
        cells = {}
        def get_range(first, last):
            key = (first, last)
            cells.setdefault(key, SimpleNamespace(WrapText=False,
                Rows=SimpleNamespace(AutoFit=lambda: None)))
            return cells[key]
        sheet = SimpleNamespace(Cells=lambda row, col: (row, col), Range=get_range)
        pivot = SimpleNamespace(TableRange1=SimpleNamespace(Row=8, Column=1,
            Columns=SimpleNamespace(Count=2)), DataBodyRange=SimpleNamespace(Row=10))
        with patch("weekly.workbook._fit_column_to_range", side_effect=[False, True]) as fit, \
             patch("weekly.workbook._autofit_preserving_hidden_rows") as heights:
            _fit_pivot_headers(sheet, pivot)
        self.assertEqual(fit.call_count, 2)
        heights.assert_called_once()
        self.assertFalse(cells[((8, 1), (9, 1))].WrapText)
        self.assertTrue(cells[((8, 2), (9, 2))].WrapText)

    def test_header_autofit_preserves_intentionally_hidden_rows(self) -> None:
        values = [SimpleNamespace(Hidden=False), SimpleNamespace(Hidden=True)]
        class Rows:
            Count = 2
            def __call__(self, index):
                return values[index - 1]
            def AutoFit(self):
                for row in values:
                    row.Hidden = False
        _autofit_preserving_hidden_rows(SimpleNamespace(Rows=Rows()))
        self.assertEqual([row.Hidden for row in values], [False, True])


if __name__ == "__main__":
    unittest.main()
