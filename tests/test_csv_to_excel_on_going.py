import unittest
from datetime import date
from types import SimpleNamespace

from csv_to_excel_on_going import (
    DEFAULT_ON_PROGRESS_PIVOT_STYLE,
    expected_on_progress_sheet_name,
    resolve_on_progress_pivot_style,
    restore_on_progress_filter_caption_layout,
)


class FakePivot:
    def __init__(self, style_name):
        self.TableStyle2 = style_name


class FakeProgressSheet:
    def __init__(self, pivots):
        self._pivots = pivots

    def PivotTables(self, name):
        return self._pivots[name]


class ResolveOnProgressPivotStyleTests(unittest.TestCase):
    def test_uses_first_available_template_pivot_style(self):
        sheet = FakeProgressSheet(
            {
                "missing-style": FakePivot(""),
                "template-pivot": FakePivot("PivotStyleLight16"),
            }
        )

        self.assertEqual(
            resolve_on_progress_pivot_style(sheet, ["missing-style", "template-pivot"]),
            "PivotStyleLight16",
        )

    def test_uses_light_style_fallback_when_template_style_is_unavailable(self):
        sheet = FakeProgressSheet({})

        self.assertEqual(
            resolve_on_progress_pivot_style(sheet, ["missing-pivot"]),
            DEFAULT_ON_PROGRESS_PIVOT_STYLE,
        )


class OnProgressSheetNameTests(unittest.TestCase):
    def test_keeps_full_month_name_when_it_fits_excel_limit(self):
        self.assertEqual(
            expected_on_progress_sheet_name(date(2026, 8, 31)),
            "ALL ORDER ON PROGRESS 31 August",
        )

    def test_abbreviates_month_when_full_name_exceeds_excel_limit(self):
        sheet_name = expected_on_progress_sheet_name(date(2026, 9, 1))

        self.assertEqual(sheet_name, "ALL ORDER ON PROGRESS 1 Sep")
        self.assertLessEqual(len(sheet_name), 31)


class FilterCaptionCleanupTests(unittest.TestCase):
    def setup_sheet(self, start=2, column=4):
        cells = {}
        class Cell:
            def __init__(self, row, col):
                self.Row, self.Column, self.Value, self.HasFormula = row, col, None, False
                self.cleared = False
            def ClearContents(self):
                self.Value, self.cleared = None, True
        def cell(row, col):
            return cells.setdefault((row, col), Cell(row, col))
        fields = []
        for offset, name in enumerate(("Process Adjustment", "Phase 18 November 2026", "YEAR FAB UPLOAD")):
            label = cell(start+offset, column)
            label.Value = "Phase" if name.startswith("Phase") else name
            fields.append(SimpleNamespace(SourceName=name, LabelRange=label))
        pivots = [SimpleNamespace(PageFields=fields, TableRange2=SimpleNamespace(
            Row=start, Column=column, Rows=SimpleNamespace(Count=8), Columns=SimpleNamespace(Count=2)))]
        return SimpleNamespace(PivotTables=lambda: pivots, Cells=cell), cells, pivots

    def test_remove_old_d1_caption_without_touching_live_d4_filter(self):
        sheet, cells, _ = self.setup_sheet()
        sheet.Cells(1, 4).Value = "YEAR FAB UPLOAD"
        sheet.Cells(4, 5).Value = "(Multiple Items)"
        self.assertEqual(restore_on_progress_filter_caption_layout(sheet), 1)
        self.assertIsNone(cells[1, 4].Value)
        self.assertEqual(cells[4, 4].Value, "YEAR FAB UPLOAD")
        self.assertEqual(cells[4, 5].Value, "(Multiple Items)")
        self.assertEqual(restore_on_progress_filter_caption_layout(sheet), 0)

    def test_valid_filter_in_d1_is_not_deleted(self):
        sheet, cells, pivots = self.setup_sheet(start=1)
        pivots[0].PageFields[0].SourceName = "YEAR FAB UPLOAD"
        pivots[0].PageFields[2].SourceName = "Process Adjustment"
        cells[1, 4].Value = "YEAR FAB UPLOAD"
        cells[3, 4].Value = "Process Adjustment"
        restore_on_progress_filter_caption_layout(sheet)
        self.assertEqual(cells[1, 4].Value, "YEAR FAB UPLOAD")
        self.assertFalse(cells[1, 4].cleared)

    def test_follow_moved_filter_positions_without_writing_fixed_cells(self):
        sheet, cells, _ = self.setup_sheet(start=8, column=7)
        sheet.Cells(6, 7).Value = "YEAR FAB UPLOAD"
        self.assertEqual(restore_on_progress_filter_caption_layout(sheet), 1)
        self.assertEqual(cells[10, 7].Value, "YEAR FAB UPLOAD")
        self.assertNotIn((1, 4), cells)

    def test_preserve_formulas_notes_and_other_pivot_cells(self):
        sheet, cells, pivots = self.setup_sheet(start=4)
        sheet.Cells(1, 4).Value = "YEAR FAB UPLOAD"
        sheet.Cells(1, 4).HasFormula = True
        sheet.Cells(2, 4).Value = "Business note"
        sheet.Cells(3, 4).Value = "YEAR FAB UPLOAD"
        pivots.append(SimpleNamespace(PageFields=[], TableRange2=SimpleNamespace(
            Row=3, Column=4, Rows=SimpleNamespace(Count=1), Columns=SimpleNamespace(Count=2))))
        self.assertEqual(restore_on_progress_filter_caption_layout(sheet), 0)
        self.assertEqual(cells[1, 4].Value, "YEAR FAB UPLOAD")
        self.assertEqual(cells[2, 4].Value, "Business note")
        self.assertFalse(cells[3, 4].cleared)


if __name__ == "__main__":
    unittest.main()
