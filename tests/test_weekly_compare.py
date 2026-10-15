import unittest
from datetime import datetime
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from weekly.compare import _business_equal, _equal, _hidden_content_rows, compare_weekly_workbooks
from weekly.regression import _normalized_value


class WeeklyCompareTests(unittest.TestCase):
    def test_reference_audit_never_hides_a_zero_versus_one_leadtime_difference(self) -> None:
        for field in ("LEADTIME FAB-CASH", "LEADTIME CASH-COMPLETE"):
            self.assertNotEqual(_normalized_value(field, 0), _normalized_value(field, 1))

    def test_comparison_preserves_text_and_timestamp_precision(self) -> None:
        self.assertTrue(_equal(2.5714285714285716, 2.571428571428572))
        self.assertFalse(_equal(0, 1))
        self.assertFalse(_equal("R  NINDYA", "R NINDYA"))
        self.assertFalse(_equal(datetime(2026, 9, 20, 22, 43, 16), datetime(2026, 9, 20, 22, 43)))

    def test_hidden_content_detection_ignores_empty_separator_rows(self) -> None:
        xml = b"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="48" hidden="1"><c r="G48"><v>Grand Total</v></c></row>
            <row r="49" hidden="1"><c r="A49" s="1"/></row>
            <row r="50"><c r="G50"><v>Visible</v></c></row>
            <row r="51" hidden="1"><c r="A51"><f>SUM(A1:A4)</f></c></row>
          </sheetData>
        </worksheet>"""
        self.assertEqual(_hidden_content_rows(BytesIO(xml)), {48, 51})

    def test_comparison_counts_differences_by_column(self) -> None:
        with TemporaryDirectory() as directory:
            paths = [Path(directory) / f"{name}.xlsx" for name in ("reference", "generated")]
            for index, path in enumerate(paths):
                book = Workbook()
                sheet = book.active
                sheet.title = "ALL DATA"
                headers = [None] * 11
                headers[0], headers[7], headers[10] = "No.", "order_status", "fab_upload"
                sheet.append(headers)
                values = [None] * 11
                values[0], values[7], values[10] = (
                    2.5 if index == 0 else 1,
                    "Complete",
                    datetime(2026, 9, 1, 1, 2, 3 if index == 0 else 0),
                )
                sheet.append(values)
                book.save(path)
            result = compare_weekly_workbooks(*paths)
            business = compare_weekly_workbooks(*paths, business_only=True)
        self.assertFalse(result["matches"])
        self.assertEqual(result["sheets"]["ALL DATA"]["differences_by_column"], {1: 1, 11: 1})
        self.assertTrue(business["matches"])
        self.assertEqual(business["sheets"]["ALL DATA"]["difference_count"], 0)

    def test_business_comparison_still_rejects_status_difference(self) -> None:
        self.assertFalse(
            _business_equal("ALL DATA", 2, 8, "Open", "Complete")
        )


if __name__ == "__main__":
    unittest.main()
