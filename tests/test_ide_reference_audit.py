import unittest
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from ide_reference_audit import compare_all_order


class IdeReferenceAuditTests(unittest.TestCase):
    def write(self, path, headers, rows):
        book = Workbook()
        sheet = book.active
        sheet.title = "ALL ORDER"
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
        book.save(path)
        book.close()

    def compare(self, headers, manual, generated, generated_headers=None):
        with TemporaryDirectory() as directory:
            left, right = Path(directory) / "manual.xlsx", Path(directory) / "generated.xlsx"
            self.write(left, headers, manual)
            self.write(right, generated_headers or headers, generated)
            return compare_all_order(left, right)

    def test_equal_values_and_schema_pass(self):
        report = self.compare(["QUOTE ID", "OTC"], [["A", 100]], [["A", 100]])
        self.assertTrue(report["exact_value_matches"])

    def test_missing_and_extra_columns_do_not_disappear_from_comparison(self):
        report = self.compare(
            ["QUOTE ID", "STATUS DELIVERY 7 April 2027"], [["A", "Complete"]],
            [["A", "Complete"]], ["QUOTE ID", "STATUS DELIVERY 8 April 2027"],
        )
        self.assertFalse(report["exact_value_matches"])
        self.assertEqual(report["schema"]["missing_columns"], ["STATUS DELIVERY 7 April 2027"])
        self.assertEqual(report["schema"]["extra_columns"], ["STATUS DELIVERY 8 April 2027"])

    def test_money_mismatch_is_a_real_value_difference(self):
        report = self.compare(["QUOTE ID", "MRC"], [["A", 0]], [["A", 660000]])
        self.assertEqual(report["value_difference_cells"], 1)
        self.assertEqual(report["changed_rows"], 1)

    def test_date_text_and_excel_date_are_reported_not_silently_equal(self):
        report = self.compare(
            ["QUOTE ID", "CASH DATE"], [["A", "26/02/2027"]],
            [["A", datetime(2027, 2, 26)]],
        )
        self.assertFalse(report["exact_value_matches"])
        self.assertTrue(report["semantic_value_matches"])
        self.assertEqual(report["representation_difference_cells"], 1)

    def test_different_year_is_not_ignored_as_formatting(self):
        report = self.compare(
            ["QUOTE ID", "ACTUAL RFS DATE"], [["A", datetime(2027, 2, 26)]],
            [["A", datetime(2026, 2, 26)]],
        )
        self.assertFalse(report["semantic_value_matches"])
        self.assertEqual(report["value_difference_cells"], 1)

    def test_control_characters_in_identifiers_are_not_hidden_by_row_alignment(self):
        report = self.compare(["QUOTE ID", "OTC"], [["A", 0]], [["A\n", 0]])
        self.assertTrue(report["quote_order_matches"])
        self.assertEqual(report["column_differences"]["QUOTE ID"]["count"], 1)
        self.assertFalse(report["exact_value_matches"])

    def test_truncated_output_fails_even_when_common_rows_match(self):
        report = self.compare(["QUOTE ID"], [["A"], ["B"]], [["A"]])
        self.assertFalse(report["quote_order_matches"])
        self.assertFalse(report["exact_value_matches"])


if __name__ == "__main__":
    unittest.main()
