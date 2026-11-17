import unittest
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

from generate_ide_tracking import InputFiles, determine_report_date, generate_ide_tracking, target_header
from ide_list_format import REQUIRED, build_rows, generate, is_list_export, normalize_status, read_records, target_bucket


class IdeListRulesTests(unittest.TestCase):
    def raw(self, **values):
        row = dict.fromkeys(REQUIRED)
        row.update({"Key": "IDE-1", "Quote ID": "Q1", "Quote Name": "New connection",
                    "Status": "SO Complete", "Target RFS": datetime(2026, 11, 30),
                    "Actual RFS Date": datetime(2026, 11, 29), "Revenue": 0})
        row.update(values)
        return row

    def reference(self, **values):
        return {"Quote ID": "Q1", "Status": "3.Delivery", "ORDER TYPE": "New Registration", **values}

    def build(self, rows, previous=None, day=date(2026, 11, 17)):
        return build_rows(rows, previous or [self.reference()], day)[0]

    def test_first_quote_history_preserves_every_unique_key_and_current_status(self):
        rows = [self.raw(), self.raw(Key="IDE-2", Status="Cancel")]
        frame = self.build(rows, [self.reference(), self.reference(Status="5.SO Complete", **{"ORDER TYPE": "Renewal"})])
        self.assertEqual(frame["Key"].tolist(), ["IDE-1", "IDE-2"])
        self.assertEqual(frame["Yesterday Status "].tolist(), ["3.Delivery", "3.Delivery"])
        self.assertEqual(frame["Status"].tolist(), ["5.SO Complete", "6.Cancel"])
        self.assertEqual(frame["ORDER TYPE"].tolist(), ["New Registration"] * 2)

    def test_new_renewal_quote_does_not_require_a_history_match(self):
        for title in ("Renewal product", "RENEWEL LICENCE ZOOM", "Renew contract"):
            with self.subTest(title=title):
                row = self.build([self.raw(**{"Quote ID": "NEW", "Quote Name": title})]).iloc[0]
                self.assertEqual(row["ORDER TYPE"], "Renewal")
                self.assertEqual(row["PROCESS"], "Non New Reg")
                self.assertEqual(row["Yesterday Status "], "#N/A")

    def test_raw_revenue_and_missing_dates_do_not_fall_back_to_history(self):
        frame = self.build([self.raw(Revenue=None, **{"SO Completion Date": None})],
                           [self.reference(Revenue=999, **{"SO Completion Date": datetime(2026, 11, 1)})])
        self.assertIsNone(frame.iloc[0]["Revenue"])
        self.assertEqual(frame.iloc[0]["MONTH COMPLETE"], "Jan-00")
        self.assertEqual(frame.iloc[0]["YEAR FAB UPLOAD"], "1900")
        self.assertEqual(frame.iloc[0]["Status"], "5.SO Complete")

    def test_aging_uses_report_date_and_target_not_actual(self):
        for difference, status in [(-6, "On Track"), (-5, "Potential Delay"), (-1, "Potential Delay"), (0, "Delay"), (1, "Delay")]:
            from datetime import timedelta
            row = self.build([self.raw(**{"Target RFS": datetime(2026, 11, 17) - timedelta(days=difference)})]).iloc[0]
            self.assertEqual(row["Aging Of RFS"], difference)
            self.assertEqual(row["STATUS ORDER"], status)

    def test_valid_target_is_not_erased_by_missing_fab_or_actual(self):
        row = self.build([self.raw(**{"FAB Upload Date": None, "Actual RFS Date": None})]).iloc[0]
        self.assertEqual(row[target_header(date(2026, 11, 17))], "Target November")
        self.assertIsNone(row["ON TIME DELIVERY"])

    def test_source_month_is_separate_from_generated_month_and_new_fields_survive(self):
        row = self.build([self.raw(**{"Month Complete": "raw-label", "SO Completion Date": datetime(2026, 10, 31),
                                     "Future Field": "kept"})]).iloc[0]
        self.assertEqual(row["Month Complete2"], "raw-label")
        self.assertEqual(row["MONTH COMPLETE"], "Oct-26")
        self.assertEqual(row["Future Field"], "kept")

    def test_unusual_source_year_is_reported_not_guessed(self):
        frame, warnings = build_rows([self.raw(**{"FAB Upload Date": datetime(2926, 7, 28)})],
                                     [self.reference()], date(2026, 11, 17))
        self.assertEqual(frame.iloc[0]["YEAR FAB UPLOAD"], "2926")
        self.assertEqual(warnings[0]["field"], "FAB Upload Date")

    def test_duplicate_key_unknown_status_and_text_dates_fail_explicitly(self):
        for rows, message in [([self.raw(), self.raw()], "unique"),
                              ([self.raw(Status="mystery")], "Unrecognized"),
                              ([self.raw(**{"Target RFS": "11/12/2026"})], "Excel date")]:
            with self.subTest(message=message), self.assertRaisesRegex(ValueError, message):
                self.build(rows)

    def test_targets_follow_calendar_month_including_year_rollover(self):
        self.assertEqual(target_bucket(datetime(2026, 12, 31), date(2027, 1, 1)), "Before January")
        self.assertEqual(target_bucket(datetime(2027, 1, 1), date(2027, 1, 1)), "Target January")
        self.assertEqual(target_bucket(None, date(2027, 1, 1)), "Target Not Yet Inputted")

    def test_all_six_statuses_are_idempotent(self):
        for original, expected in [("Quotation", "1.Quotation"), ("Pra-Delivery", "2.Pra-Delivery"),
                                   ("Delivery", "3.Delivery"), ("Waiting AOS", "4.Waiting AOS"),
                                   ("SO Complete", "5.SO Complete"), ("Cancel", "6.Cancel")]:
            self.assertEqual(normalize_status(original), expected)
            self.assertEqual(normalize_status(expected), expected)

    def test_iso_filename_and_dispatch_without_collabs(self):
        self.assertEqual(determine_report_date(None, Path("IDE-list-2026-11-17.xlsx")), date(2026, 11, 17))
        files = InputFiles(Path("raw.xlsx"), Path("reference.xlsx"))
        with patch("ide_list_format.is_list_export", return_value=True), \
             patch("ide_list_format.generate", return_value={"format_profile": "ide-list"}) as generate:
            self.assertEqual(generate_ide_tracking(files, Path("output.xlsx"), date(2026, 11, 17))["format_profile"], "ide-list")
            generate.assert_called_once()
        with patch("ide_list_format.is_list_export", return_value=False), \
             self.assertRaisesRegex(ValueError, "requires.*CSV"):
            generate_ide_tracking(files, Path("output.xlsx"), date(2026, 11, 17))

    def test_detects_content_and_preserves_first_data_row(self):
        row = self.raw()
        sheet = SimpleNamespace(iter_rows=lambda **_: iter([tuple(row), tuple(row.values())]))
        class Book:
            sheetnames = ["List"]
            def __getitem__(self, name):
                return sheet
            def close(self):
                pass
        with patch("ide_list_format.load_workbook", side_effect=lambda *a, **k: Book()):
            self.assertTrue(is_list_export(Path("renamed.xlsx")))
            self.assertEqual(read_records(Path("renamed.xlsx"), "List"), [row])

    def test_generation_failure_does_not_publish_partial_output(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            reference, output = root / "reference.xlsx", root / "output.xlsx"
            reference.write_bytes(b"template")
            output.write_bytes(b"existing-output")
            with patch("ide_list_format.read_records", side_effect=[[self.raw()], [self.reference()]]), \
                 patch("ide_list_format._write_native", side_effect=RuntimeError("native refresh failed")), \
                 self.assertRaisesRegex(RuntimeError, "native refresh failed"):
                generate(root / "raw.xlsx", reference, output, date(2026, 11, 17))
            self.assertEqual(output.read_bytes(), b"existing-output")
            self.assertEqual(reference.read_bytes(), b"template")
            self.assertFalse(list(root.glob("ide_list_*.xlsx")))


if __name__ == "__main__":
    unittest.main()
