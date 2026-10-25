import unittest
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from weekly.io import write_master_csv
from weekly.regression import compare_master_records, reference_formula_mismatches, run_reference_audit


class WeeklyRegressionTests(unittest.TestCase):
    def record(self, **updates):
        return {
            "No.": 1, "quote_num": "Q-1", "order number": "O-1",
            "fab_upload": datetime(2032, 4, 1, 9, 30), "cash_dt": datetime(2032, 4, 2),
            "service_group": "Mobile", "LEADTIME FAB-CASH": 2,
            "otc": 1_000_000, "mrc": 10_000,
            "OTC+MRC": Decimal("1010000"), "OTC+MRC (inMio)": Decimal("1.01"),
            **updates,
        }

    def test_missing_and_extra_duplicate_occurrences_fail(self):
        row = self.record()
        self.assertEqual(compare_master_records([row], [row, row])["differences_by_field"], {"<missing generated row>": 1})
        self.assertEqual(compare_master_records([row, row], [row])["differences_by_field"], {"<extra generated row>": 1})

    def test_identity_matching_accepts_reordering_but_preserves_duplicate_values(self):
        first = self.record()
        second = self.record(quote_num="Q-2")
        self.assertTrue(compare_master_records([second, first], [first, second])["matches"])
        changed = self.record(**{"LEADTIME FAB-CASH": 0})
        self.assertEqual(compare_master_records([first, changed], [first, first])["differences_by_field"], {"LEADTIME FAB-CASH": 1})

    def test_only_approved_precision_differences_are_ignored(self):
        expected = self.record()
        actual = self.record(**{"No.": 123.45, "fab_upload": datetime(2032, 4, 1, 9, 30, 55), "OTC+MRC (inMio)": 1.010000000001})
        self.assertTrue(compare_master_records([actual], [expected])["matches"])
        actual["OTC+MRC (inMio)"] = 1.0100001
        self.assertFalse(compare_master_records([actual], [expected])["matches"])
        actual = self.record(ca_name="PT. A  B")
        self.assertFalse(compare_master_records([actual], [self.record(ca_name="PT. A B")])["matches"])

    def test_reference_diagnostics_do_not_mutate_or_suppress_parity(self):
        expected = self.record(**{"LEADTIME FAB-CASH": 999, "OTC+MRC": None, "OTC+MRC (inMio)": None})
        before = dict(expected)
        self.assertEqual(reference_formula_mismatches([expected], set()), {
            "LEADTIME FAB-CASH": 1, "OTC+MRC": 1, "OTC+MRC (inMio)": 1,
        })
        self.assertFalse(compare_master_records([self.record()], [expected])["matches"])
        self.assertEqual(expected, before)

    def test_arbitrary_sample_uses_production_finalizer_without_expected_workbooks(self):
        row = self.record()

        def finalize(**kwargs):
            self.assertNotIn("expected_master", kwargs)
            self.assertNotIn("expected_mobile", kwargs)
            self.assertEqual(kwargs["report_month"], date(2032, 4, 1))
            for name in ("master-final-data.csv", "mobile-data.csv"):
                write_master_csv(Path(kwargs["output_dir"]) / name, [row])
            return {"pending_catalyst_rows": 0, "invalid_master_services": [], "issue_counts": {}}

        with patch("weekly.regression.finalize_weekly_data", side_effect=finalize) as production:
            with patch("weekly.regression.iter_master_rows", side_effect=lambda _: iter([row])):
                with patch("weekly.regression.load_holidays", return_value=set()):
                    audit = run_reference_audit(
                        bq_csv="bq.csv", previous_master="previous.xlsx", expected_master="manual.xlsx",
                        expected_mobile="manual-mobile.xlsx", catalyst_review="review.csv",
                        holiday_workbook="holidays.xlsx", report_date=date(2032, 4, 5), report_month=date(2032, 4, 1),
                    )
        production.assert_called_once()
        self.assertTrue(audit.passed, audit.failures)


if __name__ == "__main__":
    unittest.main()
