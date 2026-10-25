import os
import unittest
from datetime import date
from pathlib import Path

from weekly.regression import run_reference_audit


@unittest.skipUnless(
    os.environ.get("WEEKLY_REFERENCE_DIR"),
    "Set WEEKLY_REFERENCE_DIR to audit the supplied W38 reference without parity exemptions.",
)
class WeeklyReferenceW38Tests(unittest.TestCase):
    def test_w38_reference_exposes_missing_post_checkpoint_status_reviews(self) -> None:
        reference_dir = Path(os.environ["WEEKLY_REFERENCE_DIR"])
        audit = run_reference_audit(
            bq_csv=reference_dir / "bq-results-20260924-090453-1790240701004.csv",
            previous_master=reference_dir / "Master Raw Mobile Service - W37 2026 old format.xlsx",
            expected_master=reference_dir / "Master Raw Mobile Service - W38 2026 old format.xlsx",
            expected_mobile=reference_dir / "Master Raw Mobile Service - W38 2026.xlsx",
            holiday_workbook=reference_dir / "list national holiday.xlsx",
            catalyst_review=(
                reference_dir
                / "Master Raw Mobile Service - W38 2026 old format cek manual via catalist.xlsx"
            ),
            report_date=date(2026, 9, 24),
        )
        failed_checks = {
            name for name, check in audit.checks.items()
            if check["actual"] != check["expected"]
        }
        self.assertIn("master_field_mismatches", failed_checks, audit.failures)
        self.assertGreater(audit.checks["master_field_mismatches"]["actual"].get("order_status", 0), 0)
        self.assertFalse(audit.passed)
        self.assertEqual(audit.checks["master_row_count"]["actual"], audit.checks["master_row_count"]["expected"])


if __name__ == "__main__":
    unittest.main()
