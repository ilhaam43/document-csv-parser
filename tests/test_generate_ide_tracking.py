import unittest
from datetime import date, datetime

from generate_ide_tracking import (
    financial_value_or_zero,
    parse_excel_datetime,
    target_value,
)


class IdeDateParsingTests(unittest.TestCase):
    def test_slash_dates_follow_dashboard_day_first_format(self) -> None:
        self.assertEqual(
            parse_excel_datetime("10/01/2026"),
            datetime(2026, 1, 10),
        )
        self.assertEqual(
            parse_excel_datetime("04/05/2026"),
            datetime(2026, 5, 4),
        )

    def test_iso_dates_are_unchanged(self) -> None:
        self.assertEqual(
            parse_excel_datetime("2026-08-13 07:30:00"),
            datetime(2026, 8, 13, 7, 30),
        )

    def test_invalid_date_is_not_coerced(self) -> None:
        self.assertIsNone(parse_excel_datetime("24/26/2026"))

    def test_target_bucket_uses_parsed_calendar_month(self) -> None:
        report_date = date(2026, 8, 13)
        self.assertEqual(target_value("10/01/2026", report_date), "Before August")
        self.assertEqual(target_value("13/08/2026", report_date), "Target August")
        self.assertEqual(target_value("01/09/2026", report_date), "After August")

    def test_missing_previous_financial_value_falls_back_to_zero(self) -> None:
        self.assertEqual(financial_value_or_zero("1,250,000"), 1_250_000)
        self.assertEqual(financial_value_or_zero("#N/A"), 0)
        self.assertEqual(financial_value_or_zero(None), 0)


if __name__ == "__main__":
    unittest.main()
