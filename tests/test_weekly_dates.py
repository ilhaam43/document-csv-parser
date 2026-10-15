import unittest
from datetime import date, datetime

from weekly.dates import excel_week_number, has_explicit_seconds, parse_datetime


class WeeklyDatesTests(unittest.TestCase):
    def test_parse_datetime_accepts_bq_and_iso_formats(self) -> None:
        self.assertEqual(parse_datetime("9/15/2026 16:30"), datetime(2026, 9, 15, 16, 30))
        self.assertEqual(parse_datetime("2026-09-15"), datetime(2026, 9, 15))
        self.assertIsNone(parse_datetime(""))

    def test_bigquery_utc_suffix_preserves_source_time_and_fractional_seconds(self) -> None:
        self.assertEqual(parse_datetime("2026-10-07 23:59:59 UTC"), datetime(2026, 10, 7, 23, 59, 59))
        self.assertEqual(parse_datetime("2026-10-07 23:59:59.123456 UTC"), datetime(2026, 10, 7, 23, 59, 59, 123456))
        self.assertEqual(parse_datetime("2026-10-07 23:59:59Z"), datetime(2026, 10, 7, 23, 59, 59))

    def test_invalid_dates_and_ambiguous_timezone_labels_remain_errors(self) -> None:
        for text in ("2026-02-30 00:00:00 UTC", "2026-10-07 12:00:00 CST"):
            with self.subTest(text=text), self.assertRaisesRegex(ValueError, "Unsupported date value"):
                parse_datetime(text)

    def test_explicit_timestamp_precision(self) -> None:
        self.assertFalse(has_explicit_seconds("9/30/2026 10:02"))
        self.assertFalse(has_explicit_seconds("2026-09-30"))
        self.assertTrue(has_explicit_seconds("9/30/2026 10:02:00"))
        self.assertTrue(has_explicit_seconds("2026-09-30T10:02:29"))
        self.assertTrue(has_explicit_seconds(datetime(2026, 9, 30, 10, 2, 29)))

    def test_excel_week_number_uses_thursday_start(self) -> None:
        self.assertEqual(excel_week_number(date(2026, 9, 9)), 36)
        self.assertEqual(excel_week_number(date(2026, 9, 10)), 37)
        self.assertEqual(excel_week_number(date(2026, 9, 17)), 38)

    def test_excel_week_number_rejects_other_return_types(self) -> None:
        with self.assertRaises(ValueError):
            excel_week_number(date(2026, 9, 17), return_type=2)


if __name__ == "__main__":
    unittest.main()
