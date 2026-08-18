import unittest
from datetime import date, datetime

import pandas as pd

from generate_ide_tracking import (
    ACTUAL_RFS_DATE_HEADER,
    DEPT_HEADER,
    NEW_RFS_INITIAL_HEADER,
    QUOTE_ID_HEADER,
    apply_current_date_first_match,
    apply_previous_field_fallbacks,
    financial_value_or_zero,
    is_missing,
    load_source_new_rfs_lookup,
    normalize_month_first_dashboard_dates,
    parse_excel_datetime,
    reconcile_date_column,
    reconciled_date_value,
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

    def test_month_first_can_be_selected_for_region_sensitive_fields(self) -> None:
        self.assertEqual(
            parse_excel_datetime("08/06/2026", dayfirst=False),
            datetime(2026, 8, 6),
        )
        self.assertEqual(
            parse_excel_datetime("22/06/2026", dayfirst=False),
            datetime(2026, 6, 22),
        )

    def test_region_sensitive_dashboard_dates_are_normalized_month_first(self) -> None:
        dataframe = pd.DataFrame(
            {
                NEW_RFS_INITIAL_HEADER: ["08/06/2026"],
                ACTUAL_RFS_DATE_HEADER: ["11/05/2026"],
            }
        )

        normalize_month_first_dashboard_dates(dataframe)

        self.assertEqual(dataframe.at[0, NEW_RFS_INITIAL_HEADER], datetime(2026, 8, 6))
        self.assertEqual(dataframe.at[0, ACTUAL_RFS_DATE_HEADER], datetime(2026, 11, 5))

    def test_duplicate_first_match_does_not_fill_an_empty_current_row(self) -> None:
        dataframe = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["A", "A", "B", "B"],
                NEW_RFS_INITIAL_HEADER: [
                    datetime(2026, 4, 15),
                    datetime(2026, 4, 30),
                    datetime(2026, 6, 8),
                    pd.NA,
                ],
                ACTUAL_RFS_DATE_HEADER: [
                    datetime(2026, 4, 15),
                    datetime(2026, 4, 30),
                    datetime(2026, 7, 1),
                    pd.NA,
                ],
            }
        )

        apply_current_date_first_match(dataframe)

        self.assertEqual(dataframe.at[1, NEW_RFS_INITIAL_HEADER], datetime(2026, 4, 30))
        self.assertEqual(dataframe.at[1, ACTUAL_RFS_DATE_HEADER], datetime(2026, 4, 15))
        self.assertTrue(is_missing(dataframe.at[3, NEW_RFS_INITIAL_HEADER]))
        self.assertTrue(is_missing(dataframe.at[3, ACTUAL_RFS_DATE_HEADER]))

    def test_source_new_rfs_lookup_uses_first_match_and_us_date_parsing(self) -> None:
        class WorksheetStub:
            def iter_rows(self, **_kwargs):
                return iter(
                    [
                        ("A", *([None] * 33), "08-06-2026 20:19:55"),
                        ("A", *([None] * 33), "09-06-2026 20:19:55"),
                    ]
                )

        lookup = load_source_new_rfs_lookup(WorksheetStub(), {"A"})

        self.assertEqual(lookup["A"], datetime(2026, 8, 6))

    def test_missing_current_date_is_not_filled_from_previous(self) -> None:
        value, used_fallback = reconciled_date_value(
            pd.NA,
            datetime(2026, 7, 1),
            date(2026, 8, 14),
        )

        self.assertTrue(is_missing(value))
        self.assertFalse(used_fallback)

    def test_actual_rfs_reconciliation_uses_first_quote_lookup_key(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1", "Q-1"],
                ACTUAL_RFS_DATE_HEADER: [
                    datetime(2025, 12, 16),
                    datetime(2025, 12, 16),
                ],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                ACTUAL_RFS_DATE_HEADER: [datetime(2026, 12, 16)],
            }
        )

        values, fallback_count = reconcile_date_column(
            current,
            previous,
            ACTUAL_RFS_DATE_HEADER,
            date(2026, 8, 14),
        )

        self.assertEqual(list(values), [datetime(2026, 12, 16)] * 2)
        self.assertEqual(fallback_count, 2)

    def test_previous_field_fallback_uses_first_quote_match_and_preserves_na_error(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["A", "A", "B"],
                DEPT_HEADER: [pd.NA, pd.NA, "#N/A"],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["A", "A"],
                DEPT_HEADER: ["First Dept", "Second Dept"],
            }
        )

        fallback_count = apply_previous_field_fallbacks(current, previous, [DEPT_HEADER])

        self.assertEqual(fallback_count, 2)
        self.assertEqual(current[DEPT_HEADER].iloc[:2].tolist(), ["First Dept", "First Dept"])
        self.assertEqual(current.at[2, DEPT_HEADER], "#N/A")

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
