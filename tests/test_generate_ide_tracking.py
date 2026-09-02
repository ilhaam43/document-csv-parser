import unittest
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from generate_ide_tracking import (
    ACTUAL_RFS_DATE_HEADER,
    DEPT_HEADER,
    NEW_RFS_INITIAL_HEADER,
    QUOTE_ID_HEADER,
    apply_current_date_first_match,
    apply_previous_field_fallbacks,
    completion_thresholds,
    date_from_filename,
    financial_value_or_zero,
    hidden_row_ranges_between_sections,
    identify_target_complete_pivots,
    is_missing,
    load_collabs_lookup,
    load_source_new_rfs_lookup,
    normalize_month_first_dashboard_dates,
    parse_excel_datetime,
    reconcile_date_column,
    reconciled_date_value,
    resolve_financial_values,
    target_period,
    target_value,
    target_value_with_previous,
)


class FakeRangeDimension:
    def __init__(self, count):
        self.Count = count


class FakeTableRange:
    def __init__(self, row, column, rows, columns):
        self.Row = row
        self.Column = column
        self.Rows = FakeRangeDimension(rows)
        self.Columns = FakeRangeDimension(columns)


class FakePivot:
    def __init__(self, name, row, column, rows, columns):
        self.Name = name
        self.TableRange2 = FakeTableRange(row, column, rows, columns)


class FakePivotCollection:
    def __init__(self, pivots):
        self._pivots = pivots
        self.Count = len(pivots)

    def __call__(self, index):
        return self._pivots[index - 1]


class FakePivotWorksheet:
    def __init__(self, pivots):
        self._pivots = FakePivotCollection(pivots)

    def PivotTables(self):
        return self._pivots


class IdeDateParsingTests(unittest.TestCase):
    def test_hidden_row_ranges_follow_dynamic_section_boundaries(self) -> None:
        self.assertEqual(
            hidden_row_ranges_between_sections(25, 113),
            [(27, 110)],
        )
        self.assertEqual(
            hidden_row_ranges_between_sections(137, 166, {144}),
            [(139, 142), (145, 163)],
        )

    def test_target_complete_pivots_allow_dynamic_start_rows(self) -> None:
        summary = FakePivot("summary", 7, 1, 45, 3)
        status = FakePivot("status", 6, 4, 46, 7)
        later_left = FakePivot("later-left", 68, 1, 20, 5)
        later_right = FakePivot("later-right", 65, 8, 18, 4)
        worksheet = FakePivotWorksheet(
            [later_right, status, later_left, summary]
        )

        actual_summary, actual_status = identify_target_complete_pivots(worksheet)

        self.assertIs(actual_summary, summary)
        self.assertIs(actual_status, status)

    def test_target_complete_pivots_reject_unrelated_later_section(self) -> None:
        summary = FakePivot("summary", 7, 1, 45, 3)
        unrelated_status = FakePivot("later-right", 65, 8, 18, 4)
        worksheet = FakePivotWorksheet([summary, unrelated_status])

        with self.assertRaisesRegex(RuntimeError, "TARGET COMPLETE"):
            identify_target_complete_pivots(worksheet)

    def test_report_filename_accepts_english_month_with_any_case(self) -> None:
        expected = date(2026, 8, 18)

        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 August 2026.xlsx")), expected)
        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 august 2026.xlsx")), expected)
        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 AUGUST 2026.xlsx")), expected)

    def test_report_filename_accepts_indonesian_month_with_any_case(self) -> None:
        expected = date(2026, 8, 18)

        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 Agustus 2026.xlsx")), expected)
        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 agustus 2026.xlsx")), expected)
        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 AGUSTUS 2026.xlsx")), expected)

    def test_report_filename_accepts_english_and_indonesian_month_abbreviations(self) -> None:
        cases = {
            "IDE DASHBOARD 18 Aug 2026.xlsx": date(2026, 8, 18),
            "IDE DASHBOARD 18 AGU 2026.xlsx": date(2026, 8, 18),
            "IDE DASHBOARD 18 Agt 2026.xlsx": date(2026, 8, 18),
            "IDE DASHBOARD 18 Sep 2026.xlsx": date(2026, 9, 18),
            "IDE DASHBOARD 18 Sept 2026.xlsx": date(2026, 9, 18),
            "IDE DASHBOARD 18 Oct 2026.xlsx": date(2026, 10, 18),
            "IDE DASHBOARD 18 Okt 2026.xlsx": date(2026, 10, 18),
            "IDE DASHBOARD 18 Dec 2026.xlsx": date(2026, 12, 18),
            "IDE DASHBOARD 18 Des 2026.xlsx": date(2026, 12, 18),
        }

        for filename, expected in cases.items():
            with self.subTest(filename=filename):
                self.assertEqual(date_from_filename(Path(filename)), expected)

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

    def test_month_opening_carries_previous_rfs_dates(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                NEW_RFS_INITIAL_HEADER: [datetime(2026, 8, 9)],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                NEW_RFS_INITIAL_HEADER: [datetime(2026, 9, 8)],
            }
        )

        values, fallback_count = reconcile_date_column(
            current,
            previous,
            NEW_RFS_INITIAL_HEADER,
            date(2026, 9, 1),
        )

        self.assertEqual(list(values), [datetime(2026, 9, 8)])
        self.assertEqual(fallback_count, 1)

    def test_non_opening_day_keeps_current_rfs_date(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                NEW_RFS_INITIAL_HEADER: [datetime(2026, 8, 9)],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                NEW_RFS_INITIAL_HEADER: [datetime(2026, 9, 8)],
            }
        )

        values, fallback_count = reconcile_date_column(
            current,
            previous,
            NEW_RFS_INITIAL_HEADER,
            date(2026, 9, 2),
        )

        self.assertEqual(list(values), [datetime(2026, 8, 9)])
        self.assertEqual(fallback_count, 0)

    def test_month_opening_matches_duplicate_actual_rfs_by_occurrence(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1", "Q-1"],
                ACTUAL_RFS_DATE_HEADER: [
                    datetime(2026, 4, 15),
                    datetime(2026, 4, 15),
                ],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1", "Q-1"],
                ACTUAL_RFS_DATE_HEADER: [
                    datetime(2026, 4, 15),
                    datetime(2026, 4, 30),
                ],
            }
        )

        values, fallback_count = reconcile_date_column(
            current,
            previous,
            ACTUAL_RFS_DATE_HEADER,
            date(2026, 9, 1),
        )

        self.assertEqual(
            list(values),
            [datetime(2026, 4, 15), datetime(2026, 4, 30)],
        )
        self.assertEqual(fallback_count, 1)

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

    def test_first_calendar_day_closes_the_previous_target_month(self) -> None:
        self.assertEqual(target_period(date(2026, 9, 1)), date(2026, 8, 31))
        self.assertEqual(target_period(date(2026, 1, 1)), date(2025, 12, 31))

    def test_later_report_days_use_the_current_target_month(self) -> None:
        self.assertEqual(target_period(date(2026, 9, 2)), date(2026, 9, 2))

    def test_missing_current_target_carries_the_previous_bucket_in_same_period(self) -> None:
        self.assertEqual(
            target_value_with_previous(
                pd.NA,
                "Target August",
                date(2026, 8, 31),
            ),
            "Target August",
        )

    def test_current_target_date_takes_precedence_over_previous(self) -> None:
        self.assertEqual(
            target_value_with_previous(
                datetime(2026, 9, 15),
                "Target Not Yet Inputted",
                date(2026, 8, 31),
            ),
            "After August",
        )

    def test_existing_target_is_carried_within_the_same_period(self) -> None:
        self.assertEqual(
            target_value_with_previous(
                datetime(2026, 8, 9),
                "After August",
                date(2026, 8, 31),
            ),
            "After August",
        )

    def test_existing_target_is_reclassified_when_period_changes(self) -> None:
        self.assertEqual(
            target_value_with_previous(
                pd.NA,
                "Target August",
                date(2026, 9, 2),
            ),
            "Target Not Yet Inputted",
        )

    def test_missing_previous_financial_value_falls_back_to_zero(self) -> None:
        self.assertEqual(financial_value_or_zero("1,250,000"), 1_250_000)
        self.assertEqual(financial_value_or_zero("#N/A"), 0)
        self.assertEqual(financial_value_or_zero(None), 0)

    def test_collabs_lookup_reads_only_requested_quote_ids(self) -> None:
        with TemporaryDirectory() as temp_dir:
            collabs_path = Path(temp_dir) / "collabs.csv"
            pd.DataFrame(
                {
                    "quote_num": ["Q-1", "Q-2"],
                    "otc": ["1,250", "9,999"],
                    "mrc": ["250", "999"],
                    "unused": ["x", "y"],
                }
            ).to_csv(collabs_path, index=False)

            lookup = load_collabs_lookup(collabs_path, {"Q-1"})

        self.assertEqual(lookup, {"Q-1": (1_250, 250)})

    def test_financial_fallback_uses_previous_then_collabs_then_zero(self) -> None:
        self.assertEqual(
            resolve_financial_values(100, 200, (300, 400)),
            (100, 200, False, False),
        )
        self.assertEqual(
            resolve_financial_values("#N/A", None, (300, 400)),
            (300, 400, True, False),
        )
        self.assertEqual(
            resolve_financial_values(None, None, (None, "#N/A")),
            (0, 0, False, True),
        )

    def test_completion_thresholds_follow_report_week_and_cap_at_week_four(self) -> None:
        expectations = {
            1: (1, 30, 30, 20),
            7: (1, 30, 30, 20),
            8: (2, 40, 40, 30),
            14: (2, 40, 40, 30),
            15: (3, 50, 50, 40),
            21: (3, 50, 50, 40),
            22: (4, 60, 60, 50),
            31: (4, 60, 60, 50),
        }

        for day, expected in expectations.items():
            with self.subTest(day=day):
                thresholds = completion_thresholds(date(2026, 8, day))
                self.assertEqual(
                    (
                        thresholds.week,
                        thresholds.green_percent,
                        thresholds.yellow_percent,
                        thresholds.red_percent,
                    ),
                    expected,
                )


if __name__ == "__main__":
    unittest.main()
