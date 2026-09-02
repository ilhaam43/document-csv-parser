import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace

import pandas as pd

from csv_to_excel import (
    PHASE_HEADER_PREFIX,
    PIVOT_TEMPLATE_SERVICE_DELIVERY_DIV_HEADER,
    QUO_HEADER,
    RFS_COMMMIT_HEADER,
    SERVICE_DELIVERY_DIV_HEADER,
    TARGET_SO_COMPLETE_DATE_HEADERS,
    add_pivot_template_compatibility_columns,
    add_aging_of_rfs_column,
    apply_previous_target_date_fallback,
    clean_dataframe,
    ensure_percentage_completion_column_widths,
    fill_missing_with_quo_lookup,
    target_header_month_name,
    target_header_period,
)


class ReportOneBusinessRuleTests(unittest.TestCase):
    def test_cleaning_preserves_pm_display_key_whitespace(self) -> None:
        options = SimpleNamespace(
            normalize_headers=True,
            keep_empty=True,
            drop_empty_columns=False,
            dedupe=False,
            infer_types=False,
        )
        source = pd.DataFrame(
            {
                "#Quo": ["2-674178064625  (2)"],
                "PM": ["Muhammad Mahdi Ramadhan "],
                "Phase": [" 00-New "],
            }
        )

        result = clean_dataframe(source, options, Path("DataOrderSD-20260901.csv"))

        self.assertEqual(result.loc[0, "PM"], "Muhammad Mahdi Ramadhan ")
        self.assertEqual(result.loc[0, "Quo"], "2-674178064625  (2)")
        self.assertEqual(result.loc[0, "Phase"], "00-New")

    def test_rfs_aging_ignores_pre_2000_sentinel_dates(self) -> None:
        source = pd.DataFrame({RFS_COMMMIT_HEADER: ["1999-11-30", "2026-08-31"]})

        result = add_aging_of_rfs_column(source, date(2026, 9, 1))

        self.assertTrue(pd.isna(result.loc[0, "Aging Of RFS"]))
        self.assertEqual(result.loc[1, "Aging Of RFS"], 1)

    def test_current_pm_wins_and_previous_pm_is_only_a_blank_fallback(self) -> None:
        current = pd.Series(["Current PM", "", pd.NA], dtype="string")
        quote_ids = pd.Series(["Q-1", "Q-2", "Q-3"], dtype="string")

        result = fill_missing_with_quo_lookup(
            current,
            quote_ids,
            {"Q-1": "Old PM", "Q-2": "Fallback PM", "Q-3": "Fallback PM 2"},
        )

        self.assertEqual(result.tolist(), ["Current PM", "Fallback PM", "Fallback PM 2"])

    def test_percentage_column_width_is_enforced_after_pivot_refresh(self) -> None:
        percentage_header = "Percentage of Completion (SO Complete, Cancel & Change Target)"

        class Worksheet:
            def __init__(self) -> None:
                self.widths = {column: SimpleNamespace(ColumnWidth=10.0) for column in range(4, 20)}

            def Cells(self, row: int, column: int):
                value = percentage_header if (row, column) == (114, 10) else None
                return SimpleNamespace(Value=value)

            def Columns(self, column: int):
                return self.widths[column]

        worksheet = Worksheet()
        ensure_percentage_completion_column_widths(worksheet)

        self.assertEqual(worksheet.widths[10].ColumnWidth, 22.0)
        self.assertEqual(worksheet.widths[9].ColumnWidth, 10.0)

    def test_target_header_month_accepts_short_and_long_year_formats(self) -> None:
        self.assertEqual(target_header_month_name("TARGET  Detemined as 1 Aug 26"), "August")
        self.assertEqual(target_header_month_name("TARGET Determined as 1 August 2026"), "August")

    def test_target_header_period_normalizes_short_and_long_year_formats(self) -> None:
        self.assertEqual(target_header_period("TARGET  Detemined as 1 Aug 26"), (2026, 8))
        self.assertEqual(target_header_period("TARGET Determined as 1 September 2026"), (2026, 9))

    def test_implausible_future_target_uses_valid_previous_date_and_bucket(self) -> None:
        target_header = "TARGET  Detemined as 1 Aug 26"
        target_date_header = TARGET_SO_COMPLETE_DATE_HEADERS[0]
        data = pd.DataFrame(
            {
                QUO_HEADER: ["2-803768807413"],
                target_header: ["After August"],
                target_date_header: ["2046-05-28"],
            }
        )
        mappings = SimpleNamespace(target_so_complete_date={"2-803768807413": "2026-05-28"})

        result = apply_previous_target_date_fallback(data, mappings, date(2026, 8, 28))

        self.assertEqual(result.loc[0, target_date_header], "2026-05-28")
        self.assertEqual(result.loc[0, target_header], "Before August")

    def test_reasonable_future_target_is_not_replaced(self) -> None:
        target_header = "TARGET  Detemined as 1 Aug 26"
        target_date_header = TARGET_SO_COMPLETE_DATE_HEADERS[0]
        data = pd.DataFrame(
            {
                QUO_HEADER: ["Q-1"],
                target_header: ["After August"],
                target_date_header: ["2027-05-28"],
            }
        )
        mappings = SimpleNamespace(target_so_complete_date={"Q-1": "2026-05-28"})

        result = apply_previous_target_date_fallback(data, mappings, date(2026, 8, 28))

        self.assertEqual(result.loc[0, target_date_header], "2027-05-28")
        self.assertEqual(result.loc[0, target_header], "After August")

    def test_service_delivery_compatibility_renames_instead_of_duplicating(self) -> None:
        data = pd.DataFrame(
            {
                QUO_HEADER: ["Q-1"],
                SERVICE_DELIVERY_DIV_HEADER: ["Seradel"],
                "Phase 27 August 2026": ["04-Pre Installation"],
            }
        )

        result, hidden_columns = add_pivot_template_compatibility_columns(
            data,
            Path("DataOrderSD-20260828-060743.csv"),
            "",
            "Phase 27 August 2026",
        )

        self.assertNotIn(SERVICE_DELIVERY_DIV_HEADER, result.columns)
        self.assertIn(PIVOT_TEMPLATE_SERVICE_DELIVERY_DIV_HEADER, result.columns)
        self.assertEqual(result.loc[0, PIVOT_TEMPLATE_SERVICE_DELIVERY_DIV_HEADER], "Seradel")
        self.assertNotIn(PIVOT_TEMPLATE_SERVICE_DELIVERY_DIV_HEADER, hidden_columns)
        self.assertEqual(result.loc[0, PHASE_HEADER_PREFIX], "04-Pre Installation")


if __name__ == "__main__":
    unittest.main()
