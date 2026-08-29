from datetime import date
from pathlib import Path
from types import SimpleNamespace
import unittest

import csv_to_excel_on_going
import run_daily_pipeline


class RunDailyPipelineTests(unittest.TestCase):
    def test_ongoing_aging_defaults_to_generated_report_date(self) -> None:
        output_path = Path("Daily Tracking 28 August 2026.xlsx")

        self.assertEqual(
            run_daily_pipeline.default_ongoing_aging_date(output_path),
            date(2026, 8, 28),
        )

    def test_column1_is_treated_as_division_sales_legacy_alias(self) -> None:
        self.assertEqual(
            csv_to_excel_on_going.canonical_sales_hierarchy_header_key("Column1"),
            "divisionsales",
        )

    def test_existing_on_progress_tab_color_is_preserved(self) -> None:
        worksheet = SimpleNamespace(Tab=SimpleNamespace(Color=123456))
        csv_to_excel_on_going.preserve_or_set_on_progress_tab_color(worksheet)

        self.assertEqual(worksheet.Tab.Color, 123456)
        self.assertEqual(
            csv_to_excel_on_going.canonical_sales_hierarchy_header_key("Division_Sales"),
            "divisionsales",
        )


if __name__ == "__main__":
    unittest.main()
