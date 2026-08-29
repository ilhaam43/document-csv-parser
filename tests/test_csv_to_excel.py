import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace

import pandas as pd

from csv_to_excel import (
    PHASE_HEADER_PREFIX,
    PIVOT_TEMPLATE_SERVICE_DELIVERY_DIV_HEADER,
    QUO_HEADER,
    SERVICE_DELIVERY_DIV_HEADER,
    TARGET_SO_COMPLETE_DATE_HEADERS,
    add_pivot_template_compatibility_columns,
    apply_previous_target_date_fallback,
    target_header_month_name,
)


class ReportOneBusinessRuleTests(unittest.TestCase):
    def test_target_header_month_accepts_short_and_long_year_formats(self) -> None:
        self.assertEqual(target_header_month_name("TARGET  Detemined as 1 Aug 26"), "August")
        self.assertEqual(target_header_month_name("TARGET Determined as 1 August 2026"), "August")

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
