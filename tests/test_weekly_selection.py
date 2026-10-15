import csv
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from datetime import date

from weekly.constants import SOURCE_COLUMNS
from weekly.selection import iter_bq_rows, select_new_source_rows


class WeeklySelectionTests(unittest.TestCase):
    def _read_csv(self, headers, values):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "bq.csv"
            with path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow(headers)
                writer.writerow(values)
            return list(iter_bq_rows(path))

    def test_original_bq_headers_preserve_source_values(self) -> None:
        values = [f"value-{index}" for index in range(len(SOURCE_COLUMNS))]
        rows = self._read_csv(SOURCE_COLUMNS, values)
        self.assertEqual(rows, [dict(zip(SOURCE_COLUMNS, values))])

    def test_created_header_aliases_map_to_the_existing_contract_without_changing_values(self) -> None:
        headers = [
            {"quote_crea": "quote_created", "order_crea": "order_created"}.get(c, c)
            for c in SOURCE_COLUMNS
        ]
        values = [f"value-{index}" for index in range(len(headers))]
        rows = self._read_csv(headers, values)
        self.assertEqual(rows, [dict(zip(SOURCE_COLUMNS, values))])
        self.assertNotIn("order_created", rows[0])

    def test_conflicting_alias_columns_are_rejected_instead_of_overwritten(self) -> None:
        for alias in ("order_created", "quote_created"):
            with self.subTest(alias=alias), self.assertRaisesRegex(ValueError, "duplicate source columns"):
                self._read_csv([*SOURCE_COLUMNS, alias], [""] * (len(SOURCE_COLUMNS) + 1))

    def test_a_missing_required_source_field_is_still_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "missing required columns: order_number"):
            self._read_csv([c for c in SOURCE_COLUMNS if c != "order_number"], [""] * (len(SOURCE_COLUMNS) - 1))

    def test_filters_by_existing_quote_service_month_and_cutoff(self) -> None:
        rows = [
            self.row("NEW-1", "Mobile", "9/20/2026"),
            self.row("NEW-1", "Mobile", "9/21/2026", order="O-2"),
            self.row("OLD-1", "Mobile", "9/20/2026"),
            self.row("NEW-2", "Fixed Data", "9/20/2026"),
            self.row("NEW-3", "M2M", "8/31/2026"),
            self.row("NEW-4", "Bulk Service", "9/25/2026"),
        ]
        result = select_new_source_rows(
            rows,
            existing_quote_nums={"OLD-1"},
            report_date=date(2026, 9, 24),
        )
        self.assertEqual([row["order_number"] for row in result.selected_rows], ["O-1", "O-2"])
        self.assertEqual(result.rejected_reasons["existing_quote_num"], 1)
        self.assertEqual(result.rejected_reasons["service_group_not_in_scope"], 1)
        self.assertEqual(result.rejected_reasons["outside_reporting_month"], 1)
        self.assertEqual(result.rejected_reasons["after_report_cutoff"], 1)

    def test_report_month_can_precede_bq_cutoff_month(self) -> None:
        rows = [
            self.row("SEP", "Mobile", "9/30/2026"),
            self.row("OCT", "Mobile", "10/1/2026"),
        ]
        result = select_new_source_rows(
            rows, existing_quote_nums=set(), report_date=date(2026, 10, 2),
            report_month=date(2026, 9, 1),
        )
        self.assertEqual([row["quote_num"] for row in result.selected_rows], ["SEP"])

    @staticmethod
    def row(quote: str, service: str, fab: str, *, order: str = "O-1") -> dict[str, str]:
        return {
            "quote_num": quote,
            "service_group": service,
            "FAB_Upload": fab,
            "order_number": order,
        }


if __name__ == "__main__":
    unittest.main()
