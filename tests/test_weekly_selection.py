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

    def test_older_fab_orders_with_current_cash_or_completion_are_backfilled(self) -> None:
        rows = [
            dict(self.row("CASH", "Mobile", "9/30/2026 16:15"), cash_date="10/2/2026"),
            self.row("CURRENT", "Mobile", "10/20/2026"),
            dict(self.row("COMPLETE", "Bulk Service", "8/15/2026"),
                 cash_date="8/16/2026", order_completion_date="10/5/2026"),
        ]
        result = select_new_source_rows(rows, existing_quote_nums=set(), report_date=date(2026, 10, 23))
        self.assertEqual([row["quote_num"] for row in result.selected_rows], ["CASH", "CURRENT", "COMPLETE"])
        self.assertEqual(result.historical_backfill_rows, 2)
        self.assertEqual(result.historical_backfill_by_fab_month, {"2026-09": 1, "2026-08": 1})
        self.assertEqual(result.selected_rows[0]["FAB_Upload"], "9/30/2026 16:15")
        self.assertEqual(rows[0]["FAB_Upload"], "9/30/2026 16:15")

    def test_backfill_keeps_existing_quote_service_fab_order_and_activity_guards(self) -> None:
        rows = [
            dict(self.row("KNOWN", "Mobile", "9/30/2026"), cash_date="10/2/2026"),
            dict(self.row("FIXED", "Fixed Data", "9/30/2026"), cash_date="10/2/2026"),
            dict(self.row("NO-FAB", "Mobile", ""), cash_date="10/2/2026"),
            dict(self.row("NO-ORDER", "Mobile", "9/30/2026", order=""), cash_date="10/2/2026"),
            dict(self.row("NO-ACTIVITY", "Mobile", "9/30/2026"), cash_date="9/30/2026"),
            dict(self.row("FUTURE-ACTIVITY", "Mobile", "9/30/2026"), cash_date="10/24/2026"),
            dict(self.row("FUTURE-FAB", "Mobile", "11/1/2026"), cash_date="10/2/2026"),
        ]
        result = select_new_source_rows(rows, existing_quote_nums={"KNOWN"}, report_date=date(2026, 10, 23))
        self.assertEqual(result.selected_rows, [])
        self.assertEqual(result.historical_backfill_rows, 0)
        self.assertEqual(result.rejected_reasons, {
            "existing_quote_num": 1, "service_group_not_in_scope": 1,
            "missing_fab_upload": 1, "outside_reporting_month": 4,
        })

    def test_backfill_uses_reporting_month_not_export_month(self) -> None:
        rows = [
            dict(self.row("SEP-ACTIVITY", "M2M", "8/31/2026"), order_completion_date="9/30/2026"),
            dict(self.row("OCT-ACTIVITY", "Mobile", "8/31/2026"), cash_date="10/1/2026"),
        ]
        result = select_new_source_rows(rows, existing_quote_nums=set(), report_date=date(2026, 10, 2),
                                        report_month=date(2026, 9, 1))
        self.assertEqual([r["quote_num"] for r in result.selected_rows], ["SEP-ACTIVITY"])
        self.assertEqual(result.historical_backfill_by_fab_month, {"2026-08": 1})

    def test_backfill_handles_year_boundary_without_a_fixed_lookback(self) -> None:
        rows = [dict(self.row("LATE", "Mobile", "12/31/2025"), cash_date="1/1/2027")]
        result = select_new_source_rows(rows, existing_quote_nums=set(), report_date=date(2027, 1, 1))
        self.assertEqual(result.selected_rows, rows)
        self.assertEqual(result.historical_backfill_by_fab_month, {"2025-12": 1})

    def test_backfill_cutoff_is_inclusive_and_does_not_use_future_completion(self) -> None:
        rows = [
            dict(self.row("AT-CUTOFF", "Mobile", "9/30/2026"), order_completion_date="10/23/2026 23:59:59"),
            dict(self.row("AFTER-CUTOFF", "Mobile", "9/30/2026"), order_completion_date="10/24/2026"),
        ]
        result = select_new_source_rows(rows, existing_quote_nums=set(), report_date=date(2026, 10, 23))
        self.assertEqual([r["quote_num"] for r in result.selected_rows], ["AT-CUTOFF"])

    def test_invalid_backfill_activity_is_reported_instead_of_guessed(self) -> None:
        rows = [dict(self.row("INVALID", "Mobile", "9/30/2026"),
                     cash_date="10/2/2026", order_completion_date="not-a-date")]
        result = select_new_source_rows(rows, existing_quote_nums=set(), report_date=date(2026, 10, 23))
        self.assertEqual(result.selected_rows, [])
        self.assertEqual(result.rejected_reasons["invalid_backfill_activity"], 1)
        self.assertEqual([(i.code, i.row_number) for i in result.issues], [("invalid_backfill_activity", 2)])

    def test_backfill_is_idempotent_when_selected_quote_is_in_next_master(self) -> None:
        rows = [dict(self.row("LATE", "Mobile", "9/30/2026"), cash_date="10/2/2026")]
        first = select_new_source_rows(rows, existing_quote_nums=set(), report_date=date(2026, 10, 23))
        second = select_new_source_rows(rows, existing_quote_nums={r["quote_num"] for r in first.selected_rows},
                                        report_date=date(2026, 10, 30))
        self.assertEqual(len(first.selected_rows), 1)
        self.assertEqual(second.selected_rows, [])
        self.assertEqual(second.historical_backfill_rows, 0)

    def test_backfill_keeps_distinct_order_rows_for_one_new_quote(self) -> None:
        rows = [dict(self.row("LATE", "Mobile", "9/30/2026", order=order), cash_date="10/2/2026")
                for order in ("FIRST", "SECOND")]
        result = select_new_source_rows(rows, existing_quote_nums=set(), report_date=date(2026, 10, 23))
        self.assertEqual([r["order_number"] for r in result.selected_rows], ["FIRST", "SECOND"])
        self.assertEqual(result.historical_backfill_rows, 2)

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
