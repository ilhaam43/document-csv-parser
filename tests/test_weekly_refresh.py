import unittest
from datetime import date, datetime

from weekly.catalyst import CatalystOverride
from weekly.matching import record_key
from weekly.models import ReviewDecision
from weekly.refresh import ExistingSourceIndex, refresh_existing_record


class WeeklyRefreshTests(unittest.TestCase):
    def test_missing_cash_clears_sla_but_retains_valid_completion_period(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({"order_completion_date": datetime(2026, 9, 21),
                       "Count Completion": 1, "Week Completion": "W38",
                       "LEADTIME CASH-COMPLETE": 5, "ACHIEVEMENT": "NOT ACHIEVED"})
        result = refresh_existing_record(master, source=None, catalyst_override=None, holidays=set())
        self.assertIsNone(result["LEADTIME CASH-COMPLETE"])
        self.assertIsNone(result["ACHIEVEMENT"])
        self.assertEqual(result["Count Completion"], 1)
        self.assertEqual(result["Week Completion"], "W38")

    def test_missing_fab_clears_lead_time_without_erasing_cash_count(self) -> None:
        master = self.master(status="Open", normalized="In Progress")
        master.update({"fab_upload": None, "cash_dt": datetime(2026, 9, 18),
                       "Count Cash": 1, "LEADTIME FAB-CASH": 5})
        result = refresh_existing_record(master, source=None, catalyst_override=None, holidays=set())
        self.assertIsNone(result["LEADTIME FAB-CASH"])
        self.assertEqual(result["Count Cash"], 1)

    def test_missing_dates_clear_inherited_derived_values_without_a_date_change(self) -> None:
        master = self.master(status="In Progress", normalized="In Progress")
        master.update({
            "Count Cash": 1, "Count Completion": 1,
            "Week Completion": "W35", "Month Completion": "Aug-26", "Year Completion": "2026",
            "LEADTIME FAB-CASH": 5, "LEADTIME CASH-COMPLETE": 5,
            "ACHIEVEMENT": "NOT ACHIEVED",
        })
        result = refresh_existing_record(
            master, source=None, catalyst_override=None, holidays=set()
        )
        for field in (
            "Count Cash", "Count Completion", "Week Completion", "Month Completion",
            "Year Completion", "LEADTIME FAB-CASH", "LEADTIME CASH-COMPLETE", "ACHIEVEMENT",
        ):
            self.assertIsNone(result[field], field)

    def test_valid_historical_reviewed_values_survive_unchanged_dates(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({
            "cash_dt": datetime(2026, 9, 18),
            "order_completion_date": datetime(2026, 9, 25),
            "LEADTIME FAB-CASH": 0, "LEADTIME CASH-COMPLETE": 6,
            "ACHIEVEMENT": "ACHIEVED", "ADJUSTMENT": "PM confirmed",
        })
        result = refresh_existing_record(
            master, source=None, catalyst_override=None, holidays=set()
        )
        self.assertEqual(result["LEADTIME FAB-CASH"], 0)
        self.assertEqual(result["LEADTIME CASH-COMPLETE"], 6)
        self.assertEqual(result["ACHIEVEMENT"], "ACHIEVED")

    def test_existing_approved_adjustment_is_preserved_without_completion_date(self) -> None:
        master = self.master(status="In Progress", normalized="In Progress")
        master.update({"ADJUSTMENT": "Antrian Massloader", "ACHIEVEMENT": "ACHIEVED"})
        result = refresh_existing_record(master, source=None, catalyst_override=None, holidays=set())
        self.assertEqual(result["ACHIEVEMENT"], "ACHIEVED")
        master["ACHIEVEMENT"] = "NOT ACHIEVED"
        result = refresh_existing_record(master, source=None, catalyst_override=None, holidays=set())
        self.assertIsNone(result["ACHIEVEMENT"])

    def test_terminal_status_is_not_downgraded_by_bq(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        result = refresh_existing_record(
            master,
            source=self.source(order="O-1", status="Open"),
            catalyst_override=None,
            holidays=set(),
        )
        self.assertEqual(result["order_status"], "Complete")
        self.assertEqual(result["order_status2"], "COMPLETE")

    def test_missing_bq_match_marks_active_status_pending_without_changing_pivot_status(self) -> None:
        master = self.master(status="In Progress", normalized="In Progress")
        result = refresh_existing_record(
            master, source=None, catalyst_override=None, holidays=set()
        )
        self.assertEqual(result["order_status"], "#N/A")
        self.assertEqual(result["order_status2"], "In Progress")
        self.assertEqual(result["order_status (Confirmed by MIDO Dept.)"], "In Progress")

    def test_previous_missing_status_without_bq_match_stays_pending(self) -> None:
        master = self.master(status="#N/A", normalized="In Progress")
        result = refresh_existing_record(
            master, source=None, catalyst_override=None, holidays=set()
        )
        self.assertEqual(result["order_status"], "#N/A")
        self.assertEqual(result["order_status2"], "In Progress")

    def test_unresolved_review_preserves_prior_pivot_status(self) -> None:
        master = self.master(status="Open", normalized="In Progress")
        result = refresh_existing_record(
            master,
            source=None,
            catalyst_override=CatalystOverride(record_key("Q-1", "O-1"), "#N/A"),
            holidays=set(),
        )
        self.assertEqual(result["order_status"], "#N/A")
        self.assertEqual(result["order_status2"], "In Progress")

    def test_bq_resolves_previous_missing_status(self) -> None:
        master = self.master(status="#N/A", normalized="In Progress")
        result = refresh_existing_record(
            master,
            source=self.source(order="O-1", status="Complete"),
            catalyst_override=None,
            holidays=set(),
        )
        self.assertEqual(result["order_status"], "Complete")
        self.assertEqual(result["order_status2"], "COMPLETE")

    def test_matched_empty_bq_lifecycle_resolves_previous_missing_status(self) -> None:
        master = self.master(status="#N/A", normalized="In Progress")
        source = self.source(order="O-1", status="")
        source["order_completion_date"] = ""
        result = refresh_existing_record(
            master, source=source, catalyst_override=None, holidays=set()
        )
        self.assertEqual(result["order_status"], "In Progress")
        self.assertEqual(result["order_status2"], "In Progress")
        self.assertEqual(result["order_status (Confirmed by MIDO Dept.)"], "In Progress")

    def test_blank_bq_status_with_completion_date_stays_unresolved(self) -> None:
        master = self.master(status="#N/A", normalized="In Progress")
        result = refresh_existing_record(
            master,
            source=self.source(order="O-1", status=""),
            catalyst_override=None,
            holidays=set(),
        )
        self.assertEqual(result["order_status"], "#N/A")

    def test_reviewed_status_overrides_matched_empty_bq_lifecycle(self) -> None:
        master = self.master(status="#N/A", normalized="In Progress")
        source = self.source(order="O-1", status="")
        source["order_completion_date"] = ""
        result = refresh_existing_record(
            master,
            source=source,
            catalyst_override=None,
            review_decision=ReviewDecision(order_status="Cancelled"),
            holidays=set(),
        )
        self.assertEqual(result["order_status"], "Cancelled")
        self.assertEqual(result["order_status2"], "CANCEL")
        self.assertEqual(result["order_status (Confirmed by MIDO Dept.)"], "CANCEL")

    def test_exact_order_match_wins_for_existing_order(self) -> None:
        master = {"quote_num": "Q-1", "order number": "O-2"}
        index = ExistingSourceIndex([master])
        index.add(self.source(order="O-1", status="Cancelled"))
        index.add(self.source(order="O-2", status="Complete"))
        source, issue = index.resolve(master)
        self.assertEqual(source["order_number"], "O-2")
        self.assertIsNone(issue)

    def test_reviewed_active_status_clears_completion_and_derived_fields(self) -> None:
        master = self.master(status="Open", normalized="In Progress")
        master.update({
            "order_completion_date": datetime(2026, 9, 21),
            "Count Completion": 1,
            "Week Completion": "W39",
            "Month Completion": "Sep-26",
            "Year Completion": "2026",
            "LEADTIME CASH-COMPLETE": 2,
            "ACHIEVEMENT": "ACHIEVED",
        })
        result = refresh_existing_record(
            master,
            source=self.source(order="O-1", status="Complete"),
            catalyst_override=None,
            review_decision=ReviewDecision(order_status="Open"),
            holidays=set(),
        )
        for column in (
            "order_completion_date", "Count Completion", "Week Completion",
            "Month Completion", "Year Completion", "LEADTIME CASH-COMPLETE", "ACHIEVEMENT",
        ):
            self.assertIsNone(result[column], column)
        self.assertEqual(result["order_status2"], "In Progress")

    def test_reviewed_cancel_uses_manual_raw_label(self) -> None:
        result = refresh_existing_record(
            self.master(status="#N/A", normalized="In Progress"),
            source=None,
            catalyst_override=None,
            review_decision=ReviewDecision(order_status="CANCEL"),
            holidays=set(),
        )
        self.assertEqual(result["order_status"], "Cancelled")
        self.assertEqual(result["order_status2"], "CANCEL")

    def test_cancelled_order_does_not_keep_stale_completion_metrics(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({
            "order_completion_date": datetime(2026, 9, 21),
            "LEADTIME CASH-COMPLETE": 5,
            "ACHIEVEMENT": "NOT ACHIEVED",
        })
        result = refresh_existing_record(
            master,
            source=None,
            catalyst_override=None,
            review_decision=ReviewDecision(order_status="Cancelled"),
            holidays=set(),
        )
        self.assertEqual(result["order_status2"], "CANCEL")
        self.assertIsNone(result["order_completion_date"])
        self.assertIsNone(result["LEADTIME CASH-COMPLETE"])
        self.assertIsNone(result["ACHIEVEMENT"])

    def test_blank_order_uses_first_quote_match_and_reports_ambiguity(self) -> None:
        master = {"quote_num": "Q-1", "order number": None}
        index = ExistingSourceIndex([master])
        index.add(self.source(order="O-1", status="Complete"))
        index.add(self.source(order="O-2", status="Cancelled"))
        source, issue = index.resolve(master)
        self.assertEqual(source["order_number"], "O-1")
        self.assertEqual(issue.code, "quote_first_match_used")

    def test_catalyst_status_and_source_lifecycle_are_separate(self) -> None:
        master = {
            "quote_num": "Q-1",
            "order number": None,
            "order_status": "#N/A",
            "order_status2": "In Progress",
            "order_status (Confirmed by MIDO Dept.)": "In Progress",
            "service_group": "Mobile",
            "fab_upload": datetime(2026, 9, 17, 10, 0),
            "cash_dt": None,
            "order created": None,
            "order_completion_date": None,
            "otc": 0,
            "mrc": 25_000,
            "ADJUSTMENT": None,
        }
        source = self.source(order="O-1", status="Complete")
        override = CatalystOverride(record_key("Q-1", None), "Complete")
        result = refresh_existing_record(
            master,
            source=source,
            catalyst_override=override,
            holidays=set(),
        )
        self.assertEqual(result["order_status"], "Complete")
        self.assertEqual(result["order_status2"], "COMPLETE")
        self.assertEqual(result["order number"], "O-1")
        self.assertEqual(result["LEADTIME CASH-COMPLETE"], 2)
        self.assertEqual(result["ACHIEVEMENT"], "ACHIEVED")

    @staticmethod
    def master(*, status: str, normalized: str) -> dict:
        return {
            "quote_num": "Q-1",
            "order number": "O-1",
            "order_status": status,
            "order_status2": normalized,
            "order_status (Confirmed by MIDO Dept.)": normalized,
            "service_group": "Mobile",
            "fab_upload": datetime(2026, 9, 17),
            "cash_dt": None,
            "order_completion_date": None,
            "otc": 0,
            "mrc": 25_000,
        }

    @staticmethod
    def source(*, order: str, status: str) -> dict[str, str]:
        return {
            "quote_num": "Q-1",
            "order_number": order,
            "order_status": status,
            "cash_date": "9/18/2026",
            "order_crea": "9/20/2026",
            "order_completion_date": "9/21/2026",
        }


if __name__ == "__main__":
    unittest.main()
