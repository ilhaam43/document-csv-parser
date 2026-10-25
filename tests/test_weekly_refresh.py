import unittest
from datetime import date, datetime
from decimal import Decimal

from weekly.catalyst import CatalystOverride
from weekly.matching import record_key
from weekly.models import ReviewDecision
from weekly.refresh import ExistingSourceIndex, refresh_existing_record
from weekly.rules import build_master_record


class WeeklyRefreshTests(unittest.TestCase):
    def test_missing_historical_results_use_the_same_formula_as_new_rows(self) -> None:
        holidays = {date(2031, 12, 25)}
        new = build_master_record({
            "quote_num": "Q-NEXT", "order_number": "O-NEXT",
            "FAB_Upload": "2031-12-24 09:15:00", "cash_date": "2031-12-26",
            "order_completion_date": "2031-12-29", "order_status": "Complete",
            "service_group": "Mobile", "otc": "1000.25", "mrc": "0.75",
        }, row_number=1, holidays=holidays)
        inherited = dict(new)
        inherited.update({"Count Cash": None, "LEADTIME FAB-CASH": None,
                          "OTC+MRC": None, "OTC+MRC (inMio)": None})
        refreshed = refresh_existing_record(
            inherited, source=None, catalyst_override=None, holidays=holidays
        )
        for field in ("Count Cash", "LEADTIME FAB-CASH", "OTC+MRC", "OTC+MRC (inMio)"):
            self.assertEqual(new[field], refreshed[field], field)

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

    def test_published_leadtimes_and_achievement_are_preserved_for_unchanged_dates(self) -> None:
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

    def test_revenue_is_recalculated_without_overwriting_published_leadtime(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({
            "cash_dt": datetime(2026, 9, 18),
            "otc": 1_000_000, "mrc": 25_000,
        })
        for prior_leadtime in (None, 0, 1, 999):
            with self.subTest(prior_leadtime=prior_leadtime):
                master.update({
                    "Count Cash": None, "LEADTIME FAB-CASH": prior_leadtime,
                    "OTC+MRC": None, "OTC+MRC (inMio)": 999,
                })
                before = dict(master)
                result = refresh_existing_record(
                    master, source=None, catalyst_override=None, holidays=set()
                )
                self.assertEqual(master, before)
                self.assertEqual(result["Count Cash"], 1)
                self.assertEqual(result["LEADTIME FAB-CASH"], 2 if prior_leadtime is None else prior_leadtime)
                self.assertEqual(result["OTC+MRC"], Decimal("1025000"))
                self.assertEqual(result["OTC+MRC (inMio)"], Decimal("1.025"))

    def test_holiday_updates_do_not_silently_rewrite_published_history(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master["cash_dt"] = datetime(2026, 9, 18)
        first = refresh_existing_record(
            master, source=None, catalyst_override=None, holidays=set()
        )
        second = refresh_existing_record(
            first, source=None, catalyst_override=None, holidays={date(2026, 9, 18)}
        )
        self.assertEqual(first["LEADTIME FAB-CASH"], 2)
        self.assertEqual(second["LEADTIME FAB-CASH"], 2)

    def test_published_weekend_normalization_is_preserved(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({
            "fab_upload": datetime(2026, 9, 20, 9),
            "cash_dt": datetime(2026, 9, 20), "LEADTIME FAB-CASH": 1,
        })
        result = refresh_existing_record(
            master, source=None, catalyst_override=None, holidays=set()
        )
        self.assertEqual(result["LEADTIME FAB-CASH"], 1)

    def test_missing_weekend_leadtime_is_calculated_as_zero(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({"fab_upload": datetime(2031, 12, 28, 9),
                       "cash_dt": datetime(2031, 12, 28)})
        result = refresh_existing_record(master, source=None, catalyst_override=None, holidays=set())
        self.assertEqual(result["LEADTIME FAB-CASH"], 0)

    def test_invalid_published_values_are_repaired_not_preserved(self) -> None:
        for invalid in (True, "#VALUE!", "NaN", "Infinity", 2.5, ""):
            with self.subTest(value=invalid):
                master = self.master(status="Complete", normalized="COMPLETE")
                master.update({"cash_dt": datetime(2026, 9, 18),
                               "order_completion_date": datetime(2026, 9, 21),
                               "LEADTIME FAB-CASH": invalid, "LEADTIME CASH-COMPLETE": invalid})
                result = refresh_existing_record(master, source=None, catalyst_override=None, holidays=set())
                self.assertEqual(result["LEADTIME FAB-CASH"], 2)
                self.assertEqual(result["LEADTIME CASH-COMPLETE"], 2)

    def test_new_cash_date_recalculates_affected_history_and_achievement(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({"order_completion_date": datetime(2026, 9, 21),
                       "LEADTIME FAB-CASH": 99, "LEADTIME CASH-COMPLETE": 99,
                       "ACHIEVEMENT": "NOT ACHIEVED"})
        result = refresh_existing_record(master, source=self.source(order="O-1", status="Complete"),
                                         catalyst_override=None, holidays=set())
        self.assertEqual(result["LEADTIME FAB-CASH"], 2)
        self.assertEqual(result["LEADTIME CASH-COMPLETE"], 2)
        self.assertEqual(result["ACHIEVEMENT"], "ACHIEVED")

    def test_changed_dates_invalidate_inherited_achievement_without_new_pm_approval(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({"order_completion_date": datetime(2026, 9, 25),
                       "LEADTIME CASH-COMPLETE": 1, "ACHIEVEMENT": "ACHIEVED",
                       "ADJUSTMENT": "Published adjustment"})
        result = refresh_existing_record(master, source=self.source(order="O-1", status="Complete"),
                                         catalyst_override=None, holidays=set())
        self.assertEqual(result["LEADTIME CASH-COMPLETE"], 6)
        self.assertEqual(result["ACHIEVEMENT"], "NOT ACHIEVED")

    def test_missing_cash_complete_is_calculated_even_when_dates_have_not_changed(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({"cash_dt": datetime(2026, 9, 18),
                       "order_completion_date": datetime(2026, 9, 25),
                       "LEADTIME CASH-COMPLETE": None, "ACHIEVEMENT": "ACHIEVED"})
        result = refresh_existing_record(master, source=None, catalyst_override=None, holidays=set())
        self.assertEqual(result["LEADTIME CASH-COMPLETE"], 6)
        self.assertEqual(result["ACHIEVEMENT"], "NOT ACHIEVED")

    def test_missing_achievement_uses_published_cash_complete_without_implying_pm_approval(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({"cash_dt": datetime(2026, 9, 18),
                       "order_completion_date": datetime(2026, 9, 21),
                       "LEADTIME CASH-COMPLETE": 6, "ADJUSTMENT": "Needs PM confirmation"})
        result = refresh_existing_record(master, source=None, catalyst_override=None, holidays=set())
        self.assertEqual(result["LEADTIME CASH-COMPLETE"], 6)
        self.assertEqual(result["ACHIEVEMENT"], "NOT ACHIEVED")

    def test_metric_audit_distinguishes_calculated_published_and_approved_values(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({"cash_dt": datetime(2026, 9, 18),
                       "order_completion_date": datetime(2026, 9, 21),
                       "LEADTIME FAB-CASH": 1, "LEADTIME CASH-COMPLETE": 6,
                       "ACHIEVEMENT": "NOT ACHIEVED"})
        audit = []
        result = refresh_existing_record(master, source=None, catalyst_override=None,
            review_decision=ReviewDecision(pm_approved=True), holidays=set(), metric_decisions=audit)
        metrics = {item["metric"]: item for item in audit}
        self.assertEqual(metrics["LEADTIME FAB-CASH"], {
            "metric": "LEADTIME FAB-CASH", "previous_value": 1,
            "calculated_value": 2, "final_value": 1, "decision_method": "published_history"})
        self.assertEqual(metrics["ACHIEVEMENT"]["decision_method"], "explicit_pm_approval")
        self.assertEqual(result["ACHIEVEMENT"], "ACHIEVED")

    def test_explicit_leadtime_correction_overrides_published_history(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({"cash_dt": datetime(2026, 9, 18), "LEADTIME FAB-CASH": 1})
        result = refresh_existing_record(master, source=None, catalyst_override=None,
            review_decision=ReviewDecision(leadtime_fab_cash=2), holidays=set())
        self.assertEqual(result["LEADTIME FAB-CASH"], 2)

    def test_repeated_refresh_does_not_change_a_completed_record(self) -> None:
        master = self.master(status="Complete", normalized="COMPLETE")
        master.update({
            "cash_dt": datetime(2026, 9, 18),
            "order_completion_date": datetime(2026, 9, 21),
        })
        source = self.source(order="O-1", status="Complete")
        once = refresh_existing_record(master, source=source, catalyst_override=None, holidays=set())
        twice = refresh_existing_record(once, source=source, catalyst_override=None, holidays=set())
        self.assertEqual(once, twice)

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
