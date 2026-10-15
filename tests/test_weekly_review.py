import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from weekly.models import ReviewDecision
from weekly.refresh import refresh_existing_record
from weekly.review import load_review_decisions


class WeeklyReviewTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.path = Path(self.temporary.name) / "review.csv"
        self.rows = [
            {"quote_num": "Q-1", "order number": None},
            {"quote_num": "Q-1", "order number": None},
            {"quote_num": "Q-2", "order number": "O-2"},
        ]

    def _write(self, rows: str) -> None:
        self.path.write_text(
            "master_row,quote_num,order_number,reviewed_status,"
            "reviewed_completion_date,pm_approved,reviewed_by,review_reason,"
            "reviewed_fab_cash_leadtime\n" + rows,
            encoding="utf-8-sig",
        )

    def test_review_targets_duplicate_quote_rows_independently(self) -> None:
        self._write(
            "2,Q-1,,Cancelled,,,Dewi,Catalyst verification\n"
            "3,Q-1,,Cancelled,,,Dewi,Catalyst verification\n"
            "4,Q-2,O-2,,2026-09-09,yes,Dewi,PM confirmed correction\n"
        )
        result = load_review_decisions(self.path, self.rows)
        decisions = result.decisions
        self.assertEqual(result.count, 3)
        self.assertEqual(decisions[2].order_status, "Cancelled")
        self.assertEqual(decisions[3].order_status, "Cancelled")
        self.assertEqual(decisions[4].completion_date, datetime(2026, 9, 9))
        self.assertTrue(decisions[4].pm_approved)

    def test_checkpoint_matches_new_row_without_guessing_duplicate_keys(self) -> None:
        new_row = {
            "quote_num": "Q-3",
            "order number": "O-3",
            "fab_upload": datetime(2026, 9, 20, 22, 43),
        }
        reviewed = {
            **new_row,
            "fab_upload": datetime(2026, 9, 20, 22, 43, 16),
            "No.": 53922.1666662631,
            "LEADTIME FAB-CASH": 0,
            "order_status": "Complete",
            "order_completion_date": datetime(2026, 9, 21),
        }
        with patch("weekly.review.iter_catalyst_review", return_value=iter([reviewed])):
            result = load_review_decisions(
                self.path.with_suffix(".xlsx"), [new_row], first_new_row=2
            )
        self.assertEqual(result.unmatched_new_rows, 0)
        self.assertEqual(result.checkpoint_rows[2].fab_upload.second, 16)
        self.assertEqual(result.checkpoint_rows[2].leadtime_fab_cash, 0)
        self.assertEqual(result.decisions[2].completion_date, datetime(2026, 9, 21))

        with patch("weekly.review.iter_catalyst_review", return_value=iter([reviewed, reviewed])):
            result = load_review_decisions(
                self.path.with_suffix(".xlsx"), [new_row], first_new_row=2
            )
        self.assertEqual(result.unmatched_new_rows, 1)
        self.assertFalse(result.checkpoint_rows)

    def test_checkpoint_matches_equal_duplicate_occurrences_in_order(self) -> None:
        draft = [
            {"quote_num": "Q-3", "order number": "O-3", "fab_upload": datetime(2026, 9, 29, 9, 39)},
            {"quote_num": "Q-3", "order number": "O-3", "fab_upload": datetime(2026, 9, 29, 9, 39)},
        ]
        checkpoint = [
            {**draft[0], "No.": 10.1, "fab_upload": datetime(2026, 9, 29, 9, 39, 18), "order_status": "Open"},
            {**draft[1], "No.": 11.2, "fab_upload": datetime(2026, 9, 29, 9, 39, 19), "order_status": "Complete"},
        ]
        with patch("weekly.review.iter_catalyst_review", return_value=iter(checkpoint)):
            result = load_review_decisions(self.path.with_suffix(".xlsx"), draft, first_new_row=2)
        self.assertEqual(result.unmatched_new_rows, 0)
        self.assertEqual(result.checkpoint_rows[2].fab_upload.second, 18)
        self.assertEqual(result.checkpoint_rows[3].fab_upload.second, 19)
        self.assertEqual(result.decisions[3].order_status, "Complete")

    def test_checkpoint_does_not_require_a_manual_number(self) -> None:
        draft = [{"quote_num": "Q-3", "order number": "O-3", "fab_upload": datetime(2026, 9, 29, 9, 39)}]
        checkpoint = [{**draft[0], "No.": None, "order_status": "Complete"}]
        with patch("weekly.review.iter_catalyst_review", return_value=iter(checkpoint)):
            result = load_review_decisions(self.path.with_suffix(".xlsx"), draft, first_new_row=2)
        self.assertEqual(result.unmatched_new_rows, 0)
        self.assertIn(2, result.checkpoint_rows)

    def test_legacy_checkpoint_matches_existing_rows_after_order_number_changes(self) -> None:
        draft = [
            {"quote_num": "Q-1", "order number": "O-1", "fab_upload": datetime(2026, 9, 1), "order_status": "#N/A"},
            {"quote_num": "Q-1", "order number": "O-2", "fab_upload": datetime(2026, 9, 2), "order_status": "#N/A"},
        ]
        checkpoint = [
            {"quote_num": "Q-1", "order number": None, "fab_upload": datetime(2026, 9, 1), "order_status": "In Progress"},
            {"quote_num": "Q-1", "order number": None, "fab_upload": datetime(2026, 9, 2), "order_status": "Complete", "order_completion_date": datetime(2026, 9, 9)},
        ]
        with patch("weekly.review.iter_catalyst_review", return_value=iter(checkpoint)):
            result = load_review_decisions(self.path.with_suffix(".xlsx"), draft, first_new_row=4)
        self.assertEqual(result.decisions[2].order_status, "In Progress")
        self.assertEqual(result.decisions[3].order_status, "Complete")
        self.assertEqual(result.decisions[3].completion_date, datetime(2026, 9, 9))
        checkpoint[1]["quote_num"] = "different"
        with patch("weekly.review.iter_catalyst_review", return_value=iter(checkpoint)):
            with self.assertRaisesRegex(ValueError, "does not match"):
                load_review_decisions(self.path.with_suffix(".xlsx"), draft, first_new_row=4)

    def test_checkpoint_explicit_achievement_change_is_approval(self) -> None:
        draft = [{
            "quote_num": "Q-1", "order number": "O-1",
            "fab_upload": datetime(2026, 9, 1),
            "ACHIEVEMENT": "NOT ACHIEVED", "ADJUSTMENT": "PM note",
        }]
        checkpoint = [{**draft[0], "ACHIEVEMENT": "ACHIEVED"}]
        with patch("weekly.review.iter_catalyst_review", return_value=iter(checkpoint)):
            result = load_review_decisions(self.path.with_suffix(".xlsx"), draft, first_new_row=3)
        self.assertTrue(result.decisions[2].pm_approved)
        checkpoint[0]["ACHIEVEMENT"] = "NOT ACHIEVED"
        with patch("weekly.review.iter_catalyst_review", return_value=iter(checkpoint)):
            result = load_review_decisions(self.path.with_suffix(".xlsx"), draft, first_new_row=3)
        self.assertFalse(result.decisions)

    def test_rejects_mismatched_identity_and_ambiguous_missing_row(self) -> None:
        self._write("2,Q-2,O-2,Cancelled,,,Dewi,Catalyst verification\n")
        with self.assertRaisesRegex(ValueError, "does not match"):
            load_review_decisions(self.path, self.rows)
        self._write(",Q-1,,Cancelled,,,Dewi,Catalyst verification\n")
        with self.assertRaisesRegex(ValueError, "master_row is required"):
            load_review_decisions(self.path, self.rows)

    def test_rejects_unattributed_or_invalid_corrections(self) -> None:
        self._write("4,Q-2,O-2,,2026-09-09,yes,,\n")
        with self.assertRaisesRegex(ValueError, "reviewed_by and review_reason"):
            load_review_decisions(self.path, self.rows, require_attribution=True)
        self._write("4,Q-2,O-2,,09/09/2026,,Dewi,PM confirmation\n")
        with self.assertRaisesRegex(ValueError, "YYYY-MM-DD"):
            load_review_decisions(self.path, self.rows)
        self._write("4,Q-2,O-2,,,approved,Dewi,PM confirmation\n")
        with self.assertRaisesRegex(ValueError, "pm_approved must"):
            load_review_decisions(self.path, self.rows)

    def test_catalyst_review_accepts_status_without_attribution(self) -> None:
        self.path.write_text(
            "master_row,quote_num,order_number,reviewed_status\n"
            "2,Q-1,,In Progress\n",
            encoding="utf-8-sig",
        )
        result = load_review_decisions(self.path, self.rows)
        self.assertEqual(result.decisions[2].order_status, "In Progress")
        self.assertEqual(result.decisions[2].reviewed_by, "")
        self.assertEqual(result.decisions[2].review_reason, "")

    def test_adjustment_note_does_not_approve_without_explicit_decision(self) -> None:
        row = {
            "quote_num": "Q-2",
            "order number": "O-2",
            "order_status": "Complete",
            "service_group": "Mobile",
            "fab_upload": datetime(2026, 9, 1),
            "cash_dt": datetime(2026, 9, 2),
            "order_completion_date": datetime(2026, 9, 10),
            "LEADTIME CASH-COMPLETE": 7,
            "ACHIEVEMENT": "NOT ACHIEVED",
            "ADJUSTMENT": "Customer requested later activation",
            "otc": 0,
            "mrc": 0,
        }
        unchanged = refresh_existing_record(
            row, source=None, catalyst_override=None, holidays=set()
        )
        self.assertEqual(unchanged["ACHIEVEMENT"], "NOT ACHIEVED")
        approved = refresh_existing_record(
            row,
            source=None,
            catalyst_override=None,
            review_decision=ReviewDecision(pm_approved=True),
            holidays=set(),
        )
        self.assertEqual(approved["ACHIEVEMENT"], "ACHIEVED")

    def test_corrected_completion_recalculates_week_leadtime_and_achievement(self) -> None:
        row = {
            "quote_num": "Q-2",
            "order number": "O-2",
            "order_status": "Complete",
            "service_group": "Mobile",
            "fab_upload": datetime(2026, 9, 1),
            "cash_dt": datetime(2026, 9, 2),
            "order_completion_date": datetime(2026, 9, 10),
            "LEADTIME CASH-COMPLETE": 7,
            "ACHIEVEMENT": "NOT ACHIEVED",
            "otc": 0,
            "mrc": 0,
        }
        corrected = refresh_existing_record(
            row,
            source={"order_status": "Complete"},
            catalyst_override=None,
            review_decision=ReviewDecision(completion_date=datetime(2026, 9, 9)),
            holidays=set(),
        )
        self.assertEqual(corrected["order_completion_date"], datetime(2026, 9, 9))
        self.assertEqual(corrected["Week Completion"], "W36")
        self.assertEqual(corrected["LEADTIME CASH-COMPLETE"], 6)
        self.assertEqual(corrected["ACHIEVEMENT"], "NOT ACHIEVED")

    def test_reviewed_zero_leadtime_requires_matching_source_dates(self) -> None:
        row = {
            "quote_num": "Q-2",
            "order number": "O-2",
            "order_status": "Complete",
            "service_group": "Mobile",
            "fab_upload": datetime(2026, 9, 20, 22, 43),
            "cash_dt": datetime(2026, 9, 20),
            "order_completion_date": datetime(2026, 9, 21),
            "LEADTIME FAB-CASH": 1,
            "otc": 0,
            "mrc": 0,
        }
        reviewed = refresh_existing_record(
            row,
            source=None,
            catalyst_override=None,
            review_decision=ReviewDecision(leadtime_fab_cash=0),
            holidays=set(),
        )
        self.assertEqual(reviewed["LEADTIME FAB-CASH"], 0)
        row["cash_dt"] = datetime(2026, 9, 21)
        with self.assertRaisesRegex(ValueError, "does not match"):
            refresh_existing_record(
                row,
                source=None,
                catalyst_override=None,
                review_decision=ReviewDecision(leadtime_fab_cash=0),
                holidays=set(),
            )

    def test_zero_leadtime_is_parsed_as_a_correction(self) -> None:
        self._write("4,Q-2,O-2,,,,Dewi,Manual final confirmed,0\n")
        result = load_review_decisions(self.path, self.rows)
        self.assertEqual(result.decisions[4].leadtime_fab_cash, 0)


if __name__ == "__main__":
    unittest.main()
