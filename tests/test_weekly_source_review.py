import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from weekly.refresh import ExistingSourceIndex
from weekly.source_review import (
    automatic_source_choice,
    collect_source_conflicts,
    load_source_review_csv,
    source_with_reviewed_lifecycle,
    write_source_review_csv,
)


class WeeklySourceReviewTests(unittest.TestCase):
    def test_status_refresh_rule_matches_active_and_terminal_patterns(self) -> None:
        for status in ("Open", "In Progress", "#N/A", None):
            self.assertEqual(automatic_source_choice(status), "quote", status)
        for status in ("Complete", "COMPLETE", "Cancelled", "CANCEL"):
            self.assertEqual(automatic_source_choice(status), "order", status)

    def test_conflicts_are_limited_to_active_existing_and_new_rows(self) -> None:
        previous = [
            {"quote_num": "Q-1", "order number": "O-2", "order_status": "Open"},
            {"quote_num": "Q-2", "order number": "O-4", "order_status": "Complete"},
        ]
        index = ExistingSourceIndex(previous)
        for source in (
            {"quote_num": "Q-1", "order_number": "O-1", "order_status": "Complete", "order_completion_date": "9/30/2026"},
            {"quote_num": "Q-1", "order_number": "O-2", "order_status": "Open", "order_completion_date": ""},
            {"quote_num": "Q-2", "order_number": "O-3", "order_status": "Open", "order_completion_date": ""},
            {"quote_num": "Q-2", "order_number": "O-4", "order_status": "Complete", "order_completion_date": "9/30/2026"},
        ):
            index.add(source)
        selected = [
            {"quote_num": "Q-3", "order_number": "O-5", "order_status": "Complete", "order_completion_date": "9/30/2026"},
            {"quote_num": "Q-3", "order_number": "O-6", "order_status": "Open", "order_completion_date": ""},
        ]
        conflicts = collect_source_conflicts(previous, index, selected)
        self.assertEqual([row["master_row"] for row in conflicts], ["2", "5"])
        self.assertEqual([row["quote_status"] for row in conflicts], ["Complete", "Complete"])
        self.assertEqual([row["order_status"] for row in conflicts], ["Open", "Open"])

    def test_review_csv_checks_identity_and_accepts_optional_overrides(self) -> None:
        rows = [{
            "master_row": "5", "quote_num": "Q-3", "order_number": "O-6",
            "ca_name": "Customer", "service_group": "Mobile", "order_status": "Open",
            "order_completion_date": "", "quote_status": "Complete",
            "quote_completion_date": "9/30/2026", "reviewed_source": "",
        }]
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source-review.csv"
            write_source_review_csv(path, rows)
            self.assertEqual(load_source_review_csv(path, rows), {})
            second = dict(rows[0], master_row="6", quote_num="Q-4", order_number="O-7")
            reviewed = [dict(rows[0], reviewed_source="quote")]
            write_source_review_csv(path, reviewed)
            self.assertEqual(load_source_review_csv(path, [*rows, second]), {5: "quote"})
            write_source_review_csv(path, [rows[0], rows[0]])
            with self.assertRaisesRegex(ValueError, "duplicate"):
                load_source_review_csv(path, rows)
            write_source_review_csv(path, [dict(reviewed[0], order_number="O-7")])
            with self.assertRaisesRegex(ValueError, "identity changed"):
                load_source_review_csv(path, rows)

    def test_quote_choice_changes_only_lifecycle(self) -> None:
        order = {
            "quote_num": "Q-1", "order_number": "O-2", "order_status": "Open",
            "order_completion_date": "", "cash_date": "9/28/2026",
        }
        first = {
            "order_number": "O-1", "order_status": "Complete",
            "order_completion_date": "9/30/2026", "cash_date": "9/20/2026",
        }
        resolved = source_with_reviewed_lifecycle(order, first, "quote")
        self.assertEqual(resolved["order_number"], "O-2")
        self.assertEqual(resolved["cash_date"], "9/28/2026")
        self.assertEqual(resolved["order_status"], "Complete")
        self.assertEqual(resolved["order_completion_date"], "9/30/2026")


if __name__ == "__main__":
    unittest.main()
