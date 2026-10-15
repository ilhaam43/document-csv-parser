import tempfile
import unittest
from pathlib import Path

from weekly.catalyst import collect_catalyst_overrides, iter_catalyst_review
from weekly.matching import MasterIndex


class WeeklyMatchingTests(unittest.TestCase):
    def test_composite_key_resolves_duplicate_quote(self) -> None:
        index = MasterIndex(
            [
                {"quote_num": "Q-1", "order number": "O-1"},
                {"quote_num": "Q-1", "order number": "O-2"},
            ]
        )
        match = index.resolve({"quote_num": "Q-1", "order_number": "O-2"})
        self.assertEqual(match.row_number, 3)
        self.assertEqual(match.strategy, "quote_and_order")

    def test_ambiguous_quote_is_not_silently_matched(self) -> None:
        index = MasterIndex(
            [
                {"quote_num": "Q-1", "order number": "O-1"},
                {"quote_num": "Q-1", "order number": "O-2"},
            ]
        )
        match = index.resolve({"quote_num": "Q-1", "order_number": "O-3"})
        self.assertIsNone(match.row_number)
        self.assertEqual(match.issue.code, "ambiguous_quote")

    def test_conflicting_catalyst_status_is_rejected(self) -> None:
        overrides, issues = collect_catalyst_overrides(
            [
                {"quote_num": "Q-1", "order number": "O-1", "order_status": "Complete"},
                {"quote_num": "Q-1", "order number": "O-1", "order_status": "Cancelled"},
            ]
        )
        self.assertEqual(overrides, {})
        self.assertEqual(issues[0].code, "conflicting_catalyst_status")

    def test_checkpoint_csv_uses_only_reviewed_status(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "review.csv"
            path.write_text(
                "quote_num,order_number,current_status,reviewed_status\n"
                "Q-1,O-1,In Progress,Complete\n"
                "Q-2,O-2,In Progress,\n",
                encoding="utf-8-sig",
            )
            overrides, issues = collect_catalyst_overrides(iter_catalyst_review(path))
        self.assertEqual(len(overrides), 1)
        self.assertEqual(next(iter(overrides.values())).order_status, "Complete")
        self.assertEqual(issues, [])


if __name__ == "__main__":
    unittest.main()
