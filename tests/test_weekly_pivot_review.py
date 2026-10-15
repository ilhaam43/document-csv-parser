import tempfile
import unittest
from pathlib import Path

from weekly.pivot_review import PivotReviewDecision, load_pivot_review


class WeeklyPivotReviewTests(unittest.TestCase):
    def test_review_requires_attribution_and_preserves_decisions(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "pivot-review.csv"
            path.write_text(
                "workbook,pivot,item,visible,reviewed_by,review_reason\n"
                "master,PivotTable16,Reviewed adjustment,true,Dewi,Confirmed in pivot\n"
                "master,PivotTable20,Reviewed adjustment,true,Dewi,Confirmed in pivot\n",
                encoding="utf-8",
            )
            self.assertEqual(
                load_pivot_review(path)["master"],
                [
                    PivotReviewDecision("PivotTable16", "ADJUSTMENT", "Reviewed adjustment", True),
                    PivotReviewDecision("PivotTable20", "ADJUSTMENT", "Reviewed adjustment", True),
                ],
            )

    def test_week_position_review_is_explicit(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "pivot-review.csv"
            path.write_text(
                "workbook,pivot,field,item,visible,position,reviewed_by,review_reason\n"
                "master,PivotTable14,Week Completion,W39,,4,Dewi,Matches approved manual order\n",
                encoding="utf-8",
            )
            self.assertEqual(
                load_pivot_review(path)["master"],
                [PivotReviewDecision("PivotTable14", "Week Completion", "W39", position=4)],
            )

    def test_duplicate_or_unattributed_decisions_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "pivot-review.csv"
            header = "workbook,pivot,item,visible,reviewed_by,review_reason\n"
            row = "master,PivotTable16,Reviewed adjustment,true,Dewi,Confirmed\n"
            path.write_text(header + row + row, encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "duplicate"):
                load_pivot_review(path)
            path.write_text(header + row.replace("Dewi", ""), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "attribution"):
                load_pivot_review(path)


if __name__ == "__main__":
    unittest.main()
