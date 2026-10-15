import tempfile
import unittest
from datetime import date
from pathlib import Path

from weekly.slide import completion_month, load_slide_review, reporting_week_start, week_label


class WeeklySlideTests(unittest.TestCase):
    def test_week_labels_follow_thursday_excel_weeks(self) -> None:
        self.assertEqual(week_label(2026, 37), "W2-Sep")
        self.assertEqual(week_label(2026, 38), "W3-Sep")
        self.assertEqual(reporting_week_start(2026, 39), date(2026, 9, 24))
        self.assertEqual(reporting_week_start(2026, 40), date(2026, 10, 1))
        self.assertEqual(reporting_week_start(2026, 35), date(2026, 8, 27))
        self.assertEqual(reporting_week_start(2026, 36), date(2026, 9, 3))
        self.assertEqual(week_label(2026, 40), "W1-Oct")
        self.assertEqual(completion_month("Mar26"), date(2026, 3, 1))
        self.assertEqual(completion_month("Sep-26"), date(2026, 9, 1))
        self.assertIsNone(completion_month("W40"))

    def test_review_requires_attribution_and_numeric_values(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "slide.csv"
            header = "workbook,cell,value,reviewed_by,review_reason\n"
            path.write_text(header + "master,U21,16,Dewi,Manual total confirmed\n")
            self.assertEqual(load_slide_review(path)["master"], {"U21": 16.0})
            path.write_text(header + "master,U21,16,,\n")
            with self.assertRaisesRegex(ValueError, "reviewer and reason"):
                load_slide_review(path)
            path.write_text(header + "master,U21,not-a-number,Dewi,Confirmed\n")
            with self.assertRaisesRegex(ValueError, "numeric"):
                load_slide_review(path)
            path.write_text(header + "master,A1,16,Dewi,Confirmed\n")
            with self.assertRaisesRegex(ValueError, "S10:V100"):
                load_slide_review(path)


if __name__ == "__main__":
    unittest.main()
