import tempfile
import unittest
from datetime import date
from pathlib import Path

from weekly.slide import completion_month, load_slide_review, reporting_week_start, week_label
from weekly.workbook import _slide_summary_rows


class WeeklySlideTests(unittest.TestCase):
    def test_week_labels_follow_thursday_excel_weeks(self) -> None:
        self.assertEqual(week_label(date(2026, 9, 1), 35), "W1-Sep")
        self.assertEqual(week_label(date(2026, 9, 1), 37), "W3-Sep")
        self.assertEqual(week_label(date(2026, 9, 1), 38), "W4-Sep")
        self.assertEqual(reporting_week_start(2026, 39), date(2026, 9, 24))
        self.assertEqual(reporting_week_start(2026, 40), date(2026, 10, 1))
        self.assertEqual(reporting_week_start(2026, 35), date(2026, 8, 27))
        self.assertEqual(reporting_week_start(2026, 36), date(2026, 9, 3))
        self.assertEqual(week_label(date(2026, 10, 1), 40), "W1-Oct")
        self.assertEqual(week_label(date(2027, 1, 1), 1), "W1-Jan")
        self.assertEqual(completion_month("Mar26"), date(2026, 3, 1))
        self.assertEqual(completion_month("Sep-26"), date(2026, 9, 1))
        self.assertIsNone(completion_month("W40"))

    def test_w46_chart_uses_weeks_under_november_even_if_w44_starts_in_october(self) -> None:
        labels = [(18, "Oct-26"), (19, "W44"), (20, "Nov-26"), (21, "W44"),
                  (22, "W45"), (23, "W46"), (24, "W47"), (25, "Dec-26"),
                  (26, "W46"), (27, "Grand Total")]
        self.assertEqual(reporting_week_start(2026, 44), date(2026, 10, 29))
        self.assertEqual(_slide_summary_rows(
            labels, report_month=date(2026, 11, 1), report_week=46,
            old_labels={date(2026, 10, 1): "Oct26"}, require_current_week=True,
        ), [("Oct26", 18), ("W1-Nov", 21), ("W2-Nov", 22), ("W3-Nov", 23)])

    def test_missing_weeks_do_not_shift_labels_or_use_another_months_values(self) -> None:
        labels = [(19, "Nov-26"), (20, "W45"), (21, "Dec-26"), (22, "W46")]
        params = dict(report_month=date(2026, 11, 1), report_week=46, old_labels={})
        with self.assertRaisesRegex(RuntimeError, "no current report week"):
            _slide_summary_rows(labels, **params, require_current_week=True)
        self.assertEqual(_slide_summary_rows(labels, **params, require_current_week=False),
                         [("W1-Nov", None), ("W2-Nov", 20), ("W3-Nov", None)])

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
