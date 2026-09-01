import unittest
from datetime import date

from csv_to_excel_on_going import (
    DEFAULT_ON_PROGRESS_PIVOT_STYLE,
    expected_on_progress_sheet_name,
    resolve_on_progress_pivot_style,
)


class FakePivot:
    def __init__(self, style_name):
        self.TableStyle2 = style_name


class FakeProgressSheet:
    def __init__(self, pivots):
        self._pivots = pivots

    def PivotTables(self, name):
        return self._pivots[name]


class ResolveOnProgressPivotStyleTests(unittest.TestCase):
    def test_uses_first_available_template_pivot_style(self):
        sheet = FakeProgressSheet(
            {
                "missing-style": FakePivot(""),
                "template-pivot": FakePivot("PivotStyleLight16"),
            }
        )

        self.assertEqual(
            resolve_on_progress_pivot_style(sheet, ["missing-style", "template-pivot"]),
            "PivotStyleLight16",
        )

    def test_uses_light_style_fallback_when_template_style_is_unavailable(self):
        sheet = FakeProgressSheet({})

        self.assertEqual(
            resolve_on_progress_pivot_style(sheet, ["missing-pivot"]),
            DEFAULT_ON_PROGRESS_PIVOT_STYLE,
        )


class OnProgressSheetNameTests(unittest.TestCase):
    def test_keeps_full_month_name_when_it_fits_excel_limit(self):
        self.assertEqual(
            expected_on_progress_sheet_name(date(2026, 8, 31)),
            "ALL ORDER ON PROGRESS 31 August",
        )

    def test_abbreviates_month_when_full_name_exceeds_excel_limit(self):
        sheet_name = expected_on_progress_sheet_name(date(2026, 9, 1))

        self.assertEqual(sheet_name, "ALL ORDER ON PROGRESS 1 Sep")
        self.assertLessEqual(len(sheet_name), 31)


if __name__ == "__main__":
    unittest.main()
