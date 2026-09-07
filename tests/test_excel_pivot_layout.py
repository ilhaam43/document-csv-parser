import unittest
from datetime import date

from types import SimpleNamespace

from excel_pivot_layout import (
    completion_legend_values,
    completion_mode_for_row,
    completion_section_modes,
    section_banner_end_columns,
    update_dynamic_pivot_section_titles,
)


class SectionBannerEndColumnsTests(unittest.TestCase):
    def test_completion_legend_uses_the_week_one_threshold_band(self):
        self.assertEqual(
            completion_legend_values(date(2026, 9, 1)),
            ("Green : >30%", "Yellow : >=20%", "Red : <20%"),
        )

    def test_completion_legend_keeps_week_five_for_target_after(self):
        self.assertEqual(
            completion_legend_values(date(2026, 9, 29), "target_after"),
            ("Green : >50%", "Yellow : >=40%", "Red : <40%"),
        )

    def test_target_after_mode_is_detected_from_the_nearest_section(self):
        values = (
            ("TARGET COMPLETE SEPTEMBER 2026",),
            (None,),
            ("Percentage of Completion",),
            ("TARGET AFTER SEPTEMBER 2026",),
            (None,),
            ("Percentage of Completion",),
        )
        section_modes = completion_section_modes(values, 10)

        self.assertEqual(completion_mode_for_row(12, section_modes), "weekly")
        self.assertEqual(
            completion_mode_for_row(15, section_modes),
            "target_after",
        )

    def test_updates_month_titles_without_fixed_cell_addresses(self):
        values = (
            ("TARGET COMPLETE AUGUST 2026", None),
            (None, "ALL TARGET AUGUST"),
        )
        cells = {}

        class Cell:
            def __init__(self, value=None):
                self.Value = value

        class Sheet:
            UsedRange = SimpleNamespace(Value=values, Row=7, Column=3)

            def Cells(self, row, column):
                return cells.setdefault((row, column), Cell())

        sheet = Sheet()
        updated = update_dynamic_pivot_section_titles(
            sheet,
            {
                "target complete": "TARGET COMPLETE SEPTEMBER 2026",
                "all target": "ALL TARGET SEPTEMBER",
            },
        )

        self.assertEqual(updated, 2)
        self.assertEqual(cells[(7, 3)].Value, "TARGET COMPLETE SEPTEMBER 2026")
        self.assertEqual(cells[(8, 4)].Value, "ALL TARGET SEPTEMBER")

    def test_separates_titles_that_share_the_same_row(self):
        titles = [(1, 1), (1, 19), (159, 1)]
        pivot_ranges = [
            (9, 1, 3),
            (8, 5, 10),
            (8, 19, 23),
            (177, 1, 3),
            (176, 5, 11),
        ]
        side_headers = [(9, 11), (176, 12)]

        self.assertEqual(
            section_banner_end_columns(titles, pivot_ranges, side_headers),
            {
                (1, 1): 11,
                (1, 19): 23,
                (159, 1): 12,
            },
        )

    def test_uses_pivot_edge_when_section_has_no_side_calculation(self):
        self.assertEqual(
            section_banner_end_columns(
                titles=[(92, 1), (158, 1)],
                pivot_ranges=[(110, 1, 3), (109, 5, 6), (167, 1, 3), (166, 5, 10)],
                side_headers=[(167, 11)],
            ),
            {
                (92, 1): 6,
                (158, 1): 11,
            },
        )


if __name__ == "__main__":
    unittest.main()
