import unittest

from excel_pivot_layout import section_banner_end_columns


class SectionBannerEndColumnsTests(unittest.TestCase):
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
