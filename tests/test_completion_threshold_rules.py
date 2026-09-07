import unittest
from datetime import date

from completion_threshold_rules import (
    ONGOING_PROFILE,
    STANDARD_PROFILE,
    TARGET_AFTER_MODE,
    WEEKLY_MODE,
    completion_thresholds_for,
    report_week,
)


class CompletionThresholdRulesTests(unittest.TestCase):
    def test_week_boundaries_include_a_distinct_week_five(self) -> None:
        expected_weeks = {
            1: 1,
            6: 1,
            7: 2,
            13: 2,
            14: 3,
            20: 3,
            21: 4,
            28: 4,
            29: 5,
            30: 5,
            31: 5,
        }
        for day, expected_week in expected_weeks.items():
            with self.subTest(day=day):
                self.assertEqual(report_week(date(2026, 1, day)), expected_week)

    def test_standard_profile_matches_reports_one_three_and_four(self) -> None:
        expected = {
            WEEKLY_MODE: {
                1: (30, 20),
                2: (40, 30),
                3: (50, 40),
                4: (60, 50),
                5: (60, 50),
            },
            TARGET_AFTER_MODE: {
                1: (10, 5),
                2: (20, 10),
                3: (30, 20),
                4: (40, 30),
                5: (50, 40),
            },
        }
        self.assert_profile(STANDARD_PROFILE, expected)

    def test_ongoing_profile_matches_report_two(self) -> None:
        expected = {
            WEEKLY_MODE: {
                1: (30, 20),
                2: (40, 30),
                3: (50, 40),
                4: (99, 98),
                5: (99, 98),
            },
            TARGET_AFTER_MODE: {
                1: (10, 5),
                2: (20, 10),
                3: (30, 20),
                4: (99, 98),
                5: (99, 98),
            },
        }
        self.assert_profile(ONGOING_PROFILE, expected)

    def assert_profile(
        self,
        profile: str,
        expected: dict[str, dict[int, tuple[int, int]]],
    ) -> None:
        representative_days = {1: 1, 2: 7, 3: 14, 4: 21, 5: 29}
        for mode, week_values in expected.items():
            for week, (green, lower) in week_values.items():
                with self.subTest(profile=profile, mode=mode, week=week):
                    thresholds = completion_thresholds_for(
                        date(2026, 1, representative_days[week]),
                        profile=profile,
                        mode=mode,
                    )
                    self.assertEqual(
                        (
                            thresholds.week,
                            thresholds.green_percent,
                            thresholds.yellow_percent,
                            thresholds.red_percent,
                        ),
                        (week, green, lower, lower),
                    )


if __name__ == "__main__":
    unittest.main()
