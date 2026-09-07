import unittest

from completion_threshold_rules import TARGET_AFTER_MODE, WEEKLY_MODE
from generate_iphone_tracking import IPHONE_COMPLETION_LAYOUT_SPECS


class IphoneThresholdLayoutTests(unittest.TestCase):
    def test_all_target_block_uses_target_after_thresholds(self) -> None:
        self.assertEqual(
            tuple(mode for _, _, mode in IPHONE_COMPLETION_LAYOUT_SPECS),
            (WEEKLY_MODE, WEEKLY_MODE, TARGET_AFTER_MODE),
        )


if __name__ == "__main__":
    unittest.main()
