from __future__ import annotations

import os
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

from screenshot_upload import report_screenshot_targets, replace_inline_image_sources


class ScreenshotUploadTests(unittest.TestCase):
    def test_reports_use_stable_non_overlapping_file_numbers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.dict(os.environ, {"SCREENSHOT_ROOT": temp_dir}, clear=False):
                expected = {
                    1: ["file1.jpg", "file2.jpg", "file3.jpg", "file4.jpg"],
                    2: ["file5.jpg", "file6.jpg", "file7.jpg", "file8.jpg", "file9.jpg"],
                    3: ["file10.jpg"],
                    4: ["file11.jpg", "file12.jpg", "file13.jpg", "file14.jpg"],
                }
                generated_at = datetime(2026, 9, 1, 15, 8, 38, 123456)
                for report_number, filenames in expected.items():
                    targets = report_screenshot_targets(
                        Path(temp_dir),
                        report_number,
                        len(filenames),
                        date(2026, 9, 1),
                        generated_at,
                    )
                    self.assertEqual([path.name for path, _ in targets], filenames)
                    expected_segment = "/screenshots/2026-09-01/150838123456/"
                    self.assertTrue(all(expected_segment in url for _, url in targets))
                    self.assertTrue(
                        all(path.parent.name == "150838123456" for path, _ in targets)
                    )

    def test_replaces_content_ids_with_public_urls(self) -> None:
        message = '<img src="cid:first"><img src="cid:second">'
        result = replace_inline_image_sources(
            message,
            ["first", "second"],
            ["https://example.test/file1.jpg", "https://example.test/file2.jpg"],
        )
        self.assertNotIn("cid:", result)
        self.assertIn("https://example.test/file1.jpg", result)
        self.assertIn("https://example.test/file2.jpg", result)


if __name__ == "__main__":
    unittest.main()
