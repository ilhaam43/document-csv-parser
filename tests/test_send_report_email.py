from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from send_report_1_email import multipart_form_attachment


class SendReportEmailTests(unittest.TestCase):
    def test_multipart_message_attaches_workbook_not_screenshot(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook = Path(temp_dir) / "Daily Report.xlsx"
            workbook.write_bytes(b"test-workbook-content")
            fields = {
                "to": "recipient@example.test",
                "subject": "Testing",
                "message": '<img src="https://example.test/screenshots/2026-09-01/file1.jpg">',
            }

            body, content_type = multipart_form_attachment(fields, workbook)

        self.assertIn("multipart/form-data; boundary=", content_type)
        self.assertIn(b'name="attachment[]"; filename="Daily Report.xlsx"', body)
        self.assertIn(b"test-workbook-content", body)
        self.assertIn(b"https://example.test/screenshots/2026-09-01/file1.jpg", body)
        self.assertNotIn(b"attachment_content_id", body)
        self.assertNotIn(b"Content-ID", body)


if __name__ == "__main__":
    unittest.main()
