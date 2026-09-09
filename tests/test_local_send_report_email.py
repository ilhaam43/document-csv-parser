from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from local_send_report_email import encode_multipart


class LocalSendReportEmailTests(unittest.TestCase):
    def test_prepare_upload_uses_workbook_field(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook = Path(temp_dir) / "report.xlsx"
            workbook.write_bytes(b"workbook-data")
            body, _ = encode_multipart({"report_number": "3"}, workbook, "workbook")
        self.assertIn(b'name="workbook"; filename="report.xlsx"', body)
        self.assertNotIn(b'name="attachment[]"', body)

    def test_smtp_upload_uses_attachment_array_field(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook = Path(temp_dir) / "report.xlsx"
            workbook.write_bytes(b"workbook-data")
            body, _ = encode_multipart({"message": "<p>Report</p>"}, workbook, "attachment[]")
        self.assertIn(b'name="attachment[]"; filename="report.xlsx"', body)


if __name__ == "__main__":
    unittest.main()
