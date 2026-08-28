import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from excel_automation_lock import excel_process_lock
from pipeline_stage_runner import (
    StageProcessResult,
    _publish_atomic_output,
    is_transient_excel_error,
    wait_for_xlsx,
)


class PipelineStageRunnerTests(unittest.TestCase):
    @staticmethod
    def _write_minimal_xlsx(path: Path) -> None:
        with zipfile.ZipFile(path, "w") as workbook:
            workbook.writestr("[Content_Types].xml", "<Types />")
            workbook.writestr("_rels/.rels", "<Relationships />")
            workbook.writestr("xl/workbook.xml", "<workbook />")

    def test_process_lock_is_reentrant_in_the_same_thread(self) -> None:
        with excel_process_lock(1):
            with excel_process_lock(1):
                pass

    def test_atomic_publish_replaces_output_only_after_valid_generation(self) -> None:
        with TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "result.xlsx"
            working_path = Path(temp_dir) / ".result.working.xlsx"

            def generate(*_args, **_kwargs):
                self._write_minimal_xlsx(working_path)
                return StageProcessResult(stdout="ok", stderr="", elapsed_seconds=1.0)

            with patch("pipeline_stage_runner._run_worker", side_effect=generate):
                result = _publish_atomic_output([], output_path, working_path, 1)

            self.assertEqual(result.stdout, "ok")
            self.assertTrue(output_path.is_file())
            self.assertFalse(working_path.exists())

    def test_atomic_publish_preserves_previous_output_when_generation_fails(self) -> None:
        with TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "result.xlsx"
            working_path = Path(temp_dir) / ".result.working.xlsx"
            output_path.write_bytes(b"previous-output")

            with patch("pipeline_stage_runner._run_worker", side_effect=RuntimeError("failed")):
                with self.assertRaisesRegex(RuntimeError, "failed"):
                    _publish_atomic_output([], output_path, working_path, 1)

            self.assertEqual(output_path.read_bytes(), b"previous-output")

    def test_excel_access_and_clipboard_timeouts_are_retryable(self) -> None:
        self.assertTrue(is_transient_excel_error("Microsoft Excel cannot access the file"))
        self.assertTrue(is_transient_excel_error("We couldn't paste this data because it took too long"))
        self.assertFalse(is_transient_excel_error("Required column QUOTE ID is missing"))

    def test_wait_for_xlsx_accepts_complete_workbook_container(self) -> None:
        with TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "result.xlsx"
            self._write_minimal_xlsx(workbook_path)

            self.assertEqual(wait_for_xlsx(workbook_path, timeout_seconds=1), workbook_path.resolve())

    def test_wait_for_xlsx_rejects_zip_without_workbook_structure(self) -> None:
        with TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "result.xlsx"
            with zipfile.ZipFile(workbook_path, "w") as workbook:
                workbook.writestr("[Content_Types].xml", "<Types />")

            with self.assertRaisesRegex(RuntimeError, "missing workbook members"):
                wait_for_xlsx(workbook_path, timeout_seconds=0)

    def test_wait_for_xlsx_rejects_incomplete_output(self) -> None:
        with TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "result.xlsx"
            workbook_path.write_text("incomplete", encoding="utf-8")

            with self.assertRaisesRegex(RuntimeError, "Workbook is not ready"):
                wait_for_xlsx(workbook_path, timeout_seconds=0)


if __name__ == "__main__":
    unittest.main()
