import unittest
import warnings
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

from excel_automation_lock import excel_process_lock
from pipeline_stage_runner import (
    StageProcessResult,
    _publish_atomic_output,
    is_transient_excel_error,
    main,
    optimize_xlsx_archive,
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


class WorkbookCompressionTests(unittest.TestCase):
    def write_package(self, path, *, compression=zipfile.ZIP_STORED):
        with zipfile.ZipFile(path, 'w', compression=compression, compresslevel=9) as book:
            book.comment = b'original package comment'
            for name, value in (
                ('[Content_Types].xml', b'<Types/>'), ('_rels/.rels', b'<Relationships/>'),
                ('xl/workbook.xml', b'<workbook/>'),
                ('xl/pivotCache/pivotCacheRecords1.xml', b'<r><n v="123"/></r>' * 2000),
                ('xl/pivotCache/pivotCacheRecords2.xml', b'<r><n v="123"/></r>' * 2000),
                ('xl/media/image1.png', bytes(range(256))),
            ):
                info = zipfile.ZipInfo(name, date_time=(2026, 11, 2, 6, 0, 0))
                info.comment = b'part comment'
                info.external_attr = 0o600 << 16
                book.writestr(info, value, compress_type=compression, compresslevel=9)

    def test_compression_retains_every_part_cache_and_metadata(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'report.xlsx'
            self.write_package(path)
            with zipfile.ZipFile(path) as before:
                parts = {i.filename: (before.read(i), i.date_time, i.comment, i.external_attr)
                         for i in before.infolist()}
            old_size, new_size = optimize_xlsx_archive(path)
            self.assertLess(new_size, old_size)
            self.assertEqual(new_size, path.stat().st_size)
            with zipfile.ZipFile(path) as after:
                self.assertEqual(after.comment, b'original package comment')
                self.assertEqual(parts, {i.filename: (after.read(i), i.date_time, i.comment, i.external_attr)
                                        for i in after.infolist()})
                self.assertIsNone(after.testzip())
            self.assertEqual(list(Path(directory).iterdir()), [path])

    def test_already_optimized_package_is_not_replaced_when_no_smaller(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'report.xlsx'
            self.write_package(path, compression=zipfile.ZIP_DEFLATED)
            original = path.read_bytes()
            self.assertEqual(optimize_xlsx_archive(path), (len(original), len(original)))
            self.assertEqual(path.read_bytes(), original)

    def test_failure_retains_original_and_cleans_temporary_package(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'report.xlsx'
            self.write_package(path)
            original = path.read_bytes()
            with patch('pipeline_stage_runner.zipfile.ZipFile.writestr', side_effect=OSError('disk full')):
                with self.assertRaisesRegex(OSError, 'disk full'):
                    optimize_xlsx_archive(path)
            self.assertEqual(path.read_bytes(), original)
            self.assertEqual(list(Path(directory).iterdir()), [path])

    def test_long_report_names_do_not_expand_the_temporary_filename(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / ('report-' + 'x' * 140 + '.xlsx')
            self.write_package(path)
            original_write = zipfile.ZipFile.writestr
            def checked_write(archive, *args, **kwargs):
                self.assertLess(len(Path(archive.filename).name), 50)
                return original_write(archive, *args, **kwargs)
            with patch('pipeline_stage_runner.zipfile.ZipFile.writestr', new=checked_write):
                before, after = optimize_xlsx_archive(path)
            self.assertLess(after, before)
            self.assertEqual(list(Path(directory).iterdir()), [path])

    def test_incomplete_package_is_rejected_without_replacement(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'report.xlsx'
            with zipfile.ZipFile(path, 'w') as book:
                book.writestr('unrelated.xml', b'data')
            original = path.read_bytes()
            with self.assertRaisesRegex(ValueError, 'missing required'):
                optimize_xlsx_archive(path)
            self.assertEqual(path.read_bytes(), original)

    def test_duplicate_parts_are_rejected_without_losing_the_original(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'report.xlsx'
            self.write_package(path)
            with warnings.catch_warnings():
                warnings.simplefilter('ignore', UserWarning)
                with zipfile.ZipFile(path, 'a') as book:
                    book.writestr('xl/workbook.xml', b'<duplicate/>')
            original = path.read_bytes()
            with self.assertRaisesRegex(ValueError, 'duplicate package parts'):
                optimize_xlsx_archive(path)
            self.assertEqual(path.read_bytes(), original)

    def test_failed_integrity_check_does_not_publish_the_compressed_file(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'report.xlsx'
            self.write_package(path)
            original = path.read_bytes()
            with patch('pipeline_stage_runner.zipfile.ZipFile.testzip', return_value='xl/workbook.xml'):
                with self.assertRaisesRegex(RuntimeError, 'changed during compression'):
                    optimize_xlsx_archive(path)
            self.assertEqual(path.read_bytes(), original)
            self.assertEqual(list(Path(directory).iterdir()), [path])

    def test_pipeline_workers_compress_after_generation_and_leave_ide_unchanged(self):
        with TemporaryDirectory() as directory:
            for stage in ('daily', 'ongoing', 'iphone', 'ide'):
                with self.subTest(stage=stage), \
                     patch('pipeline_stage_runner.parse_args', return_value=SimpleNamespace(
                         stage=stage, output=Path(directory) / 'result.xlsx')), \
                     patch(f'pipeline_stage_runner._run_{stage}') as generate, \
                     patch('pipeline_stage_runner.optimize_xlsx_archive', return_value=(100, 90)) as compress:
                    compress.side_effect = lambda _: (generate.assert_called_once(), (100, 90))[1]
                    self.assertEqual(main(['test']), 0)
                    if stage == 'ide':
                        compress.assert_not_called()
                    else:
                        compress.assert_called_once()


if __name__ == "__main__":
    unittest.main()
