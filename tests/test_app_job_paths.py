import unittest
import uuid
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from fastapi.testclient import TestClient

import app as app_module


class ReportJobPathTests(unittest.TestCase):
    def test_report_1_uses_job_local_input_and_output_paths(self) -> None:
        with TemporaryDirectory() as temp_dir:
            api_work_dir = Path(temp_dir).resolve()
            client_job_id = uuid.uuid4().hex
            client = TestClient(app_module.app)
            with (
                patch.object(app_module, "API_WORK_DIR", api_work_dir),
                patch.object(app_module.REPORT_1_EXECUTOR, "submit") as submit,
            ):
                response = client.post(
                    "/convert/upload/jobs",
                    files={
                        "raw_data": (
                            "DataOrderSD-20260828-060743.csv",
                            b"quo,status\nQ1,new\n",
                            "text/csv",
                        ),
                        "yesterday_cleaned_data": (
                            "Daily Tracking 27 August 2026.xlsx",
                            b"uploaded-reference",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        ),
                    },
                    data={"job_id": client_job_id, "refresh_template": "true"},
                )

            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.json()["job_id"], client_job_id)
            submitted = submit.call_args.args
            raw_path = submitted[2]
            reference_path = submitted[3]
            output_path = submitted[4]

            self.assertEqual(raw_path.parent.name, "raw-data")
            self.assertEqual(reference_path.parent.name, "daily-reference")
            self.assertEqual(raw_path.parent.parent, reference_path.parent.parent)
            self.assertEqual(raw_path.parent.parent.name, "input")
            self.assertEqual(output_path.parent, raw_path.parent.parent.parent)
            self.assertNotEqual(output_path.parent, (app_module.APP_ROOT / app_module.CONVERSION_OUTPUT_DIR).resolve())
            self.assertEqual(output_path.name, "Daily Tracking 28 August 2026.xlsx")
            with app_module.REPORT_1_JOBS_LOCK:
                app_module.REPORT_1_JOBS.pop(client_job_id, None)

    def test_report_4_honors_client_job_id_for_page_resume(self) -> None:
        with TemporaryDirectory() as temp_dir:
            api_work_dir = Path(temp_dir).resolve()
            client_job_id = uuid.uuid4().hex
            client = TestClient(app_module.app)
            with (
                patch.object(app_module, "API_WORK_DIR", api_work_dir),
                patch.object(app_module.REPORT_4_EXECUTOR, "submit") as submit,
            ):
                response = client.post(
                    "/convert/report-4/upload/jobs",
                    files={
                        "raw_ide_workbook": (
                            "IDE DASHBOARD 7 September 2026.xlsx",
                            b"raw-ide",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        ),
                        "previous_ide_workbook": (
                            "Daily Tracking IDE 4 September 2026.xlsx",
                            b"previous-ide",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        ),
                        "collabs_csv": (
                            "bq-results.csv",
                            b"quote_id,otc,mrc\nQ1,0,0\n",
                            "text/csv",
                        ),
                    },
                    data={"job_id": client_job_id},
                )

            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.json()["job_id"], client_job_id)
            self.assertEqual(submit.call_args.args[1], client_job_id)
            self.assertEqual(submit.call_args.args[5].name, "Daily Tracking IDE 7 September 2026.xlsx")
            with app_module.REPORT_4_JOBS_LOCK:
                app_module.REPORT_4_JOBS.pop(client_job_id, None)

    def test_report_2_and_3_honor_client_job_ids_for_page_resume(self) -> None:
        cases = (
            (
                "/convert/report-2/upload/jobs",
                app_module.REPORT_2_EXECUTOR,
                app_module.REPORT_2_JOBS,
                app_module.REPORT_2_JOBS_LOCK,
                {
                    "tracking_workbook": (
                        "Daily Tracking 7 September 2026.xlsx",
                        b"daily",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                    "log_update_status": ("LogUpdateStatusOrderSD.csv", b"quo,status\nQ1,new\n", "text/csv"),
                    "previous_ongoing_workbook": (
                        "Daily Tracking 4 September 2026 On Going.xlsx",
                        b"ongoing",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                },
            ),
            (
                "/convert/report-3/upload/jobs",
                app_module.REPORT_3_EXECUTOR,
                app_module.REPORT_3_JOBS,
                app_module.REPORT_3_JOBS_LOCK,
                {
                    "tracking_workbook": (
                        "Daily Tracking 7 September 2026.xlsx",
                        b"daily",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                    "previous_iphone_workbook": (
                        "Daily Tracking Iphone 4 September 2026.xlsx",
                        b"iphone",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                },
            ),
        )

        for endpoint, executor, jobs, jobs_lock, files in cases:
            with self.subTest(endpoint=endpoint), TemporaryDirectory() as temp_dir:
                client_job_id = uuid.uuid4().hex
                client = TestClient(app_module.app)
                with (
                    patch.object(app_module, "API_WORK_DIR", Path(temp_dir).resolve()),
                    patch.object(executor, "submit") as submit,
                ):
                    response = client.post(endpoint, files=files, data={"job_id": client_job_id})

                self.assertEqual(response.status_code, 200)
                self.assertEqual(response.json()["job_id"], client_job_id)
                self.assertEqual(submit.call_args.args[1], client_job_id)
                with jobs_lock:
                    jobs.pop(client_job_id, None)

    def test_pipeline_uses_isolated_input_roles_and_job_local_outputs(self) -> None:
        with TemporaryDirectory() as temp_dir:
            api_work_dir = Path(temp_dir).resolve()
            client_job_id = uuid.uuid4().hex
            client = TestClient(app_module.app)
            with (
                patch.object(app_module, "API_WORK_DIR", api_work_dir),
                patch.object(app_module.PIPELINE_EXECUTOR, "submit") as submit,
            ):
                response = client.post(
                    "/convert/pipeline/upload/jobs",
                    files={
                        "raw_data": ("DataOrderSD-20260828-060743.csv", b"quo,status\nQ1,new\n", "text/csv"),
                        "yesterday_cleaned_data": ("Daily Tracking 27 August 2026.xlsx", b"daily", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                        "log_update_status": ("LogUpdateStatusOrderSD-20260828.csv", b"quo,status\nQ1,new\n", "text/csv"),
                        "previous_ongoing_workbook": ("Daily Tracking 27 August 2026.xlsx", b"ongoing", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                        "previous_iphone_workbook": ("Daily Tracking 27 August 2026.xlsx", b"iphone", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    },
                    data={
                        "job_id": client_job_id,
                        "refresh_template": "true",
                        "ongoing_with_pivot": "false",
                    },
                )

            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.json()["job_id"], client_job_id)
            submitted = submit.call_args.args
            input_paths = submitted[2:7]
            output_dir = submitted[7]
            zip_path = submitted[8]

            self.assertEqual(
                {path.parent.name for path in input_paths},
                {"raw-data", "daily-reference", "log-update", "ongoing-reference", "iphone-reference"},
            )
            input_root = input_paths[0].parent.parent
            self.assertTrue(all(path.parent.parent == input_root for path in input_paths))
            self.assertEqual(input_root.name, "input")
            self.assertEqual(output_dir.parent, input_root.parent)
            self.assertEqual(output_dir.name, "pipeline-output")
            self.assertEqual(zip_path.parent, input_root.parent)
            self.assertEqual(zip_path.parent.name, client_job_id)
            with app_module.PIPELINE_JOBS_LOCK:
                app_module.PIPELINE_JOBS.pop(client_job_id, None)


if __name__ == "__main__":
    unittest.main()
