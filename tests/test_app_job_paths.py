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
                    data={"refresh_template": "true"},
                )

            self.assertEqual(response.status_code, 200)
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
