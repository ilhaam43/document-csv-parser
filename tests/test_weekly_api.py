import json
import csv
import unittest
import uuid
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from fastapi.testclient import TestClient

from app import app
from weekly import api as weekly_api


class WeeklyApiTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        patcher = patch.object(weekly_api, "JOB_ROOT", self.root)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.client = TestClient(app)

    def _start(self, job_id: str | None = None, *, with_ppt: bool = False,
               report_month: str | None = None):
        files = {
            "bq_csv": ("source.csv", b"quote_num\nQ1\n", "text/csv"),
            "previous_master": ("master.xlsx", b"master", "application/octet-stream"),
            "previous_mobile": ("mobile.xlsx", b"mobile", "application/octet-stream"),
            "holiday_workbook": ("holidays.xlsx", b"holidays", "application/octet-stream"),
        }
        if with_ppt:
            files["previous_ppt"] = ("mobile.pptx", b"pptx", "application/octet-stream")
        return self.client.post(
            "/api/weekly/jobs",
            files=files,
            data={"report_date": "2026-09-24", "report_week": "38",
                  "report_month": report_month or "2026-09", "job_id": job_id or uuid.uuid4().hex},
        )

    def test_prepare_review_finalize_and_download(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit") as submit:
            response = self._start()
            self.assertEqual(response.status_code, 202)
            job_id = response.json()["job_id"]
            job_dir = self.root / job_id
            self.assertEqual(submit.call_args.args, (weekly_api._run_prepare, job_id))
            self.assertEqual(weekly_api._read_state(job_id)["report_month"], "2026-09")
            self.assertEqual((job_dir / "input" / "bq.csv").read_bytes(), b"quote_num\nQ1\n")
            self.assertEqual(self.client.get(f"/api/weekly/jobs/{job_id}/download").status_code, 409)

            def prepare(**kwargs):
                output_dir = kwargs["output_dir"]
                output_dir.mkdir(parents=True)
                (output_dir / "catalyst-review.csv").write_text(
                    "quote_num,order_number,reviewed_status\nQ1,O1,\n", encoding="utf-8"
                )
                (output_dir / "draft-master-data.csv").write_text(
                    "quote_num,order_status\nQ1,#N/A\n", encoding="utf-8"
                )
                (output_dir / "slide-review.csv").write_text(
                    "workbook,cell,value,reviewed_by,review_reason\n", encoding="utf-8"
                )
                return {"selected_source_rows": 1, "generated_master_rows": 1}

            with patch.object(weekly_api, "prepare_weekly_inputs", side_effect=prepare):
                weekly_api._run_prepare(job_id)
            status = self.client.get(f"/api/weekly/jobs/{job_id}").json()
            self.assertEqual(status["status"], "waiting_review")
            self.assertIn("Q1", self.client.get(status["review_url"]).text)
            self.assertIn("#N/A", self.client.get(status["draft_url"]).text)
            self.assertIn("workbook,cell", self.client.get(status["slide_review_url"]).text)

            review_response = self.client.post(
                f"/api/weekly/jobs/{job_id}/finalize",
                files={
                    "catalyst_review": ("review.csv", b"quote_num,order_number,reviewed_status\nQ1,O1,COMPLETE\n", "text/csv"),
                    "slide_review": ("slide.csv", b"workbook,cell,value,reviewed_by,review_reason\n", "text/csv"),
                    "pivot_review": (
                        "pivot.csv",
                        b"workbook,pivot,item,visible,reviewed_by,review_reason\n",
                        "text/csv",
                    ),
                },
            )
            self.assertEqual(review_response.status_code, 202)
            self.assertEqual(submit.call_args.args, (weekly_api._run_finalize, job_id, "catalyst-review-upload.csv", None, "slide-review-upload.csv", "pivot-review-upload.csv", None, None))

            def finalize(**kwargs):
                output_dir = kwargs["output_dir"]
                output_dir.mkdir(parents=True)
                (output_dir / "master-final-data.csv").write_text("master", encoding="utf-8")
                (output_dir / "mobile-data.csv").write_text("mobile", encoding="utf-8")
                (output_dir / "audit.json").write_text("{}", encoding="utf-8")
                (output_dir / "exceptions.csv").write_text("code\n", encoding="utf-8")
                (output_dir / "review-decisions.csv").write_text("master_row\n", encoding="utf-8")
                (output_dir / "bq-source-decisions.csv").write_text("master_row\n", encoding="utf-8")
                return {"invalid_master_services": [], "issue_counts": {}, "pending_catalyst_rows": 2}

            def workbook(**kwargs):
                kwargs["output_workbook"].write_bytes(b"xlsx")

            with (
                patch.object(weekly_api, "finalize_weekly_data", side_effect=finalize),
                patch.object(weekly_api, "write_weekly_workbook", side_effect=workbook) as writer,
            ):
                weekly_api._run_finalize(
                    job_id, "catalyst-review-upload.csv", None,
                    "slide-review-upload.csv", "pivot-review-upload.csv",
                )
            self.assertEqual(writer.call_count, 2)
            self.assertEqual(writer.call_args.kwargs["report_week"], 38)
            self.assertEqual(
                writer.call_args.kwargs["pivot_review"],
                job_dir / "input" / "pivot-review-upload.csv",
            )
            status = self.client.get(f"/api/weekly/jobs/{job_id}").json()
            self.assertEqual(status["status"], "succeeded")
            self.assertIn("2 Catalyst statuses remain unresolved", status["warnings"][0])
            self.assertIn("SLIDE GH MEETING", status["warnings"][1])
            self.assertEqual(self.client.get(status["download_url"]).status_code, 200)
            with zipfile.ZipFile(job_dir / "weekly-mobile-output.zip") as archive:
                self.assertEqual(len(archive.namelist()), 8)
                self.assertIn("exceptions.csv", archive.namelist())
                self.assertIn("review-decisions.csv", archive.namelist())
                self.assertIn("bq-source-decisions.csv", archive.namelist())
                self.assertIn("pivot-review.csv", archive.namelist())
                self.assertIn("slide-review.csv", archive.namelist())
            self.assertEqual(json.loads((job_dir / "job.json").read_text())["status"], "succeeded")

    def test_review_requires_prepared_job_and_rejects_duplicate_submission(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            response = self._start()
            job_id = response.json()["job_id"]
            self.assertEqual(self._start(job_id).status_code, 409)
            review = {
                "catalyst_review": ("review.csv", b"review", "text/csv"),
                "slide_review": ("slide.csv", b"workbook,cell,value,reviewed_by,review_reason\n", "text/csv"),
            }
            self.assertEqual(self.client.post(f"/api/weekly/jobs/{job_id}/finalize", files=review).status_code, 409)
            self.assertEqual(self.client.get("/api/weekly/jobs/not-a-uuid").status_code, 404)

    def test_inline_catalyst_review_uses_prepared_row_identity(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit") as submit:
            job_id = self._start().json()["job_id"]
            path = self.root / job_id / "preparation" / "catalyst-review.csv"
            path.parent.mkdir(parents=True)
            path.write_text(
                "master_row,quote_num,order_number,ca_name,service_group,fab_upload,current_status,reviewed_status\n"
                "42,Q-1,O-1,Customer One,Mobile,2026-09-20,#N/A,\n"
                "43,Q-2,O-2,Customer Two,Bulk Service,2026-09-21,#N/A,\n",
                encoding="utf-8-sig",
            )
            weekly_api._update_state(job_id, status="waiting_review")
            status = self.client.get(f"/api/weekly/jobs/{job_id}").json()
            rows = self.client.get(status["review_data_url"]).json()["rows"]
            self.assertEqual([row["quote_num"] for row in rows], ["Q-1", "Q-2"])

            endpoint = f"/api/weekly/jobs/{job_id}/finalize"
            for decisions in (
                {"42": "COMPLETE"},
                {"42": "COMPLETE", "43": "CANCEL", "44": "COMPLETE"},
                {"42": "", "43": "CANCEL"},
                {"42": "INVALID", "43": "CANCEL"},
            ):
                response = self.client.post(endpoint, data={"reviewed_statuses": json.dumps(decisions)})
                self.assertEqual(response.status_code, 400)
                self.assertEqual(weekly_api._read_state(job_id)["status"], "waiting_review")
                submit.assert_called_once()

            response = self.client.post(
                endpoint,
                data={"reviewed_statuses": json.dumps({"42": "Complete", "43": "Cancelled"})},
            )
            self.assertEqual(response.status_code, 202)
            output = self.root / job_id / "input" / "catalyst-review-inline.csv"
            with output.open(encoding="utf-8-sig", newline="") as handle:
                saved = list(csv.DictReader(handle))
            self.assertEqual([row["reviewed_status"] for row in saved], ["COMPLETE", "CANCEL"])
            self.assertEqual([row["quote_num"] for row in saved], ["Q-1", "Q-2"])
            self.assertEqual(
                submit.call_args.args,
                (weekly_api._run_finalize, job_id, "catalyst-review-inline.csv", None, None, None, None, None),
            )
            self.assertEqual(self.client.post(endpoint, data={"reviewed_statuses": "{}"}).status_code, 409)

    def test_inline_review_requires_prepared_rows_and_exactly_one_input_mode(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            job_id = self._start().json()["job_id"]
            weekly_api._update_state(job_id, status="waiting_review")
            endpoint = f"/api/weekly/jobs/{job_id}/finalize"
            self.assertEqual(self.client.get(f"/api/weekly/jobs/{job_id}/review-data").status_code, 404)
            self.assertEqual(self.client.post(endpoint).status_code, 400)
            self.assertEqual(self.client.post(endpoint, data={"reviewed_statuses": "{}"}).status_code, 404)
            self.assertEqual(
                self.client.post(
                    endpoint,
                    data={"reviewed_statuses": "{}"},
                    files={"catalyst_review": ("review.csv", b"review", "text/csv")},
                ).status_code,
                400,
            )

    def test_inline_source_review_uses_prepared_conflict_identity(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit") as submit:
            job_id = self._start().json()["job_id"]
            preparation = self.root / job_id / "preparation"
            preparation.mkdir(parents=True)
            (preparation / "catalyst-review.csv").write_text(
                "master_row,quote_num,order_number,reviewed_status\n", encoding="utf-8"
            )
            (preparation / "source-review.csv").write_text(
                "master_row,quote_num,order_number,ca_name,service_group,order_status,order_completion_date,quote_status,quote_completion_date,reviewed_source\n"
                "5,Q-1,O-2,Customer,Mobile,Open,,Complete,2026-09-30,\n"
                "6,Q-2,O-4,Customer,Mobile,Open,,Complete,2026-09-30,\n",
                encoding="utf-8-sig",
            )
            weekly_api._update_state(job_id, status="waiting_review", preparation_audit={"source_conflict_rows": 2})
            status = self.client.get(f"/api/weekly/jobs/{job_id}").json()
            source_rows = self.client.get(status["source_review_data_url"]).json()["rows"]
            self.assertEqual(source_rows[0]["quote_status"], "Complete")
            endpoint = f"/api/weekly/jobs/{job_id}/finalize"
            for choices in ({"5": "invalid"}, {"5": "quote", "7": "order"}):
                response = self.client.post(
                    endpoint,
                    data={"reviewed_statuses": "{}", "reviewed_sources": json.dumps(choices)},
                )
                self.assertEqual(response.status_code, 400)
            response = self.client.post(
                endpoint,
                data={"reviewed_statuses": "{}", "reviewed_sources": '{"5":"quote"}'},
            )
            self.assertEqual(response.status_code, 202)
            self.assertEqual(submit.call_args.args[-1], "source-review-inline.csv")
            with (self.root / job_id / "input" / "source-review-inline.csv").open(
                encoding="utf-8-sig", newline=""
            ) as handle:
                saved = list(csv.DictReader(handle))
            self.assertEqual(saved[0]["reviewed_source"], "quote")
            self.assertEqual(saved[1]["reviewed_source"], "")
            weekly_api._update_state(job_id, status="waiting_review")
            without_source = self.client.post(endpoint, data={"reviewed_statuses": "{}"})
            self.assertEqual(without_source.status_code, 202)
            self.assertIsNone(submit.call_args.args[-1])

    def test_legacy_review_list_is_bounded_for_browser(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            job_id = self._start().json()["job_id"]
            path = self.root / job_id / "preparation" / "catalyst-review.csv"
            path.parent.mkdir(parents=True)
            with path.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow(("master_row", "quote_num", "order_number", "reviewed_status"))
                writer.writerows((row, f"Q{row}", f"O{row}", "") for row in range(2, 5003))
            weekly_api._update_state(job_id, status="waiting_review")
            self.assertEqual(self.client.get(f"/api/weekly/jobs/{job_id}/review-data").status_code, 409)
            self.assertEqual(self.client.get(f"/api/weekly/jobs/{job_id}/review").status_code, 200)

    def test_legacy_job_can_refresh_review_without_reuploading_or_overwriting_old_files(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit") as submit:
            job_id = self._start().json()["job_id"]
            job_dir = self.root / job_id
            old_review = job_dir / "preparation" / "catalyst-review.csv"
            old_review.parent.mkdir(parents=True)
            old_review.write_text("old review data", encoding="utf-8")
            weekly_api._update_state(job_id, status="waiting_review")
            status = self.client.get(f"/api/weekly/jobs/{job_id}").json()
            self.assertNotIn("draft_url", status)
            self.assertIn("refresh-review", status["review_refresh_url"])
            response = self.client.post(status["review_refresh_url"])
            self.assertEqual(response.status_code, 202)
            self.assertEqual(submit.call_args.args, (weekly_api._run_refresh_review, job_id))
            self.assertEqual(self.client.post(status["review_refresh_url"]).status_code, 409)

            def prepare(**kwargs):
                output = kwargs["output_dir"]
                self.assertEqual(output, job_dir / "preparation-refreshed")
                output.mkdir()
                (output / "catalyst-review.csv").write_text(
                    "master_row,quote_num,order_number,reviewed_status\n42,Q-1,O-1,\n",
                    encoding="utf-8",
                )
                (output / "draft-master-data.csv").write_text("quote_num\nQ-1\n", encoding="utf-8")
                (output / "exceptions.csv").write_text("code\n", encoding="utf-8")
                return {"selected_source_rows": 1, "generated_master_rows": 1,
                        "pending_catalyst_rows": 1}

            with patch.object(weekly_api, "prepare_weekly_inputs", side_effect=prepare):
                weekly_api._run_refresh_review(job_id)
            status = self.client.get(f"/api/weekly/jobs/{job_id}").json()
            self.assertEqual(status["status"], "waiting_review")
            self.assertEqual(status["preparation_dir"], "preparation-refreshed")
            self.assertNotIn("review_refresh_url", status)
            self.assertEqual(old_review.read_text(encoding="utf-8"), "old review data")
            self.assertEqual(self.client.get(status["review_data_url"]).json()["rows"][0]["quote_num"], "Q-1")
            self.assertIn("Q-1", self.client.get(status["draft_url"]).text)

    def test_current_job_cannot_refresh_review(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            job_id = self._start().json()["job_id"]
            preparation = self.root / job_id / "preparation"
            preparation.mkdir()
            (preparation / "draft-master-data.csv").write_text("quote_num\n", encoding="utf-8")
            weekly_api._update_state(job_id, status="waiting_review")
            status = self.client.get(f"/api/weekly/jobs/{job_id}").json()
            self.assertNotIn("review_refresh_url", status)
            self.assertEqual(self.client.post(f"/api/weekly/jobs/{job_id}/refresh-review").status_code, 409)

    def test_finalize_accepts_catalyst_review_without_slide_review(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit") as submit:
            job_id = self._start().json()["job_id"]
            weekly_api._update_state(job_id, status="waiting_review")
            response = self.client.post(
                f"/api/weekly/jobs/{job_id}/finalize",
                files={
                    "catalyst_review": (
                        "review.csv",
                        b"quote_num,order_number,reviewed_status\nQ1,O1,COMPLETE\n",
                        "text/csv",
                    ),
                },
            )
            self.assertEqual(response.status_code, 202)
            self.assertEqual(
                submit.call_args.args,
                (weekly_api._run_finalize, job_id, "catalyst-review-upload.csv", None, None, None, None, None),
            )
            self.assertFalse((self.root / job_id / "input" / "slide-review-upload.csv").exists())

    def test_ppt_template_is_saved_and_generated_after_mobile_workbook(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            job_id = self._start(with_ppt=True).json()["job_id"]
        job_dir = self.root / job_id
        self.assertEqual((job_dir / "input" / "previous-mobile.pptx").read_bytes(), b"pptx")

        def finalize(**kwargs):
            output = kwargs["output_dir"]
            output.mkdir(parents=True)
            for name in ("audit.json", "exceptions.csv", "review-decisions.csv", "bq-source-decisions.csv"):
                (output / name).write_text("", encoding="utf-8")
            return {"invalid_master_services": [], "issue_counts": {}}

        def workbook(**kwargs):
            kwargs["output_workbook"].write_bytes(b"xlsx")

        def presentation(**kwargs):
            self.assertEqual(kwargs["mobile_workbook"].read_bytes(), b"xlsx")
            kwargs["output_pptx"].write_bytes(b"pptx-result")

        (job_dir / "input" / "slide.csv").write_text("", encoding="utf-8")
        (job_dir / "input" / "presentation.csv").write_text(
            "chart,series,label,value,reviewed_by,review_reason\n",
            encoding="utf-8",
        )
        with (
            patch.object(weekly_api, "finalize_weekly_data", side_effect=finalize),
            patch.object(weekly_api, "write_weekly_workbook", side_effect=workbook),
            patch.object(weekly_api, "write_weekly_presentation", side_effect=presentation) as writer,
        ):
            weekly_api._run_finalize(
                job_id, "review.csv", None, "slide.csv", None, "presentation.csv"
            )
        self.assertEqual(writer.call_count, 1)
        self.assertEqual(
            writer.call_args.kwargs["presentation_review"],
            job_dir / "input" / "presentation.csv",
        )
        state = weekly_api._read_state(job_id)
        self.assertEqual(state["status"], "succeeded")
        self.assertIn("Mobile Services W38 - 2026 (Weekly).pptx", state["output_files"])
        with zipfile.ZipFile(job_dir / "weekly-mobile-output.zip") as archive:
            self.assertEqual(archive.read("Mobile Services W38 - 2026 (Weekly).pptx"), b"pptx-result")
            self.assertIn("presentation-review.csv", archive.namelist())

    def test_finalize_accepts_separate_approved_corrections(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit") as submit:
            response = self._start()
            job_id = response.json()["job_id"]
            weekly_api._update_state(job_id, status="waiting_review")
            decisions = (
                b"master_row,quote_num,order_number,reviewed_status,reviewed_by,review_reason\n"
                b"2,Q1,O1,CANCEL,Reviewer,Confirmed in Catalyst\n"
            )
            response = self.client.post(
                f"/api/weekly/jobs/{job_id}/finalize",
                files={
                    "catalyst_review": ("review.xlsx", b"checkpoint", "application/octet-stream"),
                    "manual_decisions": ("decisions.csv", decisions, "text/csv"),
                    "slide_review": ("slide.csv", b"workbook,cell,value,reviewed_by,review_reason\n", "text/csv"),
                },
            )
            self.assertEqual(response.status_code, 202)
            self.assertEqual(
                (self.root / job_id / "input" / "manual-decisions-upload.csv").read_bytes(),
                decisions,
            )
            self.assertEqual(
                submit.call_args.args,
                (weekly_api._run_finalize, job_id, "catalyst-review-upload.xlsx", "manual-decisions-upload.csv", "slide-review-upload.csv", None, None, None),
            )

    def test_business_report_week_is_required_and_validated(self) -> None:
        files = {
            "bq_csv": ("source.csv", b"quote_num\nQ1\n", "text/csv"),
            "previous_master": ("master.xlsx", b"master", "application/octet-stream"),
            "previous_mobile": ("mobile.xlsx", b"mobile", "application/octet-stream"),
            "holiday_workbook": ("holidays.xlsx", b"holidays", "application/octet-stream"),
        }
        with patch.object(weekly_api.JOB_EXECUTOR, "submit") as submit:
            response = self.client.post(
                "/api/weekly/jobs",
                files=files,
                data={"report_date": "2026-09-24", "report_week": "54"},
            )
            self.assertEqual(response.status_code, 400)
            missing = self.client.post(
                "/api/weekly/jobs", files=files, data={"report_date": "2026-09-24"}
            )
            self.assertEqual(missing.status_code, 422)
            submit.assert_not_called()
            self.assertFalse(list(self.root.iterdir()))

    def test_report_month_is_validated_before_upload(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit") as submit:
            response = self._start(report_month="2026-10")
            self.assertEqual(response.status_code, 400)
            response = self._start(report_month="2026-13")
            self.assertEqual(response.status_code, 400)
            submit.assert_not_called()
            self.assertFalse(list(self.root.iterdir()))

    def test_conflicting_review_blocks_workbook_generation(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            response = self._start()
            job_id = response.json()["job_id"]
            weekly_api._update_state(job_id, status="queued_finalization")

            def finalize(**kwargs):
                output_dir = kwargs["output_dir"]
                output_dir.mkdir(parents=True)
                (output_dir / "exceptions.csv").write_text("conflict", encoding="utf-8")
                return {"invalid_master_services": [], "issue_counts": {"conflicting_catalyst_status": 1}}

            with (
                patch.object(weekly_api, "finalize_weekly_data", side_effect=finalize),
                patch.object(weekly_api, "write_weekly_workbook") as writer,
            ):
                weekly_api._run_finalize(job_id, "review.csv", None, "slide-review-upload.csv")
            writer.assert_not_called()
            state = self.client.get(f"/api/weekly/jobs/{job_id}").json()
            self.assertEqual(state["status"], "review_failed")
            self.assertEqual(self.client.get(state["final_exceptions_url"]).status_code, 200)
            with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
                retry = self.client.post(
                    f"/api/weekly/jobs/{job_id}/finalize",
                    files={
                        "catalyst_review": ("corrected.csv", b"quote_num,order_number,reviewed_status\n", "text/csv"),
                        "slide_review": ("slide.csv", b"workbook,cell,value,reviewed_by,review_reason\n", "text/csv"),
                    },
                )
            self.assertEqual(retry.status_code, 202)


if __name__ == "__main__":
    unittest.main()
