import json
import csv
import io
import unittest
import uuid
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from fastapi.testclient import TestClient

from app import app
from weekly import api as weekly_api
from weekly.io import iter_master_csv
from weekly.rules import build_master_record


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
               report_month: str | None = None, output_services: str = "mobile"):
        files = {
            "bq_csv": ("source.csv", b"quote_num\nQ1\n", "text/csv"),
            "previous_master": ("master.xlsx", b"master", "application/octet-stream"),
            "previous_mobile": ("mobile.xlsx", b"mobile", "application/octet-stream"),
            "holiday_workbook": ("holidays.xlsx", b"holidays", "application/octet-stream"),
        }
        if with_ppt:
            files["previous_ppt"] = ("mobile.pptx", b"pptx", "application/octet-stream")
        if output_services in {"m2m", "both"}:
            files["previous_m2m"] = ("m2m.xlsx", b"m2m", "application/octet-stream")
            if with_ppt:
                files["previous_m2m_ppt"] = ("m2m.pptx", b"m2m-pptx", "application/octet-stream")
        if output_services == "m2m":
            files.pop("previous_mobile")
        return self.client.post(
            "/api/weekly/jobs",
            files=files,
            data={"report_date": "2026-09-24", "report_week": "38",
                  "report_month": report_month or "2026-09", "job_id": job_id or uuid.uuid4().hex,
                  "output_services": output_services},
        )

    def test_selected_report_templates_are_required_before_starting_job(self) -> None:
        files = {
            "bq_csv": ("bq.csv", b"source"), "previous_master": ("master.xlsx", b"master"),
            "holiday_workbook": ("holidays.xlsx", b"holidays"),
        }
        with patch.object(weekly_api.JOB_EXECUTOR, "submit") as submit:
            for selection in ("mobile", "m2m", "both", "unknown"):
                response = self.client.post("/api/weekly/jobs", files=files,
                    data={"report_date": "2026-09-24", "report_week": "38", "output_services": selection})
                self.assertEqual(response.status_code, 400, response.text)
            submit.assert_not_called()
        self.assertFalse(list(self.root.iterdir()))

    def test_legacy_requests_without_output_selection_still_default_to_mobile(self) -> None:
        files = {
            "bq_csv": ("bq.csv", b"source"), "previous_master": ("master.xlsx", b"master"),
            "previous_mobile": ("mobile.xlsx", b"mobile"),
            "holiday_workbook": ("holidays.xlsx", b"holidays"),
        }
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            response = self.client.post("/api/weekly/jobs", files=files,
                data={"report_date": "2026-09-24", "report_week": "38"})
        self.assertEqual(response.status_code, 202, response.text)
        state = weekly_api._read_state(response.json()["job_id"])
        self.assertEqual(state["output_services"], "mobile")

    def test_m2m_and_combined_outputs_use_one_master_and_only_selected_writers(self) -> None:
        for selection in ("m2m", "both"):
            with self.subTest(selection=selection):
                with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
                    response = self._start(with_ppt=True, output_services=selection)
                self.assertEqual(response.status_code, 202, response.text)
                job_id = response.json()["job_id"]
                job_dir = self.root / job_id

                def finalize(**kwargs):
                    kwargs["output_dir"].mkdir(exist_ok=True)
                    return {"invalid_master_services": [], "issue_counts": {}}

                def workbook(**kwargs):
                    kwargs["output_workbook"].write_bytes(b"xlsx")

                def ppt(**kwargs):
                    kwargs["output_pptx"].write_bytes(b"pptx")

                with (
                    patch.object(weekly_api, "finalize_weekly_data", side_effect=finalize) as finalizer,
                    patch.object(weekly_api, "write_weekly_workbook", side_effect=workbook) as writer,
                    patch.object(weekly_api, "write_weekly_presentation", side_effect=ppt) as mobile_ppt,
                    patch.object(weekly_api, "write_m2m_presentation", side_effect=ppt) as m2m_ppt,
                ):
                    weekly_api._run_finalize(job_id, "review.csv", None, None)
                state = weekly_api._read_state(job_id)
                self.assertEqual(state["status"], "succeeded", state.get("error"))
                finalizer.assert_called_once()
                self.assertEqual(writer.call_count, 3 if selection == "both" else 2)
                self.assertEqual(mobile_ppt.call_count, int(selection == "both"))
                m2m_ppt.assert_called_once()
                m2m_call = writer.call_args.kwargs
                self.assertEqual(m2m_call["report_type"], "m2m")
                self.assertEqual(m2m_call["data_csv"].name, "m2m-data.csv")
                self.assertEqual(m2m_call["template_workbook"].name, "previous-m2m.xlsx")
                download = self.client.get(f"/api/weekly/jobs/{job_id}/download")
                self.assertEqual(download.status_code, 200)
                with zipfile.ZipFile(io.BytesIO(download.content)) as archive:
                    self.assertEqual(len(archive.namelist()), 5 if selection == "both" else 3)
                    self.assertEqual(set(archive.namelist()), set(state["output_files"]))
                    self.assertTrue(all(Path(name).suffix in {".xlsx", ".pptx"} for name in archive.namelist()))

    def test_retry_reuses_only_verified_artifacts_from_unchanged_inputs(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            job_id = self._start(with_ppt=True, output_services="both").json()["job_id"]
        job_dir = self.root / job_id
        (job_dir / "input" / "review.csv").write_text("original review", encoding="utf-8")

        def finalize(**kwargs):
            kwargs["output_dir"].mkdir(exist_ok=True)
            return {"invalid_master_services": [], "issue_counts": {}}

        def workbook(**kwargs):
            kwargs["output_workbook"].write_bytes(b"xlsx")

        def ppt(**kwargs):
            kwargs["output_pptx"].write_bytes(b"pptx")

        with (
            patch.object(weekly_api, "finalize_weekly_data", side_effect=finalize),
            patch.object(weekly_api, "write_weekly_workbook", side_effect=workbook) as writer,
            patch.object(weekly_api, "write_weekly_presentation", side_effect=ppt) as mobile_ppt,
            patch.object(weekly_api, "write_m2m_presentation", side_effect=RuntimeError("simulated PPT failure")),
        ):
            weekly_api._run_finalize(job_id, "review.csv", None, None)
            self.assertEqual(weekly_api._read_state(job_id)["status"], "review_failed")
            self.assertEqual(writer.call_count, 3)
            writer.reset_mock()
            with patch.object(weekly_api, "write_m2m_presentation", side_effect=ppt):
                weekly_api._run_finalize(job_id, "review.csv", None, None)
                self.assertEqual(weekly_api._read_state(job_id)["status"], "succeeded")
                writer.assert_not_called()
                self.assertEqual(mobile_ppt.call_count, 1)
                # A saved filename alone is insufficient evidence of a valid artifact.
                m2m_name = weekly_api._m2m_filenames(weekly_api._read_state(job_id))[0]
                (job_dir / "output" / m2m_name).write_bytes(b"tampered")
                weekly_api._run_finalize(job_id, "review.csv", None, None)
                writer.assert_called_once()
                writer.reset_mock()
                (job_dir / "input" / "review.csv").write_text("updated review", encoding="utf-8")
                weekly_api._run_finalize(job_id, "review.csv", None, None)
                self.assertEqual(writer.call_count, 3)

    def test_web_finalization_preserves_history_and_fills_missing_revenue(self) -> None:
        source = {
            "quote_num": "Q1", "order_number": "O1", "order_status": "Complete",
            "service_group": "Mobile", "FAB_Upload": "2026-09-20 09:30:00",
            "cash_date": "2026-09-20", "order_completion_date": "2026-09-21",
            "otc": "100", "mrc": "50",
        }
        previous = build_master_record(source, row_number=1)
        previous.update({"LEADTIME FAB-CASH": 1, "OTC+MRC": None, "OTC+MRC (inMio)": None})

        def workbook(**kwargs):
            rows = list(iter_master_csv(kwargs["data_csv"]))
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["LEADTIME FAB-CASH"], 1)
            self.assertEqual(rows[0]["OTC+MRC"], 150)
            self.assertEqual(str(rows[0]["OTC+MRC (inMio)"]), "0.00015")
            kwargs["output_workbook"].write_bytes(b"writer-is-tested-separately")

        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            job_id = self._start().json()["job_id"]
            with (
                patch("weekly.prepare.iter_master_rows", side_effect=lambda _: iter([previous])),
                patch("weekly.prepare.iter_bq_rows", side_effect=lambda _: iter([source])),
                patch("weekly.prepare.load_holidays", return_value=set()),
            ):
                weekly_api._run_prepare(job_id)
            response = self.client.post(
                f"/api/weekly/jobs/{job_id}/finalize",
                files={"catalyst_review": ("review.csv", b"quote_num,order_number,reviewed_status\n", "text/csv")},
            )
            self.assertEqual(response.status_code, 202)
            with (
                patch("weekly.finalize.iter_master_rows", side_effect=lambda _: iter([previous])),
                patch("weekly.finalize.iter_bq_rows", side_effect=lambda _: iter([source])),
                patch("weekly.finalize.load_holidays", return_value=set()),
                patch.object(weekly_api, "write_weekly_workbook", side_effect=workbook) as writer,
            ):
                weekly_api._run_finalize(job_id, "catalyst-review-upload.csv", None, None)
            self.assertEqual(writer.call_count, 2)
            status = self.client.get(f"/api/weekly/jobs/{job_id}").json()
            self.assertEqual(status["status"], "succeeded", status.get("error"))
            self.assertEqual(status["final_audit"]["pending_catalyst_rows"], 0)
            self.assertEqual(status["final_audit"]["metric_calculation_variances"], {"LEADTIME FAB-CASH": 1})
            with zipfile.ZipFile(io.BytesIO(self.client.get(f"/api/weekly/jobs/{job_id}/download").content)) as archive:
                self.assertEqual(set(archive.namelist()), set(status["output_files"]))

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
            download = self.client.get(status["download_url"])
            self.assertEqual(download.status_code, 200)
            expected_files = [
                "Master Raw Mobile Service - W38 2026 old format.xlsx",
                "Master Raw Mobile Service - W38 2026.xlsx",
            ]
            self.assertEqual(status["output_files"], expected_files)
            with zipfile.ZipFile(io.BytesIO(download.content)) as archive:
                self.assertEqual(archive.namelist(), expected_files)
            with zipfile.ZipFile(job_dir / "weekly-mobile-output.zip") as archive:
                self.assertEqual(archive.namelist(), expected_files)
            for name in ("audit.json", "exceptions.csv", "review-decisions.csv", "bq-source-decisions.csv"):
                self.assertTrue((job_dir / "output" / name).is_file())
            self.assertTrue((job_dir / "input" / "pivot-review-upload.csv").is_file())
            self.assertTrue((job_dir / "input" / "slide-review-upload.csv").is_file())
            self.assertEqual(self.client.get(status["final_exceptions_url"]).status_code, 200)
            self.assertEqual(json.loads((job_dir / "job.json").read_text())["status"], "succeeded")

    def test_download_cleans_legacy_archives_without_regenerating_reports(self) -> None:
        for with_ppt in (False, True):
            with self.subTest(with_ppt=with_ppt):
                with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
                    job_id = self._start(with_ppt=with_ppt).json()["job_id"]
                job_dir = self.root / job_id
                output = job_dir / "output"
                output.mkdir()
                expected_files = list(weekly_api._report_filenames(weekly_api._read_state(job_id)))
                if not with_ppt:
                    expected_files.pop()
                for name in expected_files:
                    (output / name).write_bytes(name.encode("utf-8"))
                (output / "audit.json").write_text("original audit", encoding="utf-8")
                (job_dir / "input" / "approved.csv").write_text("original review", encoding="utf-8")
                path = job_dir / "weekly-mobile-output.zip"
                with zipfile.ZipFile(path, "w") as archive:
                    for name in expected_files:
                        archive.write(output / name, arcname=name)
                    archive.write(output / "audit.json", arcname="audit.json")
                    archive.write(job_dir / "input" / "approved.csv", arcname="approved-corrections.csv")
                weekly_api._update_state(
                    job_id, status="succeeded",
                    output_files=[*expected_files, "audit.json", "approved-corrections.csv"],
                )
                status = self.client.get(f"/api/weekly/jobs/{job_id}").json()
                self.assertEqual(status["output_files"], expected_files)
                with (
                    patch.object(weekly_api, "finalize_weekly_data") as finalizer,
                    patch.object(weekly_api, "write_weekly_workbook") as workbook,
                    patch.object(weekly_api, "write_weekly_presentation") as presentation,
                ):
                    download = self.client.get(status["download_url"])
                self.assertEqual(download.status_code, 200)
                finalizer.assert_not_called()
                workbook.assert_not_called()
                presentation.assert_not_called()
                with zipfile.ZipFile(io.BytesIO(download.content)) as archive:
                    self.assertEqual(archive.namelist(), expected_files)
                    for name in expected_files:
                        self.assertEqual(archive.read(name), name.encode("utf-8"))
                self.assertEqual((output / "audit.json").read_text(), "original audit")
                self.assertEqual((job_dir / "input" / "approved.csv").read_text(), "original review")
                with patch.object(weekly_api, "_write_report_archive") as writer:
                    self.assertEqual(self.client.get(status["download_url"]).status_code, 200)
                writer.assert_not_called()

    def test_legacy_download_refuses_audit_only_archive_when_reports_are_missing(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            job_id = self._start().json()["job_id"]
        weekly_api._update_state(job_id, status="succeeded", output_files=["audit.json"])
        path = self.root / job_id / "weekly-mobile-output.zip"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("audit.json", "internal audit")
        response = self.client.get(f"/api/weekly/jobs/{job_id}/download")
        self.assertEqual(response.status_code, 404)
        with zipfile.ZipFile(path) as archive:
            self.assertEqual(archive.read("audit.json"), b"internal audit")

    def test_archive_write_failure_preserves_previous_download(self) -> None:
        with patch.object(weekly_api.JOB_EXECUTOR, "submit"):
            job_id = self._start().json()["job_id"]
        job_dir = self.root / job_id
        path = job_dir / "weekly-mobile-output.zip"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("old.xlsx", "previous report")
        previous = path.read_bytes()
        with patch.object(zipfile.ZipFile, "write", side_effect=OSError("disk full")):
            with self.assertRaisesRegex(OSError, "disk full"):
                weekly_api._write_report_archive(job_dir, ["new.xlsx"])
        self.assertEqual(path.read_bytes(), previous)
        self.assertFalse(path.with_suffix(".zip.tmp").exists())

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
        for name in ("manual.csv", "source.csv"):
            (job_dir / "input" / name).write_text("approved review", encoding="utf-8")
        with (
            patch.object(weekly_api, "finalize_weekly_data", side_effect=finalize),
            patch.object(weekly_api, "write_weekly_workbook", side_effect=workbook),
            patch.object(weekly_api, "write_weekly_presentation", side_effect=presentation) as writer,
        ):
            weekly_api._run_finalize(
                job_id, "review.csv", "manual.csv", "slide.csv", None, "presentation.csv", "source.csv"
            )
        self.assertEqual(writer.call_count, 1)
        self.assertEqual(
            writer.call_args.kwargs["presentation_review"],
            job_dir / "input" / "presentation.csv",
        )
        state = weekly_api._read_state(job_id)
        self.assertEqual(state["status"], "succeeded")
        self.assertIn("Mobile Services W38 - 2026 (Weekly).pptx", state["output_files"])
        expected_files = list(weekly_api._report_filenames(state))
        self.assertEqual(state["output_files"], expected_files)
        with zipfile.ZipFile(job_dir / "weekly-mobile-output.zip") as archive:
            self.assertEqual(archive.read("Mobile Services W38 - 2026 (Weekly).pptx"), b"pptx-result")
            self.assertEqual(archive.namelist(), expected_files)
        for name in ("manual.csv", "source.csv", "presentation.csv"):
            self.assertTrue((job_dir / "input" / name).is_file())

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
