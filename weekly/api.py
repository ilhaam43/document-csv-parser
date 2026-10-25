"""Two-stage weekly report API with a manual Catalyst review checkpoint."""

from __future__ import annotations

import csv
import hashlib
import itertools
import json
import logging
import threading
import time
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from .dates import parse_report_month
from .finalize import finalize_weekly_data
from .prepare import prepare_weekly_inputs
from .presentation import write_weekly_presentation
from .m2m_presentation import write_m2m_presentation
from .rules import normalize_order_status
from .workbook import write_weekly_workbook


logger = logging.getLogger(__name__)
JOB_ROOT = Path(__file__).resolve().parents[1] / "generated" / "weekly" / "api-jobs"
MAX_UPLOAD_BYTES = 300 * 1024 * 1024
JOB_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="weekly")
STATE_LOCK = threading.RLock()
router = APIRouter(prefix="/api/weekly", tags=["Weekly"])


def _job_dir(job_id: str) -> Path:
    try:
        return JOB_ROOT / uuid.UUID(job_id).hex
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="Weekly job not found.") from exc


def _preparation_dir(job_id: str, state: dict | None = None) -> Path:
    state = state or _read_state(job_id)
    directory = state.get("preparation_dir", "preparation")
    if directory not in {"preparation", "preparation-refreshed"}:
        raise HTTPException(status_code=500, detail="Weekly preparation directory is invalid.")
    return _job_dir(job_id) / directory


def _read_state(job_id: str) -> dict:
    path = _job_dir(job_id) / "job.json"
    try:
        with path.open(encoding="utf-8") as handle:
            return json.load(handle)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail="Weekly job not found.") from exc


def _write_state(job_id: str, state: dict) -> None:
    path = _job_dir(job_id) / "job.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".json.tmp")
    with temporary.open("w", encoding="utf-8") as handle:
        json.dump(state, handle, indent=2)
    temporary.replace(path)


def _update_state(job_id: str, **updates: object) -> dict:
    with STATE_LOCK:
        state = _read_state(job_id)
        state.update(updates, updated_at=time.time())
        _write_state(job_id, state)
        return state


def _report_filenames(state: dict) -> tuple[str, str, str]:
    week = int(state["report_week"])
    year = parse_report_month(state.get("report_month", state["report_date"][:7])).year
    return (
        f"Master Raw Mobile Service - W{week} {year} old format.xlsx",
        f"Master Raw Mobile Service - W{week} {year}.xlsx",
        f"Mobile Services W{week} - {year} (Weekly).pptx",
    )


def _output_services(state: dict) -> tuple[str, ...]:
    return {"mobile": ("mobile",), "m2m": ("m2m",), "both": ("mobile", "m2m")}[state.get("output_services", "mobile")]


def _m2m_filenames(state: dict) -> tuple[str, str]:
    week = int(state["report_week"])
    year = parse_report_month(state.get("report_month", state["report_date"][:7])).year
    return f"Master Raw M2M Service - W{week}.xlsx", f"M2M Services W{week}- {year} (Weekly).pptx"


def _allowed_report_files(state: dict) -> list[str]:
    master, mobile, ppt = _report_filenames(state)
    return [master] + ([mobile, ppt] if "mobile" in _output_services(state) else []) + (
        list(_m2m_filenames(state)) if "m2m" in _output_services(state) else [])


def _file_digest(path: Path) -> str | None:
    if not path.is_file():
        return None
    with path.open("rb") as handle:
        return hashlib.file_digest(handle, "sha256").hexdigest()


def _generation_signature(job_dir: Path, state: dict, reviews: tuple[str | None, ...]) -> str:
    files = ["bq.csv", "previous-master.xlsx", "holidays.xlsx", *filter(None, reviews)]
    for service in _output_services(state):
        files.extend([f"previous-{service}.xlsx", f"previous-{service}.pptx"])
    payload = {name: _file_digest(job_dir / "input" / name) for name in sorted(set(files))}
    # A code update invalidates saved artifacts even when uploads are unchanged.
    payload["engine"] = {p.name: _file_digest(p) for p in Path(__file__).parent.glob("*.py")}
    payload["period"] = [state["report_date"], state.get("report_month"), state["report_week"], _output_services(state)]
    return hashlib.sha256(json.dumps(payload, sort_keys=True).encode()).hexdigest()


def _write_report_archive(job_dir: Path, files: list[str]) -> Path:
    path = job_dir / "weekly-mobile-output.zip"
    temporary = path.with_suffix(".zip.tmp")
    try:
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name in files:
                archive.write(job_dir / "output" / name, arcname=name)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
    return path


async def _save_upload(upload: UploadFile, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    size = 0
    with path.open("wb") as handle:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail="Weekly upload exceeds 300 MB.")
            handle.write(chunk)


def _require_extension(upload: UploadFile, field: str, extensions: set[str]) -> None:
    if not upload.filename or Path(upload.filename).suffix.lower() not in extensions:
        raise HTTPException(
            status_code=400,
            detail=f"{field} must be one of: {', '.join(sorted(extensions))}.",
        )


def _prepared_review_rows(job_id: str) -> tuple[list[str], list[dict[str, str]]]:
    path = _preparation_dir(job_id) / "catalyst-review.csv"
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Catalyst review is not ready.")
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        columns = reader.fieldnames or []
        if not {"master_row", "quote_num", "order_number", "reviewed_status"}.issubset(columns):
            raise HTTPException(status_code=500, detail="Prepared Catalyst review has an invalid schema.")
        rows = list(itertools.islice(reader, 5001))
    if len(rows) > 5000:
        raise HTTPException(
            status_code=409,
            detail="This older job has too many review rows for direct entry. Refresh its Catalyst review from saved inputs.",
        )
    keys = [row["master_row"] for row in rows]
    if any(not key or not key.isdecimal() for key in keys) or len(keys) != len(set(keys)):
        raise HTTPException(status_code=500, detail="Prepared Catalyst review has invalid row identifiers.")
    return columns, rows


def _inline_review_rows(job_id: str, submitted: str) -> tuple[list[str], list[dict[str, str]]]:
    if len(submitted) > 1024 * 1024:
        raise HTTPException(status_code=413, detail="Catalyst review is too large.")
    try:
        decisions = json.loads(submitted)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="Catalyst statuses must be valid JSON.") from exc
    if not isinstance(decisions, dict):
        raise HTTPException(status_code=400, detail="Catalyst statuses must be keyed by master_row.")
    columns, rows = _prepared_review_rows(job_id)
    expected = {row["master_row"] for row in rows}
    if decisions.keys() != expected:
        raise HTTPException(status_code=400, detail="Every prepared Catalyst row needs one status; refresh the review list.")
    allowed = {"In Progress", "COMPLETE", "CANCEL"}
    for row in rows:
        status = decisions[row["master_row"]]
        if not isinstance(status, str) or not status.strip() or normalize_order_status(status) not in allowed:
            raise HTTPException(status_code=400, detail=f"Invalid Catalyst status for Master row {row['master_row']}.")
        row["reviewed_status"] = normalize_order_status(status)
    return columns, rows


def _prepared_source_rows(job_id: str) -> tuple[list[str], list[dict[str, str]]]:
    path = _preparation_dir(job_id) / "source-review.csv"
    if not path.is_file():
        raise HTTPException(status_code=404, detail="BQ source review is not ready.")
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        columns = reader.fieldnames or []
        required = {"master_row", "quote_num", "order_number", "reviewed_source"}
        if not required.issubset(columns):
            raise HTTPException(status_code=500, detail="Prepared source review has an invalid schema.")
        rows = list(itertools.islice(reader, 5001))
    if len(rows) > 5000:
        raise HTTPException(status_code=409, detail="Too many BQ source conflicts for direct review.")
    return columns, rows


def _inline_source_rows(job_id: str, submitted: str) -> tuple[list[str], list[dict[str, str]]]:
    if len(submitted) > 1024 * 1024:
        raise HTTPException(status_code=413, detail="BQ source review is too large.")
    try:
        decisions = json.loads(submitted)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="BQ source choices must be valid JSON.") from exc
    if not isinstance(decisions, dict):
        raise HTTPException(status_code=400, detail="BQ source choices must be keyed by master_row.")
    columns, rows = _prepared_source_rows(job_id)
    expected = {row["master_row"] for row in rows}
    if not decisions.keys() <= expected or any(choice not in {"order", "quote"} for choice in decisions.values()):
        raise HTTPException(status_code=400, detail="BQ source overrides must match prepared conflicts and choose order or quote.")
    for row in rows:
        row["reviewed_source"] = decisions.get(row["master_row"], "")
    return columns, rows


def _run_prepare(job_id: str) -> None:
    job_dir = _job_dir(job_id)
    _update_state(job_id, status="preparing", stage="Preparing Catalyst review")
    try:
        state = _read_state(job_id)
        audit = prepare_weekly_inputs(
            bq_csv=job_dir / "input" / "bq.csv",
            previous_master=job_dir / "input" / "previous-master.xlsx",
            holiday_workbook=job_dir / "input" / "holidays.xlsx",
            report_date=date.fromisoformat(state["report_date"]),
            report_month=parse_report_month(state.get("report_month", state["report_date"][:7])),
            output_dir=job_dir / "preparation",
        )
        if audit["generated_master_rows"] != audit["selected_source_rows"]:
            raise RuntimeError("Some selected BQ rows could not be built; review preparation exceptions.")
        _update_state(
            job_id,
            status="waiting_review",
            stage="Waiting for Catalyst review",
            preparation_dir="preparation",
            preparation_audit=audit,
        )
    except Exception as exc:
        logger.exception("Weekly preparation failed for %s", job_id)
        _update_state(job_id, status="failed", stage="Preparation failed", error=str(exc))


def _run_refresh_review(job_id: str) -> None:
    job_dir = _job_dir(job_id)
    try:
        state = _read_state(job_id)
        audit = prepare_weekly_inputs(
            bq_csv=job_dir / "input" / "bq.csv",
            previous_master=job_dir / "input" / "previous-master.xlsx",
            holiday_workbook=job_dir / "input" / "holidays.xlsx",
            report_date=date.fromisoformat(state["report_date"]),
            report_month=parse_report_month(state.get("report_month", state["report_date"][:7])),
            output_dir=job_dir / "preparation-refreshed",
        )
        if audit["generated_master_rows"] != audit["selected_source_rows"]:
            raise RuntimeError("Some BQ rows could not be built; inspect preparation exceptions.")
        if audit["pending_catalyst_rows"] > 5000:
            raise RuntimeError("Catalyst review still has too many rows for direct entry.")
        _update_state(
            job_id,
            status="waiting_review",
            stage="Waiting for Catalyst review",
            preparation_dir="preparation-refreshed",
            preparation_audit=audit,
            error=None,
        )
    except Exception as exc:
        logger.exception("Weekly review refresh failed for %s", job_id)
        _update_state(job_id, status="review_failed", stage="Review refresh failed", error=str(exc))


def _run_finalize(
    job_id: str,
    review_filename: str,
    decisions_filename: str | None,
    slide_filename: str | None,
    pivot_filename: str | None = None,
    presentation_filename: str | None = None,
    source_filename: str | None = None,
) -> None:
    job_dir = _job_dir(job_id)
    state = _read_state(job_id)
    report_date = date.fromisoformat(state["report_date"])
    report_month = parse_report_month(state.get("report_month", state["report_date"][:7]))
    report_week = int(state["report_week"])
    output_dir = job_dir / "output"
    try:
        signature = _generation_signature(job_dir, state, (review_filename, decisions_filename,
            slide_filename, pivot_filename, presentation_filename, source_filename))
        completed = dict(state.get("completed_artifacts", {})) if state.get("generation_signature") == signature else {}
        _update_state(job_id, status="finalizing", stage="Finalizing weekly Master data",
                      generation_signature=signature, completed_artifacts=completed)
        audit = finalize_weekly_data(
            bq_csv=job_dir / "input" / "bq.csv",
            previous_master=job_dir / "input" / "previous-master.xlsx",
            catalyst_review=job_dir / "input" / review_filename,
            source_review=(job_dir / "input" / source_filename) if source_filename else None,
            manual_decisions=(job_dir / "input" / decisions_filename) if decisions_filename else None,
            holiday_workbook=job_dir / "input" / "holidays.xlsx",
            report_date=report_date,
            report_month=report_month,
            output_dir=output_dir,
        )
        if audit["invalid_master_services"] or audit["issue_counts"].get("conflicting_catalyst_status"):
            raise RuntimeError("Final data has invalid services or conflicting Catalyst statuses; review exceptions.csv.")

        def artifact(name: str, stage: str, writer, **kwargs):
            _update_state(job_id, status="generating", stage=stage)
            path = output_dir / name
            if not completed.get(name) or completed[name] != _file_digest(path):
                writer(**kwargs)
                digest = _file_digest(path)
                if digest is None:
                    raise RuntimeError(f"Report writer did not produce {name}.")
                completed[name] = digest
                _update_state(job_id, completed_artifacts=completed)

        master_name, mobile_name, ppt_name = _report_filenames(state)
        artifact(master_name, "Generating Master workbook", write_weekly_workbook,
            data_csv=output_dir / "master-final-data.csv",
            template_workbook=job_dir / "input" / "previous-master.xlsx",
            output_workbook=output_dir / master_name,
            report_week=report_week,
            report_date=report_month,
            slide_review=(job_dir / "input" / slide_filename) if slide_filename else None,
            pivot_review=(job_dir / "input" / pivot_filename) if pivot_filename else None,
        )
        files = [master_name]
        if "mobile" in _output_services(state):
            artifact(mobile_name, "Generating Mobile workbook", write_weekly_workbook,
                data_csv=output_dir / "mobile-data.csv",
                template_workbook=job_dir / "input" / "previous-mobile.xlsx",
                output_workbook=output_dir / mobile_name,
                mobile=True,
                report_week=report_week,
                report_date=report_month,
                slide_review=(job_dir / "input" / slide_filename) if slide_filename else None,
                pivot_review=(job_dir / "input" / pivot_filename) if pivot_filename else None,
            )
            files.append(mobile_name)
            ppt_template = job_dir / "input" / "previous-mobile.pptx"
            if ppt_template.is_file():
                artifact(ppt_name, "Generating Mobile presentation", write_weekly_presentation,
                    mobile_workbook=output_dir / mobile_name,
                    template_pptx=ppt_template,
                    output_pptx=output_dir / ppt_name,
                    report_date=report_date,
                    report_month=report_month,
                    report_week=report_week,
                    presentation_review=(job_dir / "input" / presentation_filename)
                    if presentation_filename else None,
                )
                files.append(ppt_name)
        if "m2m" in _output_services(state):
            m2m_name, m2m_ppt = _m2m_filenames(state)
            artifact(m2m_name, "Generating M2M workbook", write_weekly_workbook,
                data_csv=output_dir / "m2m-data.csv", template_workbook=job_dir / "input" / "previous-m2m.xlsx",
                output_workbook=output_dir / m2m_name, report_type="m2m", report_week=report_week,
                report_date=report_month)
            files.append(m2m_name)
            if (job_dir / "input" / "previous-m2m.pptx").is_file():
                artifact(m2m_ppt, "Generating M2M presentation", write_m2m_presentation,
                    m2m_workbook=output_dir / m2m_name, template_pptx=job_dir / "input" / "previous-m2m.pptx",
                    output_pptx=output_dir / m2m_ppt, report_date=report_date, report_month=report_month,
                    report_week=report_week)
                files.append(m2m_ppt)
        _write_report_archive(job_dir, files)
        warnings = ["Verify SLIDE GH MEETING totals against the approved source before distribution."]
        pending = audit.get("pending_catalyst_rows", 0)
        if pending:
            warnings.insert(0, f"{pending} Catalyst statuses remain unresolved; verify before distribution.")
        issue_count = sum(audit.get("issue_counts", {}).values())
        if issue_count:
            warnings.append(
                f"{issue_count} data exceptions remain; use Download Exceptions CSV "
                "for details before distribution."
            )
        _update_state(
            job_id,
            status="succeeded",
            stage="Complete",
            final_audit=audit,
            warnings=warnings,
            output_files=files,
        )
    except Exception as exc:
        logger.exception("Weekly finalization failed for %s", job_id)
        _update_state(job_id, status="review_failed", stage="Generation failed", error=str(exc))


@router.post("/jobs", status_code=202)
async def start_weekly_job(
    bq_csv: UploadFile = File(...),
    previous_master: UploadFile = File(...),
    previous_mobile: UploadFile | None = File(default=None),
    previous_ppt: UploadFile | None = File(default=None),
    previous_m2m: UploadFile | None = File(default=None),
    previous_m2m_ppt: UploadFile | None = File(default=None),
    holiday_workbook: UploadFile = File(...),
    report_date: str = Form(...),
    report_week: int = Form(...),
    report_month: str | None = Form(default=None),
    job_id: str | None = Form(default=None),
    output_services: str = Form(default="mobile"),
) -> dict:
    _require_extension(bq_csv, "bq_csv", {".csv"})
    _require_extension(previous_master, "previous_master", {".xlsx"})
    if output_services not in {"mobile", "m2m", "both"}:
        raise HTTPException(status_code=400, detail="output_services must be mobile, m2m, or both.")
    selected = _output_services({"output_services": output_services})
    if "mobile" in selected and previous_mobile is None:
        raise HTTPException(status_code=400, detail="Previous Mobile workbook is required for Mobile output.")
    if "m2m" in selected and previous_m2m is None:
        raise HTTPException(status_code=400, detail="Previous M2M workbook is required for M2M output.")
    if previous_mobile is not None:
        _require_extension(previous_mobile, "previous_mobile", {".xlsx"})
    if previous_ppt is not None:
        _require_extension(previous_ppt, "previous_ppt", {".pptx"})
    if previous_m2m is not None:
        _require_extension(previous_m2m, "previous_m2m", {".xlsx"})
    if previous_m2m_ppt is not None:
        _require_extension(previous_m2m_ppt, "previous_m2m_ppt", {".pptx"})
    _require_extension(holiday_workbook, "holiday_workbook", {".xlsx"})
    try:
        parsed_report_date = date.fromisoformat(report_date)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="report_date must be YYYY-MM-DD.") from exc
    try:
        parsed_report_month = parse_report_month(report_month or report_date[:7])
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="report_month must be YYYY-MM.") from exc
    if parsed_report_month > parsed_report_date:
        raise HTTPException(status_code=400, detail="report_month cannot be after report_date.")
    if not 1 <= report_week <= 53:
        raise HTTPException(status_code=400, detail="report_week must be between 1 and 53.")
    try:
        normalized_id = uuid.UUID(job_id).hex if job_id else uuid.uuid4().hex
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="job_id must be a UUID.") from exc
    job_dir = _job_dir(normalized_id)
    with STATE_LOCK:
        if (job_dir / "job.json").exists():
            raise HTTPException(status_code=409, detail="Weekly job already exists.")
        _write_state(normalized_id, {
            "job_id": normalized_id,
            "report_date": report_date,
            "report_month": parsed_report_month.strftime("%Y-%m"),
            "report_week": report_week,
            "output_services": output_services,
            "status": "uploading",
            "stage": "Uploading files",
            "created_at": time.time(),
            "updated_at": time.time(),
        })
    try:
        for upload, name in (
            (bq_csv, "bq.csv"),
            (previous_master, "previous-master.xlsx"),
            (holiday_workbook, "holidays.xlsx"),
        ):
            await _save_upload(upload, job_dir / "input" / name)
        if previous_mobile is not None and "mobile" in selected:
            await _save_upload(previous_mobile, job_dir / "input" / "previous-mobile.xlsx")
        if previous_ppt is not None and "mobile" in selected:
            await _save_upload(previous_ppt, job_dir / "input" / "previous-mobile.pptx")
        if previous_m2m is not None and "m2m" in selected:
            await _save_upload(previous_m2m, job_dir / "input" / "previous-m2m.xlsx")
        if previous_m2m_ppt is not None and "m2m" in selected:
            await _save_upload(previous_m2m_ppt, job_dir / "input" / "previous-m2m.pptx")
    except Exception as exc:
        _update_state(normalized_id, status="failed", stage="Upload failed", error=str(exc))
        raise
    _update_state(normalized_id, status="queued", stage="Queued for preparation")
    JOB_EXECUTOR.submit(_run_prepare, normalized_id)
    return {"job_id": normalized_id, "status": "queued", "status_url": f"/api/weekly/jobs/{normalized_id}"}


@router.get("/jobs/{job_id}")
def weekly_job_status(job_id: str) -> dict:
    state = _read_state(job_id)
    response = dict(state)
    if state["status"] in {"waiting_review", "refreshing_review", "uploading_review", "queued_finalization", "finalizing", "generating", "review_failed", "succeeded"}:
        response["review_url"] = f"/api/weekly/jobs/{job_id}/review"
        response["review_data_url"] = f"/api/weekly/jobs/{job_id}/review-data"
        if (_preparation_dir(job_id, state) / "draft-master-data.csv").is_file():
            response["draft_url"] = f"/api/weekly/jobs/{job_id}/draft"
        response["slide_review_url"] = f"/api/weekly/jobs/{job_id}/slide-review"
        if (_preparation_dir(job_id, state) / "source-review.csv").is_file():
            response["source_review_url"] = f"/api/weekly/jobs/{job_id}/source-review"
            response["source_review_data_url"] = f"/api/weekly/jobs/{job_id}/source-review-data"
    if state["status"] in {"waiting_review", "review_failed"} and not (_preparation_dir(job_id, state) / "draft-master-data.csv").is_file():
        response["review_refresh_url"] = f"/api/weekly/jobs/{job_id}/refresh-review"
    if (_preparation_dir(job_id, state) / "exceptions.csv").is_file():
        response["preparation_exceptions_url"] = f"/api/weekly/jobs/{job_id}/exceptions/preparation"
    if (_job_dir(job_id) / "output" / "exceptions.csv").is_file():
        response["final_exceptions_url"] = f"/api/weekly/jobs/{job_id}/exceptions/final"
    if state["status"] == "succeeded":
        response["output_files"] = [
            name for name in _allowed_report_files(state) if name in state.get("output_files", [])
        ]
        response["download_url"] = f"/api/weekly/jobs/{job_id}/download"
    return response


@router.get("/jobs/{job_id}/review")
def download_weekly_review(job_id: str) -> FileResponse:
    _read_state(job_id)
    path = _preparation_dir(job_id) / "catalyst-review.csv"
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Catalyst review is not ready.")
    return FileResponse(path, media_type="text/csv", filename="catalyst-review.csv")


@router.get("/jobs/{job_id}/review-data")
def weekly_review_data(job_id: str) -> dict:
    _read_state(job_id)
    _, rows = _prepared_review_rows(job_id)
    return {"rows": rows}


@router.get("/jobs/{job_id}/source-review")
def download_weekly_source_review(job_id: str) -> FileResponse:
    _read_state(job_id)
    path = _preparation_dir(job_id) / "source-review.csv"
    if not path.is_file():
        raise HTTPException(status_code=404, detail="BQ source review is not ready.")
    return FileResponse(path, media_type="text/csv", filename="source-review.csv")


@router.get("/jobs/{job_id}/source-review-data")
def weekly_source_review_data(job_id: str) -> dict:
    _read_state(job_id)
    _, rows = _prepared_source_rows(job_id)
    return {"rows": rows}


@router.get("/jobs/{job_id}/draft")
def download_weekly_draft(job_id: str) -> FileResponse:
    _read_state(job_id)
    path = _preparation_dir(job_id) / "draft-master-data.csv"
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Weekly Master draft is not ready.")
    return FileResponse(path, media_type="text/csv", filename="draft-master-data.csv")


@router.get("/jobs/{job_id}/slide-review")
def download_weekly_slide_review(job_id: str) -> FileResponse:
    _read_state(job_id)
    path = _preparation_dir(job_id) / "slide-review.csv"
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Slide review is not ready.")
    return FileResponse(path, media_type="text/csv", filename="slide-review.csv")


@router.get("/jobs/{job_id}/exceptions/{stage}")
def download_weekly_exceptions(job_id: str, stage: str) -> FileResponse:
    _read_state(job_id)
    if stage not in {"preparation", "final"}:
        raise HTTPException(status_code=404, detail="Weekly exceptions not found.")
    directory = _preparation_dir(job_id) if stage == "preparation" else _job_dir(job_id) / "output"
    path = directory / "exceptions.csv"
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Weekly exceptions not found.")
    return FileResponse(path, media_type="text/csv", filename=f"{stage}-exceptions.csv")


@router.post("/jobs/{job_id}/refresh-review", status_code=202)
def refresh_weekly_review(job_id: str) -> dict:
    with STATE_LOCK:
        state = _read_state(job_id)
        if state["status"] not in {"waiting_review", "review_failed"}:
            raise HTTPException(status_code=409, detail="Weekly job is not waiting for Catalyst review.")
        if (_preparation_dir(job_id, state) / "draft-master-data.csv").is_file():
            raise HTTPException(status_code=409, detail="This job already has the current Catalyst review format.")
        _update_state(job_id, status="refreshing_review", stage="Refreshing Catalyst review", error=None)
    JOB_EXECUTOR.submit(_run_refresh_review, job_id)
    return {"job_id": job_id, "status": "refreshing_review", "status_url": f"/api/weekly/jobs/{job_id}"}


@router.post("/jobs/{job_id}/finalize", status_code=202)
async def submit_weekly_review(
    job_id: str,
    catalyst_review: UploadFile | None = File(default=None),
    reviewed_statuses: str | None = Form(default=None),
    source_review: UploadFile | None = File(default=None),
    reviewed_sources: str | None = Form(default=None),
    manual_decisions: UploadFile | None = File(default=None),
    slide_review: UploadFile | None = File(default=None),
    pivot_review: UploadFile | None = File(default=None),
    presentation_review: UploadFile | None = File(default=None),
) -> dict:
    if (catalyst_review is None) == (reviewed_statuses is None):
        raise HTTPException(status_code=400, detail="Submit either Catalyst statuses or a reviewed file.")
    if catalyst_review is not None:
        _require_extension(catalyst_review, "catalyst_review", {".csv", ".xlsx"})
    if source_review is not None and reviewed_sources is not None:
        raise HTTPException(status_code=400, detail="Submit BQ source choices or a reviewed source CSV, not both.")
    if source_review is not None:
        _require_extension(source_review, "source_review", {".csv"})
    if manual_decisions is not None:
        _require_extension(manual_decisions, "manual_decisions", {".csv"})
    if slide_review is not None:
        _require_extension(slide_review, "slide_review", {".csv"})
    if pivot_review is not None:
        _require_extension(pivot_review, "pivot_review", {".csv"})
    if presentation_review is not None:
        _require_extension(presentation_review, "presentation_review", {".csv"})
    with STATE_LOCK:
        state = _read_state(job_id)
        if state["status"] not in {"waiting_review", "review_failed"}:
            raise HTTPException(status_code=409, detail="Weekly job is not waiting for Catalyst review.")
        inline_rows = _inline_review_rows(job_id, reviewed_statuses) if reviewed_statuses is not None else None
        inline_source_rows = _inline_source_rows(job_id, reviewed_sources) if reviewed_sources is not None else None
        _update_state(job_id, status="uploading_review", stage="Uploading Catalyst review")
    review_filename = (
        "catalyst-review-inline.csv" if inline_rows is not None
        else "catalyst-review-upload" + Path(catalyst_review.filename).suffix.lower()
    )
    source_filename = (
        "source-review-inline.csv" if inline_source_rows is not None
        else "source-review-upload.csv" if source_review else None
    )
    decisions_filename = "manual-decisions-upload.csv" if manual_decisions else None
    slide_filename = "slide-review-upload.csv" if slide_review else None
    pivot_filename = "pivot-review-upload.csv" if pivot_review else None
    presentation_filename = "presentation-review-upload.csv" if presentation_review else None
    try:
        if inline_rows is not None:
            columns, rows = inline_rows
            path = _job_dir(job_id) / "input" / review_filename
            with path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=columns)
                writer.writeheader()
                writer.writerows(rows)
        else:
            await _save_upload(catalyst_review, _job_dir(job_id) / "input" / review_filename)
        if inline_source_rows is not None:
            columns, rows = inline_source_rows
            path = _job_dir(job_id) / "input" / source_filename
            with path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=columns)
                writer.writeheader()
                writer.writerows(rows)
        elif source_review is not None:
            await _save_upload(source_review, _job_dir(job_id) / "input" / source_filename)
        if manual_decisions is not None:
            await _save_upload(manual_decisions, _job_dir(job_id) / "input" / decisions_filename)
        if slide_review is not None:
            await _save_upload(slide_review, _job_dir(job_id) / "input" / slide_filename)
        if pivot_review is not None:
            await _save_upload(pivot_review, _job_dir(job_id) / "input" / pivot_filename)
        if presentation_review is not None:
            await _save_upload(
                presentation_review, _job_dir(job_id) / "input" / presentation_filename
            )
    except Exception as exc:
        _update_state(job_id, status="waiting_review", stage="Waiting for Catalyst review", error=str(exc))
        raise
    _update_state(job_id, status="queued_finalization", stage="Queued for finalization", error=None)
    JOB_EXECUTOR.submit(
        _run_finalize, job_id, review_filename, decisions_filename, slide_filename,
        pivot_filename,
        presentation_filename,
        source_filename,
    )
    return {"job_id": job_id, "status": "queued_finalization", "status_url": f"/api/weekly/jobs/{job_id}"}


@router.get("/jobs/{job_id}/download")
def download_weekly_output(job_id: str) -> FileResponse:
    state = _read_state(job_id)
    if state["status"] != "succeeded":
        raise HTTPException(status_code=409, detail="Weekly output is not ready.")
    path = _job_dir(job_id) / "weekly-mobile-output.zip"
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Weekly output is missing.")
    files = [name for name in _allowed_report_files(state)
             if name.endswith(".xlsx") or name in state.get("output_files", [])]
    # Older completed jobs also expose only reports, without rerunning generation.
    with STATE_LOCK:
        with zipfile.ZipFile(path) as archive:
            current_files = archive.namelist()
        if sorted(current_files) != sorted(files):
            if any(not (_job_dir(job_id) / "output" / name).is_file() for name in files):
                raise HTTPException(status_code=404, detail="Weekly report files are missing.")
            _write_report_archive(_job_dir(job_id), files)
    return FileResponse(path, media_type="application/zip", filename=path.name)
