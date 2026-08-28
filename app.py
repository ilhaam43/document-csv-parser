from __future__ import annotations

import contextlib
import io
import logging
import os
import shutil
import subprocess
import threading
import time
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import quote
from urllib import request as url_request
from urllib.error import HTTPError, URLError

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask

from csv_to_excel import (
    ConvertOptions,
    convert_many,
    convert_one,
    output_filename_from_csv_path,
    resolve_csv_files,
    resolve_output_path,
)
from csv_to_excel_on_going import (
    output_filename_for_tracking as ongoing_output_filename_for_tracking,
)
from excel_automation_lock import excel_process_lock as _excel_process_lock
from generate_iphone_tracking import (
    default_output_path as iphone_default_output_path,
)
from generate_ide_tracking import (
    determine_report_date as determine_ide_report_date,
    output_filename as ide_output_filename,
)
from send_report_1_email import excel_range_to_png, excel_pivot_regions_to_png, send_email, send_email_images
from send_report_2_email import REPORT_2_IMAGE_IDS, report_2_ranges
from send_report_3_email import report_3_target_complete_range
from send_report_4_email import REPORT_4_IMAGE_IDS, report_4_ranges
from pipeline_stage_runner import run_daily_stage, run_ide_stage, run_iphone_stage, run_ongoing_stage


logger = logging.getLogger(__name__)
APP_ROOT = Path(__file__).resolve().parent
CONVERSION_OUTPUT_DIR = Path("output-today")
ONGOING_CONVERSION_OUTPUT_DIR = Path("output-outgoing")
IPHONE_CONVERSION_OUTPUT_DIR = Path("output-iphone")
IDE_CONVERSION_OUTPUT_DIR = Path("output-ide")
PIPELINE_CONVERSION_OUTPUT_DIR = Path("output-pipeline")
API_WORK_DIR = Path("output-today/api")
EMAIL_WORK_DIR = Path("output-today/email")
EMAIL_2_WORK_DIR = Path("output-outgoing/email")
EMAIL_3_WORK_DIR = Path("output-iphone/email")
EMAIL_4_WORK_DIR = Path("output-ide/email")
RUNTIME_DIRS = (
    Path("input-today"),
    CONVERSION_OUTPUT_DIR,
    ONGOING_CONVERSION_OUTPUT_DIR,
    IPHONE_CONVERSION_OUTPUT_DIR,
    IDE_CONVERSION_OUTPUT_DIR,
    PIPELINE_CONVERSION_OUTPUT_DIR,
    Path("vlookup-yesterday"),
    API_WORK_DIR,
    EMAIL_WORK_DIR,
    EMAIL_2_WORK_DIR,
    EMAIL_3_WORK_DIR,
    EMAIL_4_WORK_DIR,
)


def ensure_runtime_directories() -> None:
    for directory in RUNTIME_DIRS:
        (APP_ROOT / directory).mkdir(parents=True, exist_ok=True)


ensure_runtime_directories()

REPORT_1_JOBS: dict[str, dict[str, object]] = {}
REPORT_1_JOBS_LOCK = threading.Lock()
REPORT_1_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="report-1")
REPORT_2_JOBS: dict[str, dict[str, object]] = {}
REPORT_2_JOBS_LOCK = threading.Lock()
REPORT_2_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="report-2")
REPORT_3_JOBS: dict[str, dict[str, object]] = {}
REPORT_3_JOBS_LOCK = threading.Lock()
REPORT_3_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="report-3")
REPORT_4_JOBS: dict[str, dict[str, object]] = {}
REPORT_4_JOBS_LOCK = threading.Lock()
REPORT_4_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="report-4")
PIPELINE_JOBS: dict[str, dict[str, object]] = {}
PIPELINE_JOBS_LOCK = threading.Lock()
PIPELINE_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="pipeline")
EXCEL_COM_LOCK = threading.RLock()
REPORT_1_EXCEL_TIMEOUT_SECONDS = 25 * 60
REPORT_2_EXCEL_TIMEOUT_SECONDS = 25 * 60
REPORT_3_EXCEL_TIMEOUT_SECONDS = 25 * 60
REPORT_4_EXCEL_TIMEOUT_SECONDS = 25 * 60
PIPELINE_EXCEL_TIMEOUT_SECONDS = 60 * 60

app = FastAPI(
    title="Report CSV Parser API",
    version="1.0.0",
)
app.mount("/static", StaticFiles(directory=APP_ROOT / "static"), name="static")


@contextlib.contextmanager
def _initialized_com_thread():
    """Initialize Windows COM for worker threads that automate Excel."""
    with EXCEL_COM_LOCK:
        with _excel_process_lock(PIPELINE_EXCEL_TIMEOUT_SECONDS):
            try:
                import pythoncom  # type: ignore
            except ImportError:
                yield
                return

            pythoncom.CoInitialize()
            try:
                yield
            finally:
                pythoncom.CoUninitialize()


def _run_locked_excel_stage(stage_callable, timeout_seconds: int, *args):
    """Run an isolated Excel stage while honoring the server-wide queue."""
    with EXCEL_COM_LOCK:
        with _excel_process_lock(timeout_seconds):
            return stage_callable(*args)


class ConvertPathRequest(BaseModel):
    input_path: str = Field(default="input-today")
    output_path: str | None = None
    delimiter: str | None = None
    encoding: str | None = None
    normalize_headers: bool = True
    keep_empty: bool = False
    drop_empty_columns: bool = False
    dedupe: bool = False
    infer_types: bool = False
    combine: bool = False
    refresh_template: bool = False


def _resolve_path(value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return (APP_ROOT / path).resolve()


def _download_url(request: Request, job_id: str, filename: str) -> str:
    return f"/outputs/{job_id}/{quote(filename)}"


def _upload_image_to_onedrive(image_path: Path) -> dict[str, object]:
    access_token = os.getenv("ONEDRIVE_ACCESS_TOKEN")
    drive_id = os.getenv("ONEDRIVE_DRIVE_ID")
    if not access_token or not drive_id:
        raise RuntimeError("Set ONEDRIVE_ACCESS_TOKEN and ONEDRIVE_DRIVE_ID before sending an email.")

    folder = os.getenv("ONEDRIVE_FOLDER", "Daily Reports").strip("/")
    path_parts = [part for part in folder.split("/") if part] + [image_path.name]
    encoded_path = "/".join(quote(part, safe="") for part in path_parts)
    endpoint = f"https://graph.microsoft.com/v1.0/drives/{quote(drive_id, safe='')}/root:/{encoded_path}:/content"
    upload_request = url_request.Request(
        endpoint,
        data=image_path.read_bytes(),
        method="PUT",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "image/png",
        },
    )
    try:
        with url_request.urlopen(upload_request, timeout=60) as response:
            payload = response.read().decode("utf-8", errors="replace")
    except (HTTPError, URLError) as exc:
        detail = exc.read().decode("utf-8", errors="replace") if isinstance(exc, HTTPError) else str(exc)
        raise RuntimeError(f"OneDrive upload failed: {detail}") from exc

    import json

    result = json.loads(payload) if payload else {}
    return {
        "id": result.get("id"),
        "name": result.get("name", image_path.name),
        "web_url": result.get("webUrl"),
    }


def _run_email_report(
    workbook_path: Path,
    image_dir: Path,
    recipient: str,
    subject: str,
    message: str,
    dry_run: bool,
) -> dict[str, object]:
    started = time.perf_counter()
    images = excel_pivot_regions_to_png(workbook_path, image_dir)
    result: dict[str, object] = {
        "image_filenames": [image_path.name for image_path, _, _ in images],
        "ranges": [cell_range for _, _, cell_range in images],
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }
    if dry_run:
        result["status"] = "preview_created"
        result["message"] = "Screenshot created. OneDrive and SMTP were skipped because dry_run=true."
        return result

    result["onedrive"] = [_upload_image_to_onedrive(image_path) for image_path, _, _ in images]
    smtp_endpoint = os.getenv(
        "SMTP_API_URL",
        "http://10.34.144.197/secm-portal/smtp/api_send_email",
    )
    send_email_images(smtp_endpoint, recipient, subject, message, [(image_path, content_id) for image_path, content_id, _ in images], timeout=60)
    result["status"] = "sent"
    result["message"] = "Screenshot uploaded to OneDrive and sent through the SMTP API."
    result["elapsed_seconds"] = round(time.perf_counter() - started, 3)
    return result


def _run_email_report_2(workbook_path: Path, image_path: Path, recipient: str, subject: str, message: str, dry_run: bool) -> dict[str, object]:
    started = time.perf_counter()
    images = []
    for index, (cell_range, content_id) in enumerate(report_2_ranges(workbook_path), start=1):
        output_path = image_path.parent / f"{workbook_path.stem} - report-2-pivot-{index}.png"
        excel_range_to_png(workbook_path, output_path, "PIVOT", cell_range, scale=3.0 if "legend" in content_id else 1.0)
        images.append((output_path, content_id, cell_range))
    result: dict[str, object] = {"image_filenames": [path.name for path, _, _ in images], "ranges": [cell_range for _, _, cell_range in images], "elapsed_seconds": round(time.perf_counter() - started, 3)}
    if dry_run:
        result.update(status="preview_created", message="Report 2 screenshot created. OneDrive and SMTP were skipped because dry_run=true.")
        return result
    result["onedrive"] = [_upload_image_to_onedrive(path) for path, _, _ in images]
    send_email_images(os.getenv("SMTP_API_URL", "http://10.34.144.197/secm-portal/smtp/api_send_email"), recipient, subject, message, [(path, content_id) for path, content_id, _ in images], timeout=60)
    result.update(status="sent", message="Report 2 screenshot uploaded to OneDrive and sent through the SMTP API.")
    return result


def _run_email_report_3(workbook_path: Path, image_path: Path, recipient: str, subject: str, message: str, dry_run: bool) -> dict[str, object]:
    started = time.perf_counter()
    cell_range = report_3_target_complete_range(workbook_path)
    excel_range_to_png(workbook_path, image_path, "PIVOT", cell_range)
    result: dict[str, object] = {"image_filename": image_path.name, "range": cell_range, "elapsed_seconds": round(time.perf_counter() - started, 3)}
    if dry_run:
        result.update(status="preview_created", message="Report 3 screenshot created. OneDrive and SMTP were skipped because dry_run=true.")
        return result
    result["onedrive"] = _upload_image_to_onedrive(image_path)
    send_email(os.getenv("SMTP_API_URL", "http://10.34.144.197/secm-portal/smtp/api_send_email"), recipient, subject, message, image_path, timeout=60, content_id="report-3-pivot")
    result.update(status="sent", message="Report 3 screenshot uploaded to OneDrive and sent through the SMTP API.")
    return result


def _run_email_report_4(workbook_path: Path, image_path: Path, recipient: str, subject: str, message: str, dry_run: bool) -> dict[str, object]:
    started = time.perf_counter()
    images = []
    for index, (cell_range, content_id) in enumerate(report_4_ranges(workbook_path), start=1):
        output_path = image_path.parent / f"{workbook_path.stem} - report-4-pivot-{index}.png"
        excel_range_to_png(workbook_path, output_path, "PIVOT", cell_range, scale=3.0 if "legend" in content_id else 1.0)
        images.append((output_path, content_id, cell_range))
    result: dict[str, object] = {"image_filenames": [path.name for path, _, _ in images], "ranges": [cell_range for _, _, cell_range in images], "elapsed_seconds": round(time.perf_counter() - started, 3)}
    if dry_run:
        result.update(status="preview_created", message="Report 4 screenshot created. OneDrive and SMTP were skipped because dry_run=true.")
        return result
    result["onedrive"] = [_upload_image_to_onedrive(path) for path, _, _ in images]
    send_email_images(os.getenv("SMTP_API_URL", "http://10.34.144.197/secm-portal/smtp/api_send_email"), recipient, subject, message, [(path, content_id) for path, content_id, _ in images], timeout=60)
    result.update(status="sent", message="Report 4 screenshot uploaded to OneDrive and sent through the SMTP API.")
    return result


async def _save_upload_file(upload: UploadFile, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as handle:
        while chunk := await upload.read(1024 * 1024):
            handle.write(chunk)


def _job_input_path(job_dir: Path, input_name: str, filename: str) -> Path:
    """Keep uploaded roles isolated while preserving source filenames."""
    return (job_dir / "input" / input_name / Path(filename).name).resolve()


def _now_seconds() -> float:
    return round(time.time(), 3)


def _set_report_1_job(job_id: str, **updates: object) -> None:
    with REPORT_1_JOBS_LOCK:
        job = REPORT_1_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job["updated_at"] = _now_seconds()


def _get_report_1_job(job_id: str) -> dict[str, object] | None:
    with REPORT_1_JOBS_LOCK:
        job = REPORT_1_JOBS.get(job_id)
        return dict(job) if job else None


def _set_report_2_job(job_id: str, **updates: object) -> None:
    with REPORT_2_JOBS_LOCK:
        job = REPORT_2_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job["updated_at"] = _now_seconds()


def _get_report_2_job(job_id: str) -> dict[str, object] | None:
    with REPORT_2_JOBS_LOCK:
        job = REPORT_2_JOBS.get(job_id)
        return dict(job) if job else None


def _set_report_3_job(job_id: str, **updates: object) -> None:
    with REPORT_3_JOBS_LOCK:
        job = REPORT_3_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job["updated_at"] = _now_seconds()


def _get_report_3_job(job_id: str) -> dict[str, object] | None:
    with REPORT_3_JOBS_LOCK:
        job = REPORT_3_JOBS.get(job_id)
        return dict(job) if job else None


def _set_report_4_job(job_id: str, **updates: object) -> None:
    with REPORT_4_JOBS_LOCK:
        job = REPORT_4_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job["updated_at"] = _now_seconds()


def _get_report_4_job(job_id: str) -> dict[str, object] | None:
    with REPORT_4_JOBS_LOCK:
        job = REPORT_4_JOBS.get(job_id)
        return dict(job) if job else None


def _set_pipeline_job(job_id: str, **updates: object) -> None:
    with PIPELINE_JOBS_LOCK:
        job = PIPELINE_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job["updated_at"] = _now_seconds()


def _get_pipeline_job(job_id: str) -> dict[str, object] | None:
    with PIPELINE_JOBS_LOCK:
        job = PIPELINE_JOBS.get(job_id)
        return dict(job) if job else None


def _mark_job_downloaded(job_id: str) -> None:
    updates = {
        "downloaded": True,
        "downloaded_at": _now_seconds(),
    }
    _set_report_1_job(job_id, **updates)
    _set_report_2_job(job_id, **updates)
    _set_report_3_job(job_id, **updates)
    _set_report_4_job(job_id, **updates)
    _set_pipeline_job(job_id, **updates)


def _terminate_excel_processes(reason: str) -> None:
    try:
        result = subprocess.run(
            ["taskkill.exe", "/F", "/IM", "EXCEL.EXE"],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode == 0:
            logger.warning("Stopped Excel processes %s.", reason)
    except Exception:
        logger.exception("Failed to stop Excel processes %s.", reason)


def _cleanup_excel_lock_markers(directory: Path) -> None:
    try:
        for lock_file in directory.rglob("~$*.xlsx"):
            try:
                lock_file.unlink()
            except FileNotFoundError:
                pass
            except Exception:
                logger.warning("Could not remove Excel lock marker: %s", lock_file, exc_info=True)
    except Exception:
        logger.warning("Could not scan Excel lock markers in %s", directory, exc_info=True)


def _delete_job_dir_after_download(job_id: str, job_dir: Path) -> None:
    api_root = (APP_ROOT / API_WORK_DIR).resolve()
    resolved_job_dir = job_dir.resolve()
    if resolved_job_dir.parent != api_root:
        logger.error("Refusing to delete unexpected download directory: %s", resolved_job_dir)
        return

    _mark_job_downloaded(job_id)
    for attempt in range(1, 9):
        try:
            _cleanup_excel_lock_markers(resolved_job_dir)
            shutil.rmtree(resolved_job_dir)
            logger.info("Deleted job directory after download: %s", resolved_job_dir)
            return
        except FileNotFoundError:
            return
        except Exception:
            if attempt == 8:
                logger.exception("Could not delete job directory after download: %s", resolved_job_dir)
                return
            time.sleep(1.5)


def _run_with_excel_watchdog(
    job_label: str,
    job_id: str,
    set_job,
    cleanup_dir: Path,
    timeout_seconds: int,
    work,
) -> dict[str, object] | None:
    timed_out = threading.Event()
    failed = False

    def expire_job() -> None:
        timed_out.set()
        error = (
            f"Excel processing exceeded {timeout_seconds // 60} minutes. "
            "The workbook may be too large or Excel may be waiting on a hidden dialog."
        )
        logger.error("%s background conversion timed out for job %s", job_label, job_id)
        set_job(
            job_id,
            status="failed",
            completed_at=_now_seconds(),
            error=error,
        )
        _terminate_excel_processes(f"after {job_label} job {job_id} timed out")

    try:
        with EXCEL_COM_LOCK:
            with _excel_process_lock(timeout_seconds):
                timeout_timer = threading.Timer(timeout_seconds, expire_job)
                timeout_timer.daemon = True
                timeout_timer.start()
                try:
                    result = work()
                    if timed_out.is_set():
                        return None
                    return result
                except Exception as exc:
                    failed = True
                    logger.exception("%s background conversion failed for job %s", job_label, job_id)
                    if not timed_out.is_set():
                        set_job(
                            job_id,
                            status="failed",
                            completed_at=_now_seconds(),
                            error=str(exc),
                        )
                    return None
                finally:
                    timeout_timer.cancel()
                    if timed_out.is_set() or failed:
                        _terminate_excel_processes(f"after failed {job_label} job {job_id}")
                    _cleanup_excel_lock_markers(cleanup_dir)
    except Exception as exc:
        logger.exception("%s could not acquire the Excel automation lock for job %s", job_label, job_id)
        set_job(
            job_id,
            status="failed",
            completed_at=_now_seconds(),
            error=str(exc),
        )
        return None


def _run_conversion(
    input_path: Path,
    output_path: Path | None,
    request: ConvertPathRequest,
    lookup_workbook: Path | None = None,
) -> dict[str, object]:
    csv_files = resolve_csv_files(input_path)
    args = SimpleNamespace(output=output_path, combine=request.combine)
    resolved_output = resolve_output_path(input_path, csv_files, args).resolve()
    options = ConvertOptions(
        output=resolved_output,
        delimiter=request.delimiter,
        encoding=request.encoding,
        normalize_headers=request.normalize_headers,
        keep_empty=request.keep_empty,
        drop_empty_columns=request.drop_empty_columns,
        dedupe=request.dedupe,
        infer_types=request.infer_types,
        combine=request.combine,
        refresh_template=request.refresh_template,
        lookup_workbook=lookup_workbook,
    )

    stdout = io.StringIO()
    stderr = io.StringIO()
    started = time.perf_counter()
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        with _initialized_com_thread():
            if len(csv_files) == 1 and input_path.is_file():
                convert_one(csv_files[0], resolved_output, options)
                output_files = [resolved_output]
            else:
                convert_many(csv_files, options)
                if request.combine:
                    output_files = [resolved_output]
                else:
                    output_files = [resolved_output / output_filename_from_csv_path(csv_path) for csv_path in csv_files]
    elapsed_seconds = time.perf_counter() - started

    return {
        "elapsed_seconds": round(elapsed_seconds, 3),
        "input_files": [str(path) for path in csv_files],
        "output_path": str(resolved_output),
        "output_files": [str(path) for path in output_files],
        "stdout": stdout.getvalue().strip(),
        "stderr": stderr.getvalue().strip(),
        "refresh_template": request.refresh_template,
    }


def _run_report_1_upload_job(
    job_id: str,
    input_path: Path,
    lookup_workbook_path: Path,
    output_path: Path,
    conversion_request: ConvertPathRequest,
) -> None:
    _set_report_1_job(job_id, status="running", started_at=_now_seconds())

    def work() -> dict[str, object]:
        result = run_daily_stage(
            input_path,
            lookup_workbook_path,
            output_path,
            conversion_request.refresh_template,
            delimiter=conversion_request.delimiter,
            encoding=conversion_request.encoding,
            normalize_headers=conversion_request.normalize_headers,
            keep_empty=conversion_request.keep_empty,
            drop_empty_columns=conversion_request.drop_empty_columns,
            dedupe=conversion_request.dedupe,
            infer_types=conversion_request.infer_types,
            timeout_seconds=REPORT_1_EXCEL_TIMEOUT_SECONDS,
        )
        return {
            "elapsed_seconds": round(result.elapsed_seconds, 3),
            "stdout": result.stdout,
            "stderr": result.stderr,
        }

    result = _run_with_excel_watchdog(
        "Report 1",
        job_id,
        _set_report_1_job,
        output_path.parent,
        REPORT_1_EXCEL_TIMEOUT_SECONDS,
        work,
    )
    if result is None:
        return
    _set_report_1_job(
        job_id,
        status="succeeded",
        completed_at=_now_seconds(),
        elapsed_seconds=result["elapsed_seconds"],
        filename=output_path.name,
        output_file=str(output_path),
        refresh_template=conversion_request.refresh_template,
        stdout=result.get("stdout", ""),
        stderr=result.get("stderr", ""),
    )


def _run_report_2_upload_job(
    job_id: str,
    tracking_path: Path,
    log_path: Path,
    previous_ongoing_path: Path,
    output_path: Path,
    with_pivot: bool,
) -> None:
    _set_report_2_job(job_id, status="running", started_at=_now_seconds())

    def work() -> dict[str, object]:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        result = run_ongoing_stage(
            tracking_path,
            log_path,
            previous_ongoing_path,
            output_path,
            with_pivot,
            timeout_seconds=REPORT_2_EXCEL_TIMEOUT_SECONDS,
        )
        return {
            "elapsed_seconds": round(result.elapsed_seconds, 3),
            "stdout": result.stdout,
            "stderr": result.stderr,
        }

    result = _run_with_excel_watchdog(
        "Report 2",
        job_id,
        _set_report_2_job,
        output_path.parent,
        REPORT_2_EXCEL_TIMEOUT_SECONDS,
        work,
    )
    if result is None:
        return
    _set_report_2_job(
        job_id,
        status="succeeded",
        completed_at=_now_seconds(),
        elapsed_seconds=result["elapsed_seconds"],
        filename=output_path.name,
        output_file=str(output_path),
        with_pivot=with_pivot,
        stdout=result.get("stdout", ""),
        stderr=result.get("stderr", ""),
    )


def _run_pipeline_conversion(
    data_order_csv_path: Path,
    daily_reference_workbook_path: Path,
    log_csv_path: Path,
    previous_ongoing_workbook_path: Path,
    previous_iphone_workbook_path: Path,
    output_dir: Path,
    refresh_template: bool,
    ongoing_with_pivot: bool,
) -> dict[str, object]:
    started = time.perf_counter()
    output_dir.mkdir(parents=True, exist_ok=True)

    daily_output_path = (output_dir / output_filename_from_csv_path(data_order_csv_path)).resolve()
    ongoing_output_path = (output_dir / ongoing_output_filename_for_tracking(daily_output_path)).resolve()
    iphone_output_path = (output_dir / iphone_default_output_path(daily_output_path).name).resolve()

    stdout_parts: list[str] = []
    stderr_parts: list[str] = []

    def record_stage(stage_name: str, stage_result) -> None:
        if stage_result.stdout:
            stdout_parts.append(f"[{stage_name}]\n{stage_result.stdout}")
        if stage_result.stderr:
            stderr_parts.append(f"[{stage_name}]\n{stage_result.stderr}")

    # Each generator runs in its own process. Process exit guarantees that its
    # Excel COM apartment and workbook handles are released before the next
    # generator opens the output.
    with EXCEL_COM_LOCK:
        with _excel_process_lock(PIPELINE_EXCEL_TIMEOUT_SECONDS):
            record_stage(
                "Daily Tracking",
                run_daily_stage(
                    data_order_csv_path,
                    daily_reference_workbook_path,
                    daily_output_path,
                    refresh_template,
                    timeout_seconds=REPORT_1_EXCEL_TIMEOUT_SECONDS,
                ),
            )
            record_stage(
                "Ongoing Tracking",
                run_ongoing_stage(
                    daily_output_path,
                    log_csv_path,
                    previous_ongoing_workbook_path,
                    ongoing_output_path,
                    ongoing_with_pivot,
                    timeout_seconds=REPORT_2_EXCEL_TIMEOUT_SECONDS,
                ),
            )
            record_stage(
                "iPhone Tracking",
                run_iphone_stage(
                    daily_output_path,
                    previous_iphone_workbook_path,
                    iphone_output_path,
                    REPORT_3_EXCEL_TIMEOUT_SECONDS,
                ),
            )
    elapsed_seconds = time.perf_counter() - started

    return {
        "elapsed_seconds": round(elapsed_seconds, 3),
        "output_files": [str(daily_output_path), str(ongoing_output_path), str(iphone_output_path)],
        "stdout": "\n".join(stdout_parts),
        "stderr": "\n".join(stderr_parts),
        "refresh_template": refresh_template,
        "ongoing_with_pivot": ongoing_with_pivot,
    }


def _run_report_3_upload_job(
    job_id: str,
    tracking_path: Path,
    reference_path: Path,
    output_path: Path,
) -> None:
    _set_report_3_job(job_id, status="running", started_at=_now_seconds())

    def work() -> dict[str, object]:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        result = run_iphone_stage(
            tracking_path,
            reference_path,
            output_path,
            REPORT_3_EXCEL_TIMEOUT_SECONDS,
        )
        return {
            "elapsed_seconds": round(result.elapsed_seconds, 3),
            "stdout": result.stdout,
            "stderr": result.stderr,
        }

    result = _run_with_excel_watchdog(
        "Report 3",
        job_id,
        _set_report_3_job,
        output_path.parent,
        REPORT_3_EXCEL_TIMEOUT_SECONDS,
        work,
    )
    if result is None:
        return
    _set_report_3_job(
        job_id,
        status="succeeded",
        completed_at=_now_seconds(),
        elapsed_seconds=result["elapsed_seconds"],
        filename=output_path.name,
        output_file=str(output_path),
        stdout=result.get("stdout", ""),
        stderr=result.get("stderr", ""),
    )


def _run_report_4_upload_job(
    job_id: str,
    raw_workbook_path: Path,
    previous_workbook_path: Path,
    collabs_csv_path: Path,
    output_path: Path,
    report_date: date,
) -> None:
    _set_report_4_job(job_id, status="running", started_at=_now_seconds())

    def work() -> dict[str, object]:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        stage_result = run_ide_stage(
            raw_workbook_path,
            previous_workbook_path,
            collabs_csv_path,
            output_path,
            report_date,
            REPORT_4_EXCEL_TIMEOUT_SECONDS,
        )
        return {
            **(stage_result.metadata or {}),
            "elapsed_seconds": round(stage_result.elapsed_seconds, 3),
            "report_date": report_date.isoformat(),
            "stdout": stage_result.stdout,
            "stderr": stage_result.stderr,
        }

    result = _run_with_excel_watchdog(
        "Report 4",
        job_id,
        _set_report_4_job,
        output_path.parent,
        REPORT_4_EXCEL_TIMEOUT_SECONDS,
        work,
    )
    if result is None:
        return
    _set_report_4_job(
        job_id,
        status="succeeded",
        completed_at=_now_seconds(),
        elapsed_seconds=result["elapsed_seconds"],
        filename=output_path.name,
        output_file=str(output_path),
        report_date=result["report_date"],
        rows=result["rows"],
        columns=result["columns"],
        pivot_tables=result["pivot_tables"],
        collabs_fallback_rows=result["collabs_fallback_rows"],
        zero_fallback_rows=result["zero_fallback_rows"],
        stdout=result.get("stdout", ""),
        stderr=result.get("stderr", ""),
    )


def _run_pipeline_upload_job(
    job_id: str,
    raw_path: Path,
    daily_reference_path: Path,
    log_path: Path,
    previous_ongoing_path: Path,
    previous_iphone_path: Path,
    output_dir: Path,
    pipeline_zip_path: Path,
    refresh_template: bool,
    ongoing_with_pivot: bool,
) -> None:
    _set_pipeline_job(job_id, status="running", started_at=_now_seconds())

    def work() -> dict[str, object]:
        result = _run_pipeline_conversion(
            raw_path,
            daily_reference_path,
            log_path,
            previous_ongoing_path,
            previous_iphone_path,
            output_dir,
            refresh_template,
            ongoing_with_pivot,
        )
        output_files = [Path(path) for path in result["output_files"]]
        with zipfile.ZipFile(pipeline_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for output_file in output_files:
                archive.write(output_file, arcname=output_file.name)
        return result

    result = _run_with_excel_watchdog(
        "Daily Pipeline",
        job_id,
        _set_pipeline_job,
        pipeline_zip_path.parent,
        PIPELINE_EXCEL_TIMEOUT_SECONDS,
        work,
    )
    if result is None:
        return
    _set_pipeline_job(
        job_id,
        status="succeeded",
        completed_at=_now_seconds(),
        elapsed_seconds=result["elapsed_seconds"],
        filename=pipeline_zip_path.name,
        output_files=[Path(path).name for path in result["output_files"]],
        output_file=str(pipeline_zip_path),
        refresh_template=refresh_template,
        ongoing_with_pivot=ongoing_with_pivot,
        stdout=result.get("stdout", ""),
        stderr=result.get("stderr", ""),
    )


@app.get("/")
def home() -> RedirectResponse:
    return RedirectResponse(url="/report-1", status_code=302)


@app.get("/report-1")
def upload_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "index.html", media_type="text/html")


@app.get("/report-2")
def report_2_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "report-2.html", media_type="text/html")


@app.get("/report-3")
def report_3_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "report-3.html", media_type="text/html")


@app.get("/report-4")
def report_4_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "report-4.html", media_type="text/html")


@app.get("/pipeline")
def pipeline_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "pipeline.html", media_type="text/html")


@app.get("/email-report-1")
def email_report_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "email-report.html", media_type="text/html")


@app.get("/email-report")
def legacy_email_report_page() -> RedirectResponse:
    return RedirectResponse(url="/email-report-1", status_code=307)


@app.get("/email-report-2")
def email_report_2_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "email-report-2.html", media_type="text/html")


@app.get("/email-report-3")
def email_report_3_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "email-report-3.html", media_type="text/html")


@app.get("/email-report-4")
def email_report_4_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "email-report-4.html", media_type="text/html")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/email-report-1/send")
@app.post("/email-report/send")
async def send_email_report(
    workbook: UploadFile = File(...),
    recipient: str = Form(...),
    subject: str = Form(default="Daily Tracking Report 1"),
    message: str = Form(default="<p>Daily Tracking Report 1</p><p><img src=\"cid:target-complete-table\" alt=\"Target complete table\"></p><p><img src=\"cid:target-complete-legend\" alt=\"Target complete legend\"></p><p><img src=\"cid:target-after-table\" alt=\"Target after table\"></p><p><img src=\"cid:target-after-legend\" alt=\"Target after legend\"></p>"),
    dry_run: bool = Form(default=True),
) -> dict[str, object]:
    if not workbook.filename or not workbook.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="workbook must be an .xlsx or .xlsm file.")
    if not recipient.strip():
        raise HTTPException(status_code=400, detail="recipient is required.")
    required_content_ids = ("target-complete-table", "target-complete-legend", "target-after-table", "target-after-legend")
    if any(f"cid:{content_id}" not in message for content_id in required_content_ids):
        raise HTTPException(status_code=400, detail="message must contain all four target image Content-IDs for inline Outlook rendering.")

    job_dir = (APP_ROOT / EMAIL_WORK_DIR / uuid.uuid4().hex).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = job_dir / Path(workbook.filename).name
    image_dir = job_dir / "images"
    try:
        await _save_upload_file(workbook, workbook_path)
        result = await run_in_threadpool(
            _run_email_report,
            workbook_path,
            image_dir,
            recipient.strip(),
            subject.strip(),
            message,
            dry_run,
        )
        result["image_dir"] = str(image_dir)
        return result
    except Exception as exc:
        logger.exception("Email report failed for %s", workbook_path)
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/email-report-2/send")
async def send_email_report_2(
    workbook: UploadFile = File(...),
    recipient: str = Form(...),
    subject: str = Form(default="Daily Tracking Report 2 Ongoing"),
    message: str = Form(default='<p>Daily Tracking Report 2 Ongoing</p><p><img src="cid:report-2-ongoing-table" alt="Report 2 ongoing table"></p><p><img src="cid:report-2-ongoing-legend" alt="Report 2 ongoing legend"></p><p><img src="cid:report-2-after-table" alt="Report 2 target after table"></p><p><img src="cid:report-2-after-legend" alt="Report 2 target after legend"></p><p><img src="cid:report-2-aging-table" alt="Report 2 order aging table"></p>'),
    dry_run: bool = Form(default=True),
) -> dict[str, object]:
    if not workbook.filename or not workbook.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="workbook must be an .xlsx or .xlsm file.")
    if not recipient.strip():
        raise HTTPException(status_code=400, detail="recipient is required.")
    if any(f"cid:{content_id}" not in message for content_id in REPORT_2_IMAGE_IDS):
        raise HTTPException(status_code=400, detail="message must contain all five Report 2 image Content-IDs for inline Outlook rendering.")

    job_dir = (APP_ROOT / EMAIL_2_WORK_DIR / uuid.uuid4().hex).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = job_dir / Path(workbook.filename).name
    image_path = job_dir / f"{workbook_path.stem} - report-2-pivot.png"
    try:
        await _save_upload_file(workbook, workbook_path)
        result = await run_in_threadpool(_run_email_report_2, workbook_path, image_path, recipient.strip(), subject.strip(), message, dry_run)
        result["image_path"] = str(image_path)
        return result
    except Exception as exc:
        logger.exception("Report 2 email failed for %s", workbook_path)
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/email-report-3/send")
async def send_email_report_3(
    workbook: UploadFile = File(...),
    recipient: str = Form(...),
    subject: str = Form(default="Daily Tracking Report 3 iPhone"),
    message: str = Form(default='<p>Daily Tracking Report 3 iPhone</p><p><img src="cid:report-3-pivot" alt="Report 3 iPhone PIVOT"></p>'),
    dry_run: bool = Form(default=True),
) -> dict[str, object]:
    if not workbook.filename or not workbook.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="workbook must be an .xlsx or .xlsm file.")
    if not recipient.strip():
        raise HTTPException(status_code=400, detail="recipient is required.")
    if "cid:report-3-pivot" not in message:
        raise HTTPException(status_code=400, detail="message must contain cid:report-3-pivot for inline Outlook rendering.")
    job_dir = (APP_ROOT / EMAIL_3_WORK_DIR / uuid.uuid4().hex).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = job_dir / Path(workbook.filename).name
    image_path = job_dir / f"{workbook_path.stem} - report-3-pivot.png"
    try:
        await _save_upload_file(workbook, workbook_path)
        result = await run_in_threadpool(_run_email_report_3, workbook_path, image_path, recipient.strip(), subject.strip(), message, dry_run)
        result["image_path"] = str(image_path)
        return result
    except Exception as exc:
        logger.exception("Report 3 email failed for %s", workbook_path)
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/email-report-4/send")
async def send_email_report_4(
    workbook: UploadFile = File(...),
    recipient: str = Form(...),
    subject: str = Form(default="Daily Tracking Report 4 IDE"),
    message: str = Form(default='<p>Daily Tracking Report 4 IDE</p><p><img src="cid:report-4-complete-table" alt="Report 4 Target Complete table"></p><p><img src="cid:report-4-complete-legend" alt="Report 4 Target Complete legend"></p><p><img src="cid:report-4-after-table" alt="Report 4 Target After table"></p><p><img src="cid:report-4-after-legend" alt="Report 4 Target After legend"></p>'),
    dry_run: bool = Form(default=True),
) -> dict[str, object]:
    if not workbook.filename or not workbook.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="workbook must be an .xlsx or .xlsm file.")
    if not recipient.strip():
        raise HTTPException(status_code=400, detail="recipient is required.")
    if any(f"cid:{content_id}" not in message for content_id in REPORT_4_IMAGE_IDS):
        raise HTTPException(status_code=400, detail="message must contain all four Report 4 image Content-IDs for inline Outlook rendering.")
    job_dir = (APP_ROOT / EMAIL_4_WORK_DIR / uuid.uuid4().hex).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = job_dir / Path(workbook.filename).name
    image_path = job_dir / f"{workbook_path.stem} - report-4-pivot.png"
    try:
        await _save_upload_file(workbook, workbook_path)
        result = await run_in_threadpool(_run_email_report_4, workbook_path, image_path, recipient.strip(), subject.strip(), message, dry_run)
        result["image_path"] = str(image_path)
        return result
    except Exception as exc:
        logger.exception("Report 4 email failed for %s", workbook_path)
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/jobs/report-1/{job_id}")
def report_1_job_status(request: Request, job_id: str) -> dict[str, object]:
    try:
        uuid.UUID(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="Job not found.") from exc

    job = _get_report_1_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")

    response = {
        "job_id": job_id,
        "status": job["status"],
        "created_at": job["created_at"],
        "updated_at": job["updated_at"],
        "raw_data_filename": job.get("raw_data_filename"),
        "yesterday_cleaned_data_filename": job.get("yesterday_cleaned_data_filename"),
        "filename": job.get("filename"),
        "elapsed_seconds": job.get("elapsed_seconds"),
        "refresh_template": job.get("refresh_template"),
        "error": job.get("error"),
        "downloaded": job.get("downloaded", False),
        "downloaded_at": job.get("downloaded_at"),
    }
    if job.get("status") == "succeeded" and job.get("filename") and not job.get("downloaded"):
        response["success"] = True
        response["message"] = "CSV converted successfully."
        response["download_url"] = _download_url(request, job_id, str(job["filename"]))
    return response


@app.get("/jobs/report-2/{job_id}")
def report_2_job_status(request: Request, job_id: str) -> dict[str, object]:
    try:
        uuid.UUID(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="Job not found.") from exc

    job = _get_report_2_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")

    response = {
        "job_id": job_id,
        "status": job["status"],
        "created_at": job["created_at"],
        "updated_at": job["updated_at"],
        "tracking_workbook_filename": job.get("tracking_workbook_filename"),
        "log_update_status_filename": job.get("log_update_status_filename"),
        "previous_ongoing_workbook_filename": job.get("previous_ongoing_workbook_filename"),
        "filename": job.get("filename"),
        "elapsed_seconds": job.get("elapsed_seconds"),
        "with_pivot": job.get("with_pivot"),
        "error": job.get("error"),
        "downloaded": job.get("downloaded", False),
        "downloaded_at": job.get("downloaded_at"),
    }
    if job.get("status") == "succeeded" and job.get("filename") and not job.get("downloaded"):
        response["success"] = True
        response["message"] = "Report 2 workbook generated successfully."
        response["download_url"] = _download_url(request, job_id, str(job["filename"]))
    return response


@app.get("/jobs/report-3/{job_id}")
def report_3_job_status(request: Request, job_id: str) -> dict[str, object]:
    try:
        uuid.UUID(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="Job not found.") from exc

    job = _get_report_3_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")

    response = {
        "job_id": job_id,
        "status": job["status"],
        "created_at": job["created_at"],
        "updated_at": job["updated_at"],
        "tracking_workbook_filename": job.get("tracking_workbook_filename"),
        "previous_iphone_workbook_filename": job.get("previous_iphone_workbook_filename"),
        "filename": job.get("filename"),
        "elapsed_seconds": job.get("elapsed_seconds"),
        "error": job.get("error"),
        "downloaded": job.get("downloaded", False),
        "downloaded_at": job.get("downloaded_at"),
    }
    if job.get("status") == "succeeded" and job.get("filename") and not job.get("downloaded"):
        response["success"] = True
        response["message"] = "Report 3 iPhone workbook generated successfully."
        response["download_url"] = _download_url(request, job_id, str(job["filename"]))
    return response


@app.get("/jobs/report-4/{job_id}")
def report_4_job_status(request: Request, job_id: str) -> dict[str, object]:
    try:
        uuid.UUID(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="Job not found.") from exc

    job = _get_report_4_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")

    response = {
        "job_id": job_id,
        "status": job["status"],
        "created_at": job["created_at"],
        "updated_at": job["updated_at"],
        "raw_ide_workbook_filename": job.get("raw_ide_workbook_filename"),
        "previous_ide_workbook_filename": job.get("previous_ide_workbook_filename"),
        "collabs_csv_filename": job.get("collabs_csv_filename"),
        "filename": job.get("filename"),
        "report_date": job.get("report_date"),
        "rows": job.get("rows"),
        "columns": job.get("columns"),
        "pivot_tables": job.get("pivot_tables"),
        "collabs_fallback_rows": job.get("collabs_fallback_rows"),
        "zero_fallback_rows": job.get("zero_fallback_rows"),
        "elapsed_seconds": job.get("elapsed_seconds"),
        "error": job.get("error"),
        "downloaded": job.get("downloaded", False),
        "downloaded_at": job.get("downloaded_at"),
    }
    if job.get("status") == "succeeded" and job.get("filename") and not job.get("downloaded"):
        response["success"] = True
        response["message"] = "Report 4 IDE workbook generated successfully."
        response["download_url"] = _download_url(request, job_id, str(job["filename"]))
    return response


@app.get("/jobs/pipeline/{job_id}")
def pipeline_job_status(request: Request, job_id: str) -> dict[str, object]:
    try:
        uuid.UUID(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="Job not found.") from exc

    job = _get_pipeline_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")

    response = {
        "job_id": job_id,
        "status": job["status"],
        "created_at": job["created_at"],
        "updated_at": job["updated_at"],
        "raw_data_filename": job.get("raw_data_filename"),
        "yesterday_cleaned_data_filename": job.get("yesterday_cleaned_data_filename"),
        "log_update_status_filename": job.get("log_update_status_filename"),
        "previous_ongoing_workbook_filename": job.get("previous_ongoing_workbook_filename"),
        "previous_iphone_workbook_filename": job.get("previous_iphone_workbook_filename"),
        "filename": job.get("filename"),
        "output_files": job.get("output_files"),
        "elapsed_seconds": job.get("elapsed_seconds"),
        "refresh_template": job.get("refresh_template"),
        "ongoing_with_pivot": job.get("ongoing_with_pivot"),
        "error": job.get("error"),
        "downloaded": job.get("downloaded", False),
        "downloaded_at": job.get("downloaded_at"),
    }
    if job.get("status") == "succeeded" and job.get("filename") and not job.get("downloaded"):
        response["success"] = True
        response["message"] = "Daily pipeline generated successfully."
        response["download_url"] = _download_url(request, job_id, str(job["filename"]))
    return response


@app.get("/outputs/{job_id}/{filename}")
def download_output(job_id: str, filename: str) -> FileResponse:
    try:
        uuid.UUID(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="Output file not found.") from exc

    safe_filename = Path(filename).name
    if safe_filename != filename:
        raise HTTPException(status_code=404, detail="Output file not found.")

    job_dir = (APP_ROOT / API_WORK_DIR / job_id).resolve()
    output_path = (job_dir / safe_filename).resolve()
    if output_path.parent != job_dir or not output_path.is_file():
        raise HTTPException(status_code=404, detail="Output file not found.")

    media_type = (
        "application/zip"
        if output_path.suffix.lower() == ".zip"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return FileResponse(
        output_path,
        media_type=media_type,
        filename=output_path.name,
        background=BackgroundTask(_delete_job_dir_after_download, job_id, job_dir),
    )


@app.post("/convert/path")
async def convert_path(request: ConvertPathRequest) -> dict[str, object]:
    input_path = _resolve_path(request.input_path)
    output_path = _resolve_path(request.output_path) if request.output_path else None
    try:
        return await run_in_threadpool(_run_conversion, input_path, output_path, request)
    except Exception as exc:
        logger.exception("Path conversion failed for input %s", input_path)
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/convert/upload/jobs")
async def start_report_1_upload_job(
    request: Request,
    raw_data: UploadFile = File(...),
    yesterday_cleaned_data: UploadFile = File(...),
    delimiter: str | None = Form(default=None),
    encoding: str | None = Form(default=None),
    normalize_headers: bool = Form(default=True),
    keep_empty: bool = Form(default=False),
    drop_empty_columns: bool = Form(default=False),
    dedupe: bool = Form(default=False),
    infer_types: bool = Form(default=False),
    refresh_template: bool = Form(default=False),
) -> dict[str, object]:
    if not raw_data.filename or not raw_data.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="raw_data must be a .csv file.")
    if not yesterday_cleaned_data.filename or not yesterday_cleaned_data.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="yesterday_cleaned_data must be a .xlsx file.")

    job_id = uuid.uuid4().hex
    job_dir = (APP_ROOT / API_WORK_DIR / job_id).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    input_path = _job_input_path(job_dir, "raw-data", raw_data.filename)
    lookup_workbook_path = _job_input_path(job_dir, "daily-reference", yesterday_cleaned_data.filename)
    output_filename = output_filename_from_csv_path(input_path)
    output_path = (job_dir / output_filename).resolve()

    try:
        await _save_upload_file(raw_data, input_path)
        await _save_upload_file(yesterday_cleaned_data, lookup_workbook_path)
    except Exception as exc:
        logger.exception("Report 1 upload save failed for job %s", job_id)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    conversion_request = ConvertPathRequest(
        input_path=str(input_path),
        output_path=str(output_path),
        delimiter=delimiter or None,
        encoding=encoding or None,
        normalize_headers=normalize_headers,
        keep_empty=keep_empty,
        drop_empty_columns=drop_empty_columns,
        dedupe=dedupe,
        infer_types=infer_types,
        refresh_template=refresh_template,
    )

    with REPORT_1_JOBS_LOCK:
        REPORT_1_JOBS[job_id] = {
            "status": "queued",
            "created_at": _now_seconds(),
            "updated_at": _now_seconds(),
            "raw_data_filename": input_path.name,
            "yesterday_cleaned_data_filename": lookup_workbook_path.name,
            "filename": output_path.name,
            "refresh_template": refresh_template,
        }

    REPORT_1_EXECUTOR.submit(
        _run_report_1_upload_job,
        job_id,
        input_path,
        lookup_workbook_path,
        output_path,
        conversion_request,
    )

    return {
        "success": True,
        "message": "Report 1 conversion queued.",
        "job_id": job_id,
        "status": "queued",
        "status_url": f"/jobs/report-1/{job_id}",
        "filename": output_path.name,
        "raw_data_filename": input_path.name,
        "yesterday_cleaned_data_filename": lookup_workbook_path.name,
        "refresh_template": refresh_template,
    }


@app.post("/convert/upload")
async def convert_upload(
    request: Request,
    raw_data: UploadFile = File(...),
    yesterday_cleaned_data: UploadFile = File(...),
    delimiter: str | None = Form(default=None),
    encoding: str | None = Form(default=None),
    normalize_headers: bool = Form(default=True),
    keep_empty: bool = Form(default=False),
    drop_empty_columns: bool = Form(default=False),
    dedupe: bool = Form(default=False),
    infer_types: bool = Form(default=False),
    refresh_template: bool = Form(default=False),
) -> dict[str, object]:
    if not raw_data.filename or not raw_data.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="raw_data must be a .csv file.")
    if not yesterday_cleaned_data.filename or not yesterday_cleaned_data.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="yesterday_cleaned_data must be a .xlsx file.")

    job_dir = (APP_ROOT / API_WORK_DIR / uuid.uuid4().hex).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    input_path = _job_input_path(job_dir, "raw-data", raw_data.filename)
    lookup_workbook_path = _job_input_path(job_dir, "daily-reference", yesterday_cleaned_data.filename)
    output_filename = output_filename_from_csv_path(input_path)
    output_path = (job_dir / output_filename).resolve()

    try:
        await _save_upload_file(raw_data, input_path)
        await _save_upload_file(yesterday_cleaned_data, lookup_workbook_path)

        conversion_request = ConvertPathRequest(
            input_path=str(input_path),
            output_path=str(output_path),
            delimiter=delimiter or None,
            encoding=encoding or None,
            normalize_headers=normalize_headers,
            keep_empty=keep_empty,
            drop_empty_columns=drop_empty_columns,
            dedupe=dedupe,
            infer_types=infer_types,
            refresh_template=refresh_template,
        )
        stage_result = await run_in_threadpool(
            _run_locked_excel_stage,
            run_daily_stage,
            REPORT_1_EXCEL_TIMEOUT_SECONDS,
            input_path,
            lookup_workbook_path,
            output_path,
            conversion_request.refresh_template,
            conversion_request.delimiter,
            conversion_request.encoding,
            conversion_request.normalize_headers,
            conversion_request.keep_empty,
            conversion_request.drop_empty_columns,
            conversion_request.dedupe,
            conversion_request.infer_types,
            REPORT_1_EXCEL_TIMEOUT_SECONDS,
        )
        result = {"elapsed_seconds": round(stage_result.elapsed_seconds, 3)}
    except Exception as exc:
        logger.exception("Report 1 upload conversion failed for job %s", job_dir.name)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return {
        "success": True,
        "message": "CSV converted successfully.",
        "elapsed_seconds": result["elapsed_seconds"],
        "filename": output_path.name,
        "raw_data_filename": input_path.name,
        "yesterday_cleaned_data_filename": lookup_workbook_path.name,
        "output_file": str(output_path),
        "download_url": _download_url(request, job_dir.name, output_path.name),
        "refresh_template": refresh_template,
    }


@app.post("/convert/report-2/upload/jobs")
async def start_report_2_upload_job(
    request: Request,
    tracking_workbook: UploadFile = File(...),
    log_update_status: UploadFile = File(...),
    previous_ongoing_workbook: UploadFile = File(...),
    with_pivot: bool = Form(default=False),
) -> dict[str, object]:
    if not tracking_workbook.filename or not tracking_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="tracking_workbook must be a .xlsx file.")
    if not log_update_status.filename or not log_update_status.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="log_update_status must be a .csv file.")
    if not previous_ongoing_workbook.filename or not previous_ongoing_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="previous_ongoing_workbook must be a .xlsx file.")

    job_id = uuid.uuid4().hex
    job_dir = (APP_ROOT / API_WORK_DIR / job_id).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    tracking_path = _job_input_path(job_dir, "daily-tracking", tracking_workbook.filename)
    log_path = _job_input_path(job_dir, "log-update", log_update_status.filename)
    previous_ongoing_path = _job_input_path(job_dir, "ongoing-reference", previous_ongoing_workbook.filename)
    output_filename = ongoing_output_filename_for_tracking(tracking_path)
    output_path = (job_dir / output_filename).resolve()

    try:
        await _save_upload_file(tracking_workbook, tracking_path)
        await _save_upload_file(log_update_status, log_path)
        await _save_upload_file(previous_ongoing_workbook, previous_ongoing_path)
    except Exception as exc:
        logger.exception("Report 2 upload save failed for job %s", job_id)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    with REPORT_2_JOBS_LOCK:
        REPORT_2_JOBS[job_id] = {
            "status": "queued",
            "created_at": _now_seconds(),
            "updated_at": _now_seconds(),
            "tracking_workbook_filename": tracking_path.name,
            "log_update_status_filename": log_path.name,
            "previous_ongoing_workbook_filename": previous_ongoing_path.name,
            "filename": output_path.name,
            "with_pivot": with_pivot,
        }

    REPORT_2_EXECUTOR.submit(
        _run_report_2_upload_job,
        job_id,
        tracking_path,
        log_path,
        previous_ongoing_path,
        output_path,
        with_pivot,
    )

    return {
        "success": True,
        "message": "Report 2 conversion queued.",
        "job_id": job_id,
        "status": "queued",
        "status_url": f"/jobs/report-2/{job_id}",
        "filename": output_path.name,
        "tracking_workbook_filename": tracking_path.name,
        "log_update_status_filename": log_path.name,
        "previous_ongoing_workbook_filename": previous_ongoing_path.name,
        "with_pivot": with_pivot,
    }


@app.post("/convert/report-2/upload")
async def convert_report_2_upload(
    request: Request,
    tracking_workbook: UploadFile = File(...),
    log_update_status: UploadFile = File(...),
    previous_ongoing_workbook: UploadFile = File(...),
    with_pivot: bool = Form(default=False),
) -> dict[str, object]:
    if not tracking_workbook.filename or not tracking_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="tracking_workbook must be a .xlsx file.")
    if not log_update_status.filename or not log_update_status.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="log_update_status must be a .csv file.")
    if not previous_ongoing_workbook.filename or not previous_ongoing_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="previous_ongoing_workbook must be a .xlsx file.")

    job_dir = (APP_ROOT / API_WORK_DIR / uuid.uuid4().hex).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    tracking_path = _job_input_path(job_dir, "daily-tracking", tracking_workbook.filename)
    log_path = _job_input_path(job_dir, "log-update", log_update_status.filename)
    previous_ongoing_path = _job_input_path(job_dir, "ongoing-reference", previous_ongoing_workbook.filename)
    output_filename = ongoing_output_filename_for_tracking(tracking_path)
    output_path = (job_dir / output_filename).resolve()
    conversion_output_path = output_path

    try:
        await _save_upload_file(tracking_workbook, tracking_path)
        await _save_upload_file(log_update_status, log_path)
        await _save_upload_file(previous_ongoing_workbook, previous_ongoing_path)
        conversion_output_path.parent.mkdir(parents=True, exist_ok=True)

        stage_result = await run_in_threadpool(
            _run_locked_excel_stage,
            run_ongoing_stage,
            REPORT_2_EXCEL_TIMEOUT_SECONDS,
            tracking_path,
            log_path,
            previous_ongoing_path,
            conversion_output_path,
            with_pivot,
            None,
            REPORT_2_EXCEL_TIMEOUT_SECONDS,
        )
        result = {"elapsed_seconds": round(stage_result.elapsed_seconds, 3)}
    except Exception as exc:
        logger.exception("Report 2 upload conversion failed for job %s", job_dir.name)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return {
        "success": True,
        "message": "Report 2 workbook generated successfully.",
        "elapsed_seconds": result["elapsed_seconds"],
        "filename": output_path.name,
        "tracking_workbook_filename": tracking_path.name,
        "log_update_status_filename": log_path.name,
        "previous_ongoing_workbook_filename": previous_ongoing_path.name,
        "output_file": str(output_path),
        "download_url": _download_url(request, job_dir.name, output_path.name),
        "with_pivot": with_pivot,
    }


@app.post("/convert/report-3/upload/jobs")
async def start_report_3_upload_job(
    request: Request,
    tracking_workbook: UploadFile = File(...),
    previous_iphone_workbook: UploadFile = File(...),
) -> dict[str, object]:
    if not tracking_workbook.filename or not tracking_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="tracking_workbook must be a .xlsx file.")
    if not previous_iphone_workbook.filename or not previous_iphone_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="previous_iphone_workbook must be a .xlsx file.")

    job_id = uuid.uuid4().hex
    job_dir = (APP_ROOT / API_WORK_DIR / job_id).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    tracking_path = _job_input_path(job_dir, "daily-tracking", tracking_workbook.filename)
    reference_path = _job_input_path(job_dir, "iphone-reference", previous_iphone_workbook.filename)
    output_filename = iphone_default_output_path(tracking_path).name
    output_path = (job_dir / output_filename).resolve()

    try:
        await _save_upload_file(tracking_workbook, tracking_path)
        await _save_upload_file(previous_iphone_workbook, reference_path)
    except Exception as exc:
        logger.exception("Report 3 upload save failed for job %s", job_id)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    with REPORT_3_JOBS_LOCK:
        REPORT_3_JOBS[job_id] = {
            "status": "queued",
            "created_at": _now_seconds(),
            "updated_at": _now_seconds(),
            "tracking_workbook_filename": tracking_path.name,
            "previous_iphone_workbook_filename": reference_path.name,
            "filename": output_path.name,
        }

    REPORT_3_EXECUTOR.submit(
        _run_report_3_upload_job,
        job_id,
        tracking_path,
        reference_path,
        output_path,
    )

    return {
        "success": True,
        "message": "Report 3 conversion queued.",
        "job_id": job_id,
        "status": "queued",
        "status_url": f"/jobs/report-3/{job_id}",
        "filename": output_path.name,
        "tracking_workbook_filename": tracking_path.name,
        "previous_iphone_workbook_filename": reference_path.name,
    }


@app.post("/convert/report-3/upload")
async def convert_report_3_upload(
    request: Request,
    tracking_workbook: UploadFile = File(...),
    previous_iphone_workbook: UploadFile = File(...),
) -> dict[str, object]:
    if not tracking_workbook.filename or not tracking_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="tracking_workbook must be a .xlsx file.")
    if not previous_iphone_workbook.filename or not previous_iphone_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="previous_iphone_workbook must be a .xlsx file.")

    job_dir = (APP_ROOT / API_WORK_DIR / uuid.uuid4().hex).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    tracking_path = _job_input_path(job_dir, "daily-tracking", tracking_workbook.filename)
    reference_path = _job_input_path(job_dir, "iphone-reference", previous_iphone_workbook.filename)
    output_filename = iphone_default_output_path(tracking_path).name
    output_path = (job_dir / output_filename).resolve()

    try:
        await _save_upload_file(tracking_workbook, tracking_path)
        await _save_upload_file(previous_iphone_workbook, reference_path)
        stage_result = await run_in_threadpool(
            _run_locked_excel_stage,
            run_iphone_stage,
            REPORT_3_EXCEL_TIMEOUT_SECONDS,
            tracking_path,
            reference_path,
            output_path,
            REPORT_3_EXCEL_TIMEOUT_SECONDS,
        )
        result = {"elapsed_seconds": round(stage_result.elapsed_seconds, 3)}
    except Exception as exc:
        logger.exception("Report 3 upload conversion failed for job %s", job_dir.name)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return {
        "success": True,
        "message": "Report 3 iPhone workbook generated successfully.",
        "elapsed_seconds": result["elapsed_seconds"],
        "filename": output_path.name,
        "tracking_workbook_filename": tracking_path.name,
        "previous_iphone_workbook_filename": reference_path.name,
        "output_file": str(output_path),
        "download_url": _download_url(request, job_dir.name, output_path.name),
    }


@app.post("/convert/report-4/upload/jobs")
async def start_report_4_upload_job(
    request: Request,
    raw_ide_workbook: UploadFile = File(...),
    previous_ide_workbook: UploadFile = File(...),
    collabs_csv: UploadFile = File(...),
) -> dict[str, object]:
    if not raw_ide_workbook.filename or not raw_ide_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="raw_ide_workbook must be a .xlsx file.")
    if not previous_ide_workbook.filename or not previous_ide_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="previous_ide_workbook must be a .xlsx file.")
    if not collabs_csv.filename or not collabs_csv.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="collabs_csv must be a .csv file.")
    job_id = uuid.uuid4().hex
    job_dir = (APP_ROOT / API_WORK_DIR / job_id).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    raw_path = _job_input_path(job_dir, "raw-ide", raw_ide_workbook.filename)
    previous_path = _job_input_path(job_dir, "ide-reference", previous_ide_workbook.filename)
    collabs_path = _job_input_path(job_dir, "collabs", collabs_csv.filename)

    try:
        await _save_upload_file(raw_ide_workbook, raw_path)
        await _save_upload_file(previous_ide_workbook, previous_path)
        await _save_upload_file(collabs_csv, collabs_path)
        resolved_report_date = determine_ide_report_date(None, raw_path)
    except Exception as exc:
        logger.exception("Report 4 upload save failed for job %s", job_id)
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    output_path = (job_dir / ide_output_filename(resolved_report_date)).resolve()
    with REPORT_4_JOBS_LOCK:
        REPORT_4_JOBS[job_id] = {
            "status": "queued",
            "created_at": _now_seconds(),
            "updated_at": _now_seconds(),
            "raw_ide_workbook_filename": raw_path.name,
            "previous_ide_workbook_filename": previous_path.name,
            "collabs_csv_filename": collabs_path.name,
            "report_date": resolved_report_date.isoformat(),
            "filename": output_path.name,
        }

    REPORT_4_EXECUTOR.submit(
        _run_report_4_upload_job,
        job_id,
        raw_path,
        previous_path,
        collabs_path,
        output_path,
        resolved_report_date,
    )

    return {
        "success": True,
        "message": "Report 4 conversion queued.",
        "job_id": job_id,
        "status": "queued",
        "status_url": f"/jobs/report-4/{job_id}",
        "filename": output_path.name,
        "report_date": resolved_report_date.isoformat(),
        "raw_ide_workbook_filename": raw_path.name,
        "previous_ide_workbook_filename": previous_path.name,
        "collabs_csv_filename": collabs_path.name,
    }


@app.post("/convert/pipeline/upload/jobs")
async def start_pipeline_upload_job(
    request: Request,
    raw_data: UploadFile = File(...),
    yesterday_cleaned_data: UploadFile = File(...),
    log_update_status: UploadFile = File(...),
    previous_ongoing_workbook: UploadFile = File(...),
    previous_iphone_workbook: UploadFile = File(...),
    job_id: str | None = Form(default=None),
    refresh_template: bool = Form(default=True),
    ongoing_with_pivot: bool = Form(default=False),
) -> dict[str, object]:
    if not raw_data.filename or not raw_data.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="raw_data must be a .csv file.")
    if not yesterday_cleaned_data.filename or not yesterday_cleaned_data.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="yesterday_cleaned_data must be a .xlsx file.")
    if not log_update_status.filename or not log_update_status.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="log_update_status must be a .csv file.")
    if not previous_ongoing_workbook.filename or not previous_ongoing_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="previous_ongoing_workbook must be a .xlsx file.")
    if not previous_iphone_workbook.filename or not previous_iphone_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="previous_iphone_workbook must be a .xlsx file.")

    if job_id:
        try:
            job_id = uuid.UUID(job_id).hex
        except ValueError as exc:
            raise HTTPException(status_code=400, detail="job_id must be a valid UUID.") from exc
    else:
        job_id = uuid.uuid4().hex

    with PIPELINE_JOBS_LOCK:
        if job_id in PIPELINE_JOBS:
            raise HTTPException(status_code=409, detail="Pipeline job already exists.")
        PIPELINE_JOBS[job_id] = {
            "status": "uploading",
            "created_at": _now_seconds(),
            "updated_at": _now_seconds(),
        }

    job_dir = (APP_ROOT / API_WORK_DIR / job_id).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    raw_path = _job_input_path(job_dir, "raw-data", raw_data.filename)
    daily_reference_path = _job_input_path(job_dir, "daily-reference", yesterday_cleaned_data.filename)
    log_path = _job_input_path(job_dir, "log-update", log_update_status.filename)
    previous_ongoing_path = _job_input_path(job_dir, "ongoing-reference", previous_ongoing_workbook.filename)
    previous_iphone_path = _job_input_path(job_dir, "iphone-reference", previous_iphone_workbook.filename)
    output_dir = (job_dir / "pipeline-output").resolve()
    pipeline_zip_path = (job_dir / "daily-pipeline-output.zip").resolve()

    try:
        await _save_upload_file(raw_data, raw_path)
        await _save_upload_file(yesterday_cleaned_data, daily_reference_path)
        await _save_upload_file(log_update_status, log_path)
        await _save_upload_file(previous_ongoing_workbook, previous_ongoing_path)
        await _save_upload_file(previous_iphone_workbook, previous_iphone_path)
    except Exception as exc:
        logger.exception("Pipeline upload save failed for job %s", job_id)
        _set_pipeline_job(
            job_id,
            status="failed",
            completed_at=_now_seconds(),
            error=str(exc),
        )
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    with PIPELINE_JOBS_LOCK:
        PIPELINE_JOBS[job_id].update({
            "status": "queued",
            "updated_at": _now_seconds(),
            "raw_data_filename": raw_path.name,
            "yesterday_cleaned_data_filename": daily_reference_path.name,
            "log_update_status_filename": log_path.name,
            "previous_ongoing_workbook_filename": previous_ongoing_path.name,
            "previous_iphone_workbook_filename": previous_iphone_path.name,
            "filename": pipeline_zip_path.name,
            "refresh_template": refresh_template,
            "ongoing_with_pivot": ongoing_with_pivot,
        })

    PIPELINE_EXECUTOR.submit(
        _run_pipeline_upload_job,
        job_id,
        raw_path,
        daily_reference_path,
        log_path,
        previous_ongoing_path,
        previous_iphone_path,
        output_dir,
        pipeline_zip_path,
        refresh_template,
        ongoing_with_pivot,
    )

    return {
        "success": True,
        "message": "Daily pipeline queued.",
        "job_id": job_id,
        "status": "queued",
        "status_url": f"/jobs/pipeline/{job_id}",
        "filename": pipeline_zip_path.name,
        "raw_data_filename": raw_path.name,
        "yesterday_cleaned_data_filename": daily_reference_path.name,
        "log_update_status_filename": log_path.name,
        "previous_ongoing_workbook_filename": previous_ongoing_path.name,
        "previous_iphone_workbook_filename": previous_iphone_path.name,
        "refresh_template": refresh_template,
        "ongoing_with_pivot": ongoing_with_pivot,
    }


@app.post("/convert/pipeline/upload")
async def convert_pipeline_upload(
    request: Request,
    raw_data: UploadFile = File(...),
    yesterday_cleaned_data: UploadFile = File(...),
    log_update_status: UploadFile = File(...),
    previous_ongoing_workbook: UploadFile = File(...),
    previous_iphone_workbook: UploadFile = File(...),
    refresh_template: bool = Form(default=True),
    ongoing_with_pivot: bool = Form(default=False),
) -> dict[str, object]:
    if not raw_data.filename or not raw_data.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="raw_data must be a .csv file.")
    if not yesterday_cleaned_data.filename or not yesterday_cleaned_data.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="yesterday_cleaned_data must be a .xlsx file.")
    if not log_update_status.filename or not log_update_status.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="log_update_status must be a .csv file.")
    if not previous_ongoing_workbook.filename or not previous_ongoing_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="previous_ongoing_workbook must be a .xlsx file.")
    if not previous_iphone_workbook.filename or not previous_iphone_workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="previous_iphone_workbook must be a .xlsx file.")

    job_dir = (APP_ROOT / API_WORK_DIR / uuid.uuid4().hex).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    raw_path = _job_input_path(job_dir, "raw-data", raw_data.filename)
    daily_reference_path = _job_input_path(job_dir, "daily-reference", yesterday_cleaned_data.filename)
    log_path = _job_input_path(job_dir, "log-update", log_update_status.filename)
    previous_ongoing_path = _job_input_path(job_dir, "ongoing-reference", previous_ongoing_workbook.filename)
    previous_iphone_path = _job_input_path(job_dir, "iphone-reference", previous_iphone_workbook.filename)
    conversion_output_dir = (job_dir / "pipeline-output").resolve()
    pipeline_zip_path = job_dir / "daily-pipeline-output.zip"

    try:
        await _save_upload_file(raw_data, raw_path)
        await _save_upload_file(yesterday_cleaned_data, daily_reference_path)
        await _save_upload_file(log_update_status, log_path)
        await _save_upload_file(previous_ongoing_workbook, previous_ongoing_path)
        await _save_upload_file(previous_iphone_workbook, previous_iphone_path)

        result = await run_in_threadpool(
            _run_pipeline_conversion,
            raw_path,
            daily_reference_path,
            log_path,
            previous_ongoing_path,
            previous_iphone_path,
            conversion_output_dir,
            refresh_template,
            ongoing_with_pivot,
        )

        output_files = [Path(path) for path in result["output_files"]]
        with zipfile.ZipFile(pipeline_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for output_file in output_files:
                archive.write(output_file, arcname=output_file.name)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return {
        "success": True,
        "message": "Daily pipeline generated successfully.",
        "elapsed_seconds": result["elapsed_seconds"],
        "filename": pipeline_zip_path.name,
        "output_files": [Path(path).name for path in result["output_files"]],
        "download_url": _download_url(request, job_dir.name, pipeline_zip_path.name),
        "refresh_template": refresh_template,
        "ongoing_with_pivot": ongoing_with_pivot,
    }
