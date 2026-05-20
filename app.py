from __future__ import annotations

import contextlib
import io
import shutil
import time
import uuid
import zipfile
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from csv_to_excel import (
    ConvertOptions,
    convert_many,
    convert_one,
    output_filename_from_csv_path,
    resolve_csv_files,
    resolve_output_path,
)
from csv_to_excel_on_going import (
    apply_logic_to_workbook as apply_ongoing_logic_to_workbook,
    output_filename_for_tracking as ongoing_output_filename_for_tracking,
)
from generate_iphone_tracking import (
    default_output_path as iphone_default_output_path,
    generate_iphone_tracking,
)


APP_ROOT = Path(__file__).resolve().parent
CONVERSION_OUTPUT_DIR = Path("output-today")
ONGOING_CONVERSION_OUTPUT_DIR = Path("output-outgoing")
IPHONE_CONVERSION_OUTPUT_DIR = Path("output-iphone")
PIPELINE_CONVERSION_OUTPUT_DIR = Path("output-pipeline")
API_WORK_DIR = Path("output-today/api")

app = FastAPI(
    title="Report CSV Parser API",
    version="1.0.0",
)
app.mount("/static", StaticFiles(directory=APP_ROOT / "static"), name="static")


@contextlib.contextmanager
def _initialized_com_thread():
    """Initialize Windows COM for worker threads that automate Excel."""
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
    base_url = str(request.base_url).rstrip("/")
    return f"{base_url}/outputs/{job_id}/{quote(filename)}"


async def _save_upload_file(upload: UploadFile, destination: Path) -> None:
    with destination.open("wb") as handle:
        while chunk := await upload.read(1024 * 1024):
            handle.write(chunk)


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


def _run_ongoing_conversion(
    tracking_workbook_path: Path,
    log_csv_path: Path,
    previous_ongoing_workbook_path: Path,
    output_path: Path,
    with_pivot: bool,
) -> dict[str, object]:
    stdout = io.StringIO()
    stderr = io.StringIO()
    started = time.perf_counter()
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        with _initialized_com_thread():
            apply_ongoing_logic_to_workbook(
                tracking_workbook_path,
                log_csv_path,
                output_path,
                with_pivot=with_pivot,
                engine="com",
                clone_pivot_template=True,
                base_workbook_path=previous_ongoing_workbook_path,
            )
    elapsed_seconds = time.perf_counter() - started

    return {
        "elapsed_seconds": round(elapsed_seconds, 3),
        "output_path": str(output_path),
        "stdout": stdout.getvalue().strip(),
        "stderr": stderr.getvalue().strip(),
        "with_pivot": with_pivot,
    }


def _run_iphone_conversion(
    tracking_workbook_path: Path,
    reference_workbook_path: Path,
    output_path: Path,
) -> dict[str, object]:
    stdout = io.StringIO()
    stderr = io.StringIO()
    started = time.perf_counter()
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        with _initialized_com_thread():
            generate_iphone_tracking(tracking_workbook_path, reference_workbook_path, output_path)
    elapsed_seconds = time.perf_counter() - started

    return {
        "elapsed_seconds": round(elapsed_seconds, 3),
        "output_path": str(output_path),
        "stdout": stdout.getvalue().strip(),
        "stderr": stderr.getvalue().strip(),
    }


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
    stdout = io.StringIO()
    stderr = io.StringIO()
    started = time.perf_counter()
    output_dir.mkdir(parents=True, exist_ok=True)

    daily_output_path = (output_dir / output_filename_from_csv_path(data_order_csv_path)).resolve()
    ongoing_output_path = (output_dir / ongoing_output_filename_for_tracking(daily_output_path)).resolve()
    iphone_output_path = (output_dir / iphone_default_output_path(daily_output_path).name).resolve()

    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        with _initialized_com_thread():
            daily_request = ConvertPathRequest(
                input_path=str(data_order_csv_path),
                output_path=str(daily_output_path),
                refresh_template=refresh_template,
            )
            _run_conversion(
                data_order_csv_path,
                daily_output_path,
                daily_request,
                daily_reference_workbook_path,
            )
            _run_ongoing_conversion(
                daily_output_path,
                log_csv_path,
                previous_ongoing_workbook_path,
                ongoing_output_path,
                ongoing_with_pivot,
            )
            _run_iphone_conversion(
                daily_output_path,
                previous_iphone_workbook_path,
                iphone_output_path,
            )
    elapsed_seconds = time.perf_counter() - started

    return {
        "elapsed_seconds": round(elapsed_seconds, 3),
        "output_files": [str(daily_output_path), str(ongoing_output_path), str(iphone_output_path)],
        "stdout": stdout.getvalue().strip(),
        "stderr": stderr.getvalue().strip(),
        "refresh_template": refresh_template,
        "ongoing_with_pivot": ongoing_with_pivot,
    }


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


@app.get("/pipeline")
def pipeline_page() -> FileResponse:
    return FileResponse(APP_ROOT / "templates" / "pipeline.html", media_type="text/html")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


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
    return FileResponse(output_path, media_type=media_type, filename=output_path.name)


@app.post("/convert/path")
async def convert_path(request: ConvertPathRequest) -> dict[str, object]:
    input_path = _resolve_path(request.input_path)
    output_path = _resolve_path(request.output_path) if request.output_path else None
    try:
        return await run_in_threadpool(_run_conversion, input_path, output_path, request)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


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
    input_path = job_dir / Path(raw_data.filename).name
    lookup_workbook_path = job_dir / Path(yesterday_cleaned_data.filename).name
    output_filename = output_filename_from_csv_path(input_path)
    conversion_output_path = (APP_ROOT / CONVERSION_OUTPUT_DIR / output_filename).resolve()
    output_path = job_dir / output_filename

    try:
        await _save_upload_file(raw_data, input_path)
        await _save_upload_file(yesterday_cleaned_data, lookup_workbook_path)

        conversion_request = ConvertPathRequest(
            input_path=str(input_path),
            output_path=str(conversion_output_path),
            delimiter=delimiter or None,
            encoding=encoding or None,
            normalize_headers=normalize_headers,
            keep_empty=keep_empty,
            drop_empty_columns=drop_empty_columns,
            dedupe=dedupe,
            infer_types=infer_types,
            refresh_template=refresh_template,
        )
        result = await run_in_threadpool(
            _run_conversion,
            input_path,
            conversion_output_path,
            conversion_request,
            lookup_workbook_path,
        )
        shutil.copy2(conversion_output_path, output_path)
    except Exception as exc:
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
    tracking_path = job_dir / Path(tracking_workbook.filename).name
    log_path = job_dir / Path(log_update_status.filename).name
    previous_ongoing_path = job_dir / Path(previous_ongoing_workbook.filename).name
    output_filename = ongoing_output_filename_for_tracking(tracking_path)
    conversion_output_path = (APP_ROOT / ONGOING_CONVERSION_OUTPUT_DIR / output_filename).resolve()
    output_path = job_dir / output_filename

    try:
        await _save_upload_file(tracking_workbook, tracking_path)
        await _save_upload_file(log_update_status, log_path)
        await _save_upload_file(previous_ongoing_workbook, previous_ongoing_path)
        conversion_output_path.parent.mkdir(parents=True, exist_ok=True)

        result = await run_in_threadpool(
            _run_ongoing_conversion,
            tracking_path,
            log_path,
            previous_ongoing_path,
            conversion_output_path,
            with_pivot,
        )
        shutil.copy2(conversion_output_path, output_path)
    except Exception as exc:
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
    tracking_path = job_dir / Path(tracking_workbook.filename).name
    reference_path = job_dir / Path(previous_iphone_workbook.filename).name
    output_filename = iphone_default_output_path(tracking_path).name
    conversion_output_path = (APP_ROOT / IPHONE_CONVERSION_OUTPUT_DIR / output_filename).resolve()
    output_path = job_dir / output_filename

    try:
        await _save_upload_file(tracking_workbook, tracking_path)
        await _save_upload_file(previous_iphone_workbook, reference_path)
        conversion_output_path.parent.mkdir(parents=True, exist_ok=True)

        result = await run_in_threadpool(
            _run_iphone_conversion,
            tracking_path,
            reference_path,
            conversion_output_path,
        )
        shutil.copy2(conversion_output_path, output_path)
    except Exception as exc:
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
    input_dir = job_dir / "input"
    input_dir.mkdir(parents=True, exist_ok=True)

    raw_path = input_dir / Path(raw_data.filename).name
    daily_reference_path = input_dir / Path(yesterday_cleaned_data.filename).name
    log_path = input_dir / Path(log_update_status.filename).name
    previous_ongoing_path = input_dir / Path(previous_ongoing_workbook.filename).name
    previous_iphone_path = input_dir / Path(previous_iphone_workbook.filename).name
    conversion_output_dir = (APP_ROOT / PIPELINE_CONVERSION_OUTPUT_DIR).resolve()
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
