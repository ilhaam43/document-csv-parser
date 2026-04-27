from __future__ import annotations

import contextlib
import io
import time
import uuid
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from csv_to_excel_api import (
    ConvertOptions,
    convert_many,
    convert_one,
    output_filename_from_csv_path,
    resolve_csv_files,
    resolve_output_path,
)


APP_ROOT = Path(__file__).resolve().parent
API_WORK_DIR = Path("output-today/api")

app = FastAPI(
    title="Report CSV Parser API",
    version="1.0.0",
)


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

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_path.name,
    )


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
    output_path = job_dir / output_filename_from_csv_path(input_path)

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
        result = await run_in_threadpool(_run_conversion, input_path, output_path, conversion_request, lookup_workbook_path)
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
