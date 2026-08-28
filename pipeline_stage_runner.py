#!/usr/bin/env python3
"""Run one daily-pipeline stage in an isolated Python process."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
import uuid
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path


APP_ROOT = Path(__file__).resolve().parent
DEFAULT_STAGE_TIMEOUT_SECONDS = 25 * 60
FILE_READY_TIMEOUT_SECONDS = 30
MAX_TRANSIENT_STAGE_ATTEMPTS = 3
REQUIRED_XLSX_MEMBERS = {
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
}
TRANSIENT_EXCEL_ERROR_MARKERS = (
    "cannot access the file",
    "being used by another process",
    "same name as a currently open workbook",
    "couldn't paste this data because it took too long",
    "call was rejected by callee",
    "ole error",
)


@dataclass(frozen=True)
class StageProcessResult:
    stdout: str
    stderr: str
    elapsed_seconds: float
    metadata: dict[str, object] | None = None


def wait_for_xlsx(path: Path, timeout_seconds: int = FILE_READY_TIMEOUT_SECONDS) -> Path:
    """Wait until an isolated stage has produced a complete, readable workbook."""
    resolved = path.resolve()
    deadline = time.monotonic() + timeout_seconds
    last_error = "file does not exist"

    while True:
        try:
            if not resolved.is_file() or resolved.stat().st_size == 0:
                raise FileNotFoundError(last_error)
            with zipfile.ZipFile(resolved) as workbook:
                missing_members = REQUIRED_XLSX_MEMBERS.difference(workbook.namelist())
                if missing_members:
                    raise ValueError(
                        f"missing workbook members: {', '.join(sorted(missing_members))}"
                    )
                corrupt_member = workbook.testzip()
                if corrupt_member is not None:
                    raise ValueError(f"corrupt workbook member: {corrupt_member}")
            return resolved
        except (FileNotFoundError, OSError, ValueError, zipfile.BadZipFile) as exc:
            last_error = str(exc)
            remaining_seconds = deadline - time.monotonic()
            if remaining_seconds <= 0:
                break
            time.sleep(min(0.5, remaining_seconds))

    raise RuntimeError(f"Workbook is not ready after {timeout_seconds} seconds: {resolved} ({last_error})")


def is_transient_excel_error(message: str) -> bool:
    normalized = message.casefold()
    return any(marker in normalized for marker in TRANSIENT_EXCEL_ERROR_MARKERS)


def _run_worker(arguments: list[str], output_path: Path, timeout_seconds: int) -> StageProcessResult:
    command = [sys.executable, str(Path(__file__).resolve()), *arguments]
    creation_flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
    started = time.perf_counter()
    deadline = time.monotonic() + timeout_seconds

    for attempt in range(1, MAX_TRANSIENT_STAGE_ATTEMPTS + 1):
        remaining_seconds = max(1, int(deadline - time.monotonic()))
        try:
            completed = subprocess.run(
                command,
                cwd=APP_ROOT,
                capture_output=True,
                text=True,
                timeout=remaining_seconds,
                check=False,
                creationflags=creation_flags,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(
                f"Pipeline stage exceeded {timeout_seconds // 60} minutes while generating {output_path.name}."
            ) from exc

        stdout = completed.stdout.strip()
        stderr = completed.stderr.strip()
        if completed.returncode == 0:
            wait_for_xlsx(output_path)
            return StageProcessResult(
                stdout=stdout,
                stderr=stderr,
                elapsed_seconds=time.perf_counter() - started,
            )

        detail = stderr or stdout or f"worker exited with code {completed.returncode}"
        if (
            attempt == MAX_TRANSIENT_STAGE_ATTEMPTS
            or not is_transient_excel_error(detail)
            or time.monotonic() >= deadline
        ):
            raise RuntimeError(f"Pipeline stage failed for {output_path.name}: {detail}")

        time.sleep(2 * attempt)

    raise AssertionError("unreachable")


def _working_output_path(output_path: Path) -> Path:
    working_directory = output_path.parent / f".{output_path.stem}.{uuid.uuid4().hex}.work"
    return working_directory / output_path.name


def _publish_atomic_output(
    arguments: list[str],
    output_path: Path,
    working_output_path: Path,
    timeout_seconds: int,
) -> StageProcessResult:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        result = _run_worker(arguments, working_output_path, timeout_seconds)
        working_output_path.replace(output_path)
        wait_for_xlsx(output_path)
        return result
    finally:
        try:
            working_output_path.unlink(missing_ok=True)
        except OSError:
            pass
        try:
            working_output_path.parent.rmdir()
        except OSError:
            pass


def run_daily_stage(
    data_order_csv: Path,
    reference_workbook: Path,
    output_path: Path,
    refresh_template: bool,
    delimiter: str | None = None,
    encoding: str | None = None,
    normalize_headers: bool = True,
    keep_empty: bool = False,
    drop_empty_columns: bool = False,
    dedupe: bool = False,
    infer_types: bool = False,
    timeout_seconds: int = DEFAULT_STAGE_TIMEOUT_SECONDS,
) -> StageProcessResult:
    working_output_path = _working_output_path(output_path)
    arguments = [
        "daily",
        "--input",
        str(data_order_csv.resolve()),
        "--reference",
        str(reference_workbook.resolve()),
        "--output",
        str(working_output_path.resolve()),
    ]
    if delimiter:
        arguments.extend(("--delimiter", delimiter))
    if encoding:
        arguments.extend(("--encoding", encoding))
    if not normalize_headers:
        arguments.append("--no-normalize-headers")
    if keep_empty:
        arguments.append("--keep-empty")
    if drop_empty_columns:
        arguments.append("--drop-empty-columns")
    if dedupe:
        arguments.append("--dedupe")
    if infer_types:
        arguments.append("--infer-types")
    if not refresh_template:
        arguments.append("--skip-template-refresh")
    return _publish_atomic_output(arguments, output_path, working_output_path, timeout_seconds)


def run_ongoing_stage(
    tracking_workbook: Path,
    log_csv: Path,
    reference_workbook: Path,
    output_path: Path,
    with_pivot: bool,
    aging_date: date | None = None,
    timeout_seconds: int = DEFAULT_STAGE_TIMEOUT_SECONDS,
) -> StageProcessResult:
    working_output_path = _working_output_path(output_path)
    arguments = [
        "ongoing",
        "--tracking",
        str(tracking_workbook.resolve()),
        "--log-update",
        str(log_csv.resolve()),
        "--reference",
        str(reference_workbook.resolve()),
        "--output",
        str(working_output_path.resolve()),
    ]
    if with_pivot:
        arguments.append("--with-pivot")
    if aging_date is not None:
        arguments.extend(("--aging-date", aging_date.isoformat()))
    return _publish_atomic_output(arguments, output_path, working_output_path, timeout_seconds)


def run_iphone_stage(
    tracking_workbook: Path,
    reference_workbook: Path,
    output_path: Path,
    timeout_seconds: int = DEFAULT_STAGE_TIMEOUT_SECONDS,
) -> StageProcessResult:
    working_output_path = _working_output_path(output_path)
    return _publish_atomic_output(
        [
            "iphone",
            "--tracking",
            str(tracking_workbook.resolve()),
            "--reference",
            str(reference_workbook.resolve()),
            "--output",
            str(working_output_path.resolve()),
        ],
        output_path,
        working_output_path,
        timeout_seconds,
    )


def run_ide_stage(
    raw_workbook: Path,
    previous_workbook: Path,
    collabs_csv: Path,
    output_path: Path,
    report_date: date,
    timeout_seconds: int = DEFAULT_STAGE_TIMEOUT_SECONDS,
) -> StageProcessResult:
    working_output_path = _working_output_path(output_path)
    metadata_path = working_output_path.with_suffix(".json")
    try:
        result = _publish_atomic_output(
            [
                "ide",
                "--raw",
                str(raw_workbook.resolve()),
                "--reference",
                str(previous_workbook.resolve()),
                "--collabs",
                str(collabs_csv.resolve()),
                "--report-date",
                report_date.isoformat(),
                "--output",
                str(working_output_path.resolve()),
                "--metadata-output",
                str(metadata_path.resolve()),
            ],
            output_path,
            working_output_path,
            timeout_seconds,
        )
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        return StageProcessResult(
            stdout=result.stdout,
            stderr=result.stderr,
            elapsed_seconds=result.elapsed_seconds,
            metadata=metadata,
        )
    finally:
        metadata_path.unlink(missing_ok=True)
        try:
            working_output_path.parent.rmdir()
        except OSError:
            pass


def _run_daily(args: argparse.Namespace) -> None:
    from csv_to_excel import ConvertOptions, convert_one

    options = ConvertOptions(
        output=args.output,
        delimiter=args.delimiter,
        encoding=args.encoding,
        normalize_headers=not args.no_normalize_headers,
        keep_empty=args.keep_empty,
        drop_empty_columns=args.drop_empty_columns,
        dedupe=args.dedupe,
        infer_types=args.infer_types,
        combine=False,
        refresh_template=not args.skip_template_refresh,
        lookup_workbook=args.reference,
    )
    convert_one(args.input, args.output, options)


def _run_ongoing(args: argparse.Namespace) -> None:
    from csv_to_excel_on_going import apply_logic_to_workbook

    apply_logic_to_workbook(
        args.tracking,
        args.log_update,
        args.output,
        with_pivot=args.with_pivot,
        engine="com",
        clone_pivot_template=True,
        aging_date=args.aging_date,
        base_workbook_path=args.reference,
    )


def _run_iphone(args: argparse.Namespace) -> None:
    from generate_iphone_tracking import generate_iphone_tracking

    generate_iphone_tracking(args.tracking, args.reference, args.output)


def _run_ide(args: argparse.Namespace) -> None:
    from generate_ide_tracking import InputFiles, generate_ide_tracking

    result = generate_ide_tracking(
        InputFiles(raw=args.raw, previous=args.reference, collabs=args.collabs),
        args.output,
        args.report_date,
    )
    args.metadata_output.write_text(json.dumps(result, default=str), encoding="utf-8")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="stage", required=True)

    daily = subparsers.add_parser("daily")
    daily.add_argument("--input", type=Path, required=True)
    daily.add_argument("--reference", type=Path, required=True)
    daily.add_argument("--output", type=Path, required=True)
    daily.add_argument("--delimiter")
    daily.add_argument("--encoding")
    daily.add_argument("--no-normalize-headers", action="store_true")
    daily.add_argument("--keep-empty", action="store_true")
    daily.add_argument("--drop-empty-columns", action="store_true")
    daily.add_argument("--dedupe", action="store_true")
    daily.add_argument("--infer-types", action="store_true")
    daily.add_argument("--skip-template-refresh", action="store_true")

    ongoing = subparsers.add_parser("ongoing")
    ongoing.add_argument("--tracking", type=Path, required=True)
    ongoing.add_argument("--log-update", type=Path, required=True)
    ongoing.add_argument("--reference", type=Path, required=True)
    ongoing.add_argument("--output", type=Path, required=True)
    ongoing.add_argument("--with-pivot", action="store_true")
    ongoing.add_argument("--aging-date", type=lambda value: datetime.strptime(value, "%Y-%m-%d").date())

    iphone = subparsers.add_parser("iphone")
    iphone.add_argument("--tracking", type=Path, required=True)
    iphone.add_argument("--reference", type=Path, required=True)
    iphone.add_argument("--output", type=Path, required=True)

    ide = subparsers.add_parser("ide")
    ide.add_argument("--raw", type=Path, required=True)
    ide.add_argument("--reference", type=Path, required=True)
    ide.add_argument("--collabs", type=Path, required=True)
    ide.add_argument("--report-date", type=lambda value: datetime.strptime(value, "%Y-%m-%d").date(), required=True)
    ide.add_argument("--output", type=Path, required=True)
    ide.add_argument("--metadata-output", type=Path, required=True)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    args.output.parent.mkdir(parents=True, exist_ok=True)
    try:
        if args.stage == "daily":
            _run_daily(args)
        elif args.stage == "ongoing":
            _run_ongoing(args)
        elif args.stage == "iphone":
            _run_iphone(args)
        else:
            _run_ide(args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
