#!/usr/bin/env python3
"""API-friendly wrapper for generating Daily Tracking iPhone workbooks."""

from __future__ import annotations

import contextlib
import io
import time
from dataclasses import dataclass
from pathlib import Path

from generate_iphone_tracking import (
    DEFAULT_INPUT_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_REFERENCE_DIR,
    default_output_path,
    generate_iphone_tracking,
    output_path_for_directory,
    resolve_workbook_path,
)


@dataclass(frozen=True)
class IPhoneTrackingRequest:
    input_path: Path = Path(DEFAULT_INPUT_DIR)
    reference_path: Path = Path(DEFAULT_REFERENCE_DIR)
    output_path: Path = Path(DEFAULT_OUTPUT_DIR)


@dataclass(frozen=True)
class IPhoneTrackingResult:
    elapsed_seconds: float
    input_workbook: Path
    reference_workbook: Path
    output_workbook: Path
    stdout: str
    stderr: str


def resolve_output_workbook(output_path: Path, input_workbook: Path) -> Path:
    if output_path.suffix.lower() == ".xlsx":
        return output_path

    return output_path_for_directory(output_path, input_workbook)


def run_iphone_tracking(request: IPhoneTrackingRequest | None = None) -> IPhoneTrackingResult:
    request = request or IPhoneTrackingRequest()

    input_workbook = resolve_workbook_path(request.input_path).resolve()
    reference_workbook = resolve_workbook_path(request.reference_path).resolve()
    output_workbook = resolve_output_workbook(request.output_path, input_workbook).resolve()

    stdout = io.StringIO()
    stderr = io.StringIO()
    started = time.perf_counter()
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        generate_iphone_tracking(input_workbook, reference_workbook, output_workbook)

    elapsed_seconds = time.perf_counter() - started
    return IPhoneTrackingResult(
        elapsed_seconds=round(elapsed_seconds, 3),
        input_workbook=input_workbook,
        reference_workbook=reference_workbook,
        output_workbook=output_workbook,
        stdout=stdout.getvalue().strip(),
        stderr=stderr.getvalue().strip(),
    )


def run_iphone_tracking_from_paths(
    input_path: str | Path = DEFAULT_INPUT_DIR,
    reference_path: str | Path = DEFAULT_REFERENCE_DIR,
    output_path: str | Path = DEFAULT_OUTPUT_DIR,
) -> IPhoneTrackingResult:
    return run_iphone_tracking(
        IPhoneTrackingRequest(
            input_path=Path(input_path),
            reference_path=Path(reference_path),
            output_path=Path(output_path),
        )
    )


def output_filename_from_input_workbook(input_workbook: str | Path) -> str:
    return default_output_path(Path(input_workbook)).name
