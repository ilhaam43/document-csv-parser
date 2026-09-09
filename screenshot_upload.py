"""Publish report screenshots through the server's public HTTPS domain."""

from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path


DEFAULT_SCREENSHOT_BASE_URL = "https://auto-report-secm.lintasarta.dev/screenshots"
REPORT_FILE_OFFSETS = {1: 1, 2: 5, 3: 10, 4: 11}


def screenshot_date_folder(report_date: date | None = None) -> str:
    return (report_date or date.today()).strftime(os.getenv("SCREENSHOT_DATE_FORMAT", "%Y-%m-%d"))


def report_screenshot_targets(
    app_root: Path,
    report_number: int,
    image_count: int,
    report_date: date | None = None,
    generated_at: datetime | None = None,
) -> list[tuple[Path, str]]:
    if report_number not in REPORT_FILE_OFFSETS:
        raise ValueError(f"Unsupported report number: {report_number}")
    if image_count < 1:
        raise ValueError("image_count must be at least 1")

    root = Path(os.getenv("SCREENSHOT_ROOT", str(app_root / "screenshots"))).resolve()
    capture_time = generated_at or datetime.now()
    date_folder = screenshot_date_folder(report_date or capture_time.date())
    timestamp_folder = capture_time.strftime(
        os.getenv("SCREENSHOT_TIMESTAMP_FORMAT", "%H%M%S%f")
    )
    output_dir = root / date_folder / timestamp_folder
    output_dir.mkdir(parents=True, exist_ok=True)

    base_url = os.getenv("SCREENSHOT_BASE_URL", DEFAULT_SCREENSHOT_BASE_URL).rstrip("/")
    first_file = REPORT_FILE_OFFSETS[report_number]
    return [
        (
            output_dir / f"file{first_file + index}.jpg",
            f"{base_url}/{date_folder}/{timestamp_folder}/file{first_file + index}.jpg",
        )
        for index in range(image_count)
    ]


def replace_inline_image_sources(message: str, content_ids: list[str], public_urls: list[str]) -> str:
    if len(content_ids) != len(public_urls):
        raise ValueError("Each inline image Content-ID must have one public URL")
    for content_id, public_url in zip(content_ids, public_urls):
        message = message.replace(f"cid:{content_id}", public_url)
    return message
