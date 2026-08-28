#!/usr/bin/env python3
"""Run the daily tracking, ongoing, and iPhone generators as one pipeline."""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

import csv_to_excel
import csv_to_excel_on_going
import generate_iphone_tracking
from excel_automation_lock import excel_process_lock
from pipeline_stage_runner import run_daily_stage, run_iphone_stage, run_ongoing_stage


def require_file(path: Path, label: str) -> Path:
    resolved = path.resolve()
    if not resolved.is_file():
        raise FileNotFoundError(f"{label} not found: {resolved}")
    return resolved


def newest_file(directory: Path, pattern: str, label: str) -> Path:
    resolved = directory.resolve()
    if not resolved.is_dir():
        raise FileNotFoundError(f"{label} directory not found: {resolved}")

    candidates = sorted(
        path for path in resolved.glob(pattern) if path.is_file() and not path.name.startswith("~$")
    )
    if not candidates:
        raise FileNotFoundError(f"No {label} file found in {resolved} with pattern {pattern}")
    return max(candidates, key=lambda path: path.stat().st_mtime).resolve()


def resolve_file_or_newest(path: Path, pattern: str, label: str) -> Path:
    return newest_file(path, pattern, label) if path.is_dir() else require_file(path, label)


def default_daily_output_path(output_dir: Path, data_order_csv: Path) -> Path:
    return output_dir / csv_to_excel.output_filename_from_csv_path(data_order_csv)


def default_ongoing_output_path(output_dir: Path, daily_tracking_output: Path) -> Path:
    return output_dir / csv_to_excel_on_going.output_filename_for_tracking(daily_tracking_output)


def default_iphone_output_path(output_dir: Path, daily_tracking_output: Path) -> Path:
    return output_dir / generate_iphone_tracking.default_output_path(daily_tracking_output).name


def convert_daily_tracking(
    data_order_csv: Path,
    daily_reference_workbook: Path,
    output_path: Path,
    skip_template_refresh: bool,
) -> Path:
    run_daily_stage(
        data_order_csv,
        daily_reference_workbook,
        output_path,
        refresh_template=not skip_template_refresh,
    )
    return output_path


def convert_ongoing_tracking(
    daily_tracking_workbook: Path,
    log_update_csv: Path,
    ongoing_reference_workbook: Path,
    output_path: Path,
    with_pivot: bool,
    keep_temp: bool,
    aging_date: date,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not keep_temp:
        run_ongoing_stage(
            daily_tracking_workbook,
            log_update_csv,
            ongoing_reference_workbook,
            output_path,
            with_pivot,
            aging_date,
        )
        return output_path

    # Debug mode retains the historical in-process temporary directory so its
    # intermediate workbook can be inspected after a failure.
    temp_dir = Path(tempfile.mkdtemp(prefix="daily_pipeline_ongoing_", dir=output_path.parent))
    try:
        validate_dir = temp_dir / "validate"
        validate_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ongoing_reference_workbook, validate_dir / ongoing_reference_workbook.name)

        temp_output_dir = temp_dir / "output-outgoing"
        temp_output = temp_output_dir / output_path.name
        csv_to_excel_on_going.apply_logic_to_workbook(
            daily_tracking_workbook,
            log_update_csv,
            temp_output,
            with_pivot=with_pivot,
            engine="com",
            clone_pivot_template=True,
            aging_date=aging_date,
            base_workbook_path=ongoing_reference_workbook,
        )
        shutil.copy2(temp_output, output_path)
    finally:
        print(f"[info] Kept ongoing temp directory: {temp_dir}")

    return output_path


def convert_iphone_tracking(
    daily_tracking_workbook: Path,
    iphone_reference_workbook: Path,
    output_path: Path,
) -> Path:
    run_iphone_stage(
        daily_tracking_workbook,
        iphone_reference_workbook,
        output_path,
    )
    return output_path


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate daily tracking, ongoing tracking, and iPhone tracking workbooks in one flow.",
    )
    parser.add_argument(
        "--data-order",
        type=Path,
        default=Path("input-today"),
        help="DataOrder CSV file or directory. Directory mode uses the newest *.csv file.",
    )
    parser.add_argument(
        "--daily-reference",
        type=Path,
        default=Path("vlookup-yesterday"),
        help="Daily Tracking yesterday workbook or directory. Directory mode uses the newest *.xlsx file.",
    )
    parser.add_argument(
        "--log-update",
        type=Path,
        help="LogUpdateStatusOrderSD CSV file or directory. Required unless --skip-ongoing is used.",
    )
    parser.add_argument(
        "--ongoing-reference",
        type=Path,
        help="Daily Tracking Ongoing yesterday workbook or directory. Required unless --skip-ongoing is used.",
    )
    parser.add_argument(
        "--iphone-reference",
        type=Path,
        default=Path("vlookup-iphone"),
        help="Daily Tracking iPhone yesterday workbook or directory. Required unless --skip-iphone is used.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=Path("output-pipeline"),
        help="Directory for all generated workbooks.",
    )
    parser.add_argument("--daily-output", type=Path, help="Optional explicit Daily Tracking output workbook path.")
    parser.add_argument("--ongoing-output", type=Path, help="Optional explicit Ongoing output workbook path.")
    parser.add_argument("--iphone-output", type=Path, help="Optional explicit iPhone output workbook path.")
    parser.add_argument("--skip-ongoing", action="store_true", help="Only skip the ongoing tracking step.")
    parser.add_argument("--skip-iphone", action="store_true", help="Only skip the iPhone tracking step.")
    parser.add_argument(
        "--skip-template-refresh",
        action="store_true",
        help="Skip Excel COM template refresh for the Daily Tracking step.",
    )
    parser.add_argument(
        "--ongoing-with-pivot",
        action="store_true",
        help="Run the slower ongoing ALL ORDER ON PROGRESS pivot refresh too.",
    )
    parser.add_argument(
        "--ongoing-aging-date",
        type=lambda value: datetime.strptime(value, "%Y-%m-%d").date(),
        default=None,
        help="Date used to calculate ongoing Pre-Installation Aging. Defaults to today's date.",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep temporary working folders for debugging.",
    )
    return parser.parse_args(argv)


def _main_locked(args: argparse.Namespace) -> int:
    try:
        data_order_csv = resolve_file_or_newest(args.data_order, "*.csv", "DataOrder")
        daily_reference = resolve_file_or_newest(args.daily_reference, "*.xlsx", "Daily Tracking reference")
        output_dir = args.output_dir.resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

        daily_output = (args.daily_output.resolve() if args.daily_output else default_daily_output_path(output_dir, data_order_csv))
        ongoing_output = args.ongoing_output.resolve() if args.ongoing_output else default_ongoing_output_path(output_dir, daily_output)
        iphone_output = args.iphone_output.resolve() if args.iphone_output else default_iphone_output_path(output_dir, daily_output)

        print("[1/3] Generating Daily Tracking...")
        daily_output = convert_daily_tracking(
            data_order_csv,
            daily_reference,
            daily_output,
            args.skip_template_refresh,
        )

        generated_outputs = [daily_output]

        if not args.skip_ongoing:
            if args.log_update is None:
                raise ValueError("--log-update is required unless --skip-ongoing is used.")
            if args.ongoing_reference is None:
                raise ValueError("--ongoing-reference is required unless --skip-ongoing is used.")

            log_update = resolve_file_or_newest(args.log_update, "LogUpdateStatusOrderSD-*.csv", "Log update")
            ongoing_reference = resolve_file_or_newest(
                args.ongoing_reference,
                "*.xlsx",
                "Ongoing reference",
            )
            print("[2/3] Generating Ongoing Tracking...")
            generated_outputs.append(
                convert_ongoing_tracking(
                    daily_output,
                    log_update,
                    ongoing_reference,
                    ongoing_output,
                    args.ongoing_with_pivot,
                    args.keep_temp,
                    args.ongoing_aging_date or date.today(),
                )
            )
        else:
            print("[2/3] Skipped Ongoing Tracking")

        if not args.skip_iphone:
            iphone_reference = resolve_file_or_newest(args.iphone_reference, "*.xlsx", "iPhone reference")
            print("[3/3] Generating iPhone Tracking...")
            generated_outputs.append(convert_iphone_tracking(daily_output, iphone_reference, iphone_output))
        else:
            print("[3/3] Skipped iPhone Tracking")

        print("Generated outputs:")
        for output in generated_outputs:
            print(f"- {output}")
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    return 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        with excel_process_lock(60 * 60):
            return _main_locked(args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
