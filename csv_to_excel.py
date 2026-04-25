#!/usr/bin/env python3
"""Parse and clean CSV files, then export them to Excel workbooks."""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable

import pandas as pd
from charset_normalizer import from_path


COMMON_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin1")
NULL_LIKE_VALUES = {"", "na", "n/a", "null", "none", "nan", "-"}
EXCEL_MAX_SHEET_NAME_LENGTH = 31
TARGET_HEADER_INSERT_AFTER = "Quo"
TARGET_SOURCE_HEADER = "Target + On Hold Duration"


@dataclass(frozen=True)
class ConvertOptions:
    output: Path
    delimiter: str | None
    encoding: str | None
    normalize_headers: bool
    keep_empty: bool
    drop_empty_columns: bool
    dedupe: bool
    infer_types: bool
    combine: bool


def detect_encoding(path: Path) -> str:
    """Detect file encoding with a conservative fallback list."""
    result = from_path(path).best()
    if result and result.encoding:
        return result.encoding

    for encoding in COMMON_ENCODINGS:
        try:
            path.read_text(encoding=encoding)
            return encoding
        except UnicodeDecodeError:
            continue

    return "latin1"


def detect_delimiter(path: Path, encoding: str) -> str:
    sample = path.read_text(encoding=encoding, errors="replace")[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        return dialect.delimiter
    except csv.Error:
        return ","


def normalize_column_name(value: object, index: int) -> str:
    raw = str(value).replace("\ufeff", "").replace("_", " ").strip()
    raw = re.sub(r"^#+", "", raw).strip()
    raw = re.sub(r"\s+", " ", raw)
    if not raw:
        return f"Column {index + 1}"

    words = []
    for word in raw.split(" "):
        if word.lower() == "otc":
            words.append("OTC")
        elif word.isupper() or re.fullmatch(r"[A-Z0-9#()/-]+", word):
            words.append(word)
        else:
            words.append(word[:1].upper() + word[1:])

    return " ".join(words)


def make_unique_columns(columns: Iterable[object], normalize: bool) -> list[str]:
    names: list[str] = []
    seen: dict[str, int] = {}

    for index, column in enumerate(columns):
        name = normalize_column_name(column, index) if normalize else str(column).strip()
        name = name or f"Column {index + 1}"

        count = seen.get(name, 0)
        seen[name] = count + 1
        names.append(name if count == 0 else f"{name} ({count + 1})")

    return names


def escape_excel_formula_text(value: object) -> object:
    if isinstance(value, str) and value.startswith("="):
        return f"'{value}"
    return value


def escape_excel_formulas(df: pd.DataFrame) -> pd.DataFrame:
    escaped = df.copy()
    escaped.columns = [escape_excel_formula_text(column) for column in escaped.columns]

    for column in escaped.columns:
        if pd.api.types.is_object_dtype(escaped[column]) or pd.api.types.is_string_dtype(escaped[column]):
            escaped[column] = escaped[column].map(escape_excel_formula_text, na_action="ignore")

    return escaped


def current_target_header() -> str:
    today = date.today()
    return f"TARGET  Detemined as 1 {today.strftime('%B %Y')}"


def current_target_values(source: pd.Series) -> pd.Series:
    today = date.today()
    month_name = today.strftime("%B")
    target_month = pd.Period(today, freq="M")
    source_months = pd.to_datetime(source, errors="coerce", format="mixed").dt.to_period("M")
    values = pd.Series("Target Not Yet Inputted", index=source.index, dtype="string")

    values[source_months < target_month] = f"Before {month_name}"
    values[source_months == target_month] = f"This {month_name}"
    values[source_months > target_month] = f"After {month_name}"

    return values


def add_target_header(df: pd.DataFrame) -> pd.DataFrame:
    if TARGET_HEADER_INSERT_AFTER not in df.columns or TARGET_SOURCE_HEADER not in df.columns:
        return df

    result = df.copy()
    target_header = current_target_header()
    if target_header in result.columns:
        result[target_header] = current_target_values(result[TARGET_SOURCE_HEADER])
        return result

    insert_at = result.columns.get_loc(TARGET_HEADER_INSERT_AFTER) + 1
    result.insert(insert_at, target_header, current_target_values(result[TARGET_SOURCE_HEADER]))
    return result


def clean_dataframe(df: pd.DataFrame, options: ConvertOptions) -> pd.DataFrame:
    df = df.copy()
    df.columns = make_unique_columns(df.columns, options.normalize_headers)
    df = add_target_header(df)

    for column in df.columns:
        if pd.api.types.is_object_dtype(df[column]) or pd.api.types.is_string_dtype(df[column]):
            cleaned = (
                df[column]
                .astype("string")
                .str.replace("\ufeff", "", regex=False)
                .str.strip()
                .str.replace(r"\s+", " ", regex=True)
            )
            df[column] = cleaned.mask(cleaned.str.lower().isin(NULL_LIKE_VALUES), pd.NA)

    if not options.keep_empty:
        df = df.dropna(how="all")

    if options.drop_empty_columns:
        df = df.dropna(axis="columns", how="all")

    if options.dedupe:
        df = df.drop_duplicates()

    if options.infer_types:
        df = infer_column_types(df)

    return escape_excel_formulas(df).reset_index(drop=True)


def infer_column_types(df: pd.DataFrame) -> pd.DataFrame:
    converted = df.copy()

    for column in converted.columns:
        series = converted[column]
        non_empty = series.dropna()
        if non_empty.empty:
            continue

        numeric_candidate = (
            non_empty.astype("string")
            .str.replace(",", "", regex=False)
            .str.replace("%", "", regex=False)
        )
        numeric = pd.to_numeric(numeric_candidate, errors="coerce")
        if numeric.notna().mean() >= 0.9:
            converted[column] = pd.to_numeric(
                series.astype("string").str.replace(",", "", regex=False).str.replace("%", "", regex=False),
                errors="coerce",
            )
            continue

        parsed_dates = pd.to_datetime(non_empty, errors="coerce", format="mixed")
        if parsed_dates.notna().mean() >= 0.9:
            converted[column] = pd.to_datetime(series, errors="coerce", format="mixed")

    return converted


def read_csv(path: Path, options: ConvertOptions) -> pd.DataFrame:
    encoding = options.encoding or detect_encoding(path)
    delimiter = options.delimiter or detect_delimiter(path, encoding)

    with path.open("r", encoding=encoding, errors="replace", newline="") as handle:
        rows = list(csv.reader(handle, delimiter=delimiter, skipinitialspace=True))

    first_data_row = next((index for index, row in enumerate(rows) if any(cell.strip() for cell in row)), None)
    if first_data_row is None:
        raise ValueError(f"CSV file is empty: {path}")

    header = rows[first_data_row]
    data_rows = rows[first_data_row + 1 :]
    width = max([len(header), *(len(row) for row in data_rows)] or [len(header)])

    if len(header) < width:
        header = [*header, *(f"Extra Column {index + 1}" for index in range(len(header), width))]

    normalized_rows = [
        [*row, *([""] * (width - len(row)))] if len(row) < width else row[:width]
        for row in data_rows
    ]

    return pd.DataFrame(normalized_rows, columns=header, dtype="string")


def sanitize_sheet_name(path: Path, used: set[str]) -> str:
    base = re.sub(r"[\[\]\:\*\?\/\\]", "_", path.stem).strip() or "Sheet"
    base = base[:EXCEL_MAX_SHEET_NAME_LENGTH]
    sheet_name = base
    suffix = 2

    while sheet_name in used:
        suffix_text = f"_{suffix}"
        sheet_name = f"{base[: EXCEL_MAX_SHEET_NAME_LENGTH - len(suffix_text)]}{suffix_text}"
        suffix += 1

    used.add(sheet_name)
    return sheet_name


def resolve_csv_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path]

    if not input_path.is_dir():
        raise FileNotFoundError(f"Input path does not exist: {input_path}")

    files = sorted(input_path.glob("*.csv"))
    if not files:
        raise FileNotFoundError(f"No CSV files found in: {input_path}")

    return files


def default_output_path(input_path: Path, combine: bool) -> Path:
    if input_path.is_file():
        return input_path.with_suffix(".xlsx")

    if combine:
        return input_path / "combined.xlsx"

    return input_path / "excel"


def resolve_output_path(input_path: Path, csv_files: list[Path], args: argparse.Namespace) -> Path:
    output = args.output or default_output_path(input_path, args.combine)

    if input_path.is_file() and (output.exists() and output.is_dir()):
        return output / f"{input_path.stem}.xlsx"

    if input_path.is_dir() and args.combine and output.suffix.lower() != ".xlsx":
        return output / "combined.xlsx"

    if input_path.is_dir() and not args.combine and output.suffix.lower() == ".xlsx":
        raise ValueError("Use --combine when writing multiple CSV files to one .xlsx output file.")

    if len(csv_files) == 1 and input_path.is_dir() and args.combine and output.suffix.lower() != ".xlsx":
        return output / "combined.xlsx"

    return output


def convert_one(csv_path: Path, output_path: Path, options: ConvertOptions) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df = clean_dataframe(read_csv(csv_path, options), options)
    df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"Wrote {output_path}")


def convert_many(csv_files: list[Path], options: ConvertOptions) -> None:
    if options.combine:
        options.output.parent.mkdir(parents=True, exist_ok=True)
        used_sheet_names: set[str] = set()
        with pd.ExcelWriter(options.output, engine="openpyxl") as writer:
            for csv_path in csv_files:
                df = clean_dataframe(read_csv(csv_path, options), options)
                df.to_excel(writer, sheet_name=sanitize_sheet_name(csv_path, used_sheet_names), index=False)
        print(f"Wrote {options.output}")
        return

    options.output.mkdir(parents=True, exist_ok=True)
    for csv_path in csv_files:
        convert_one(csv_path, options.output / f"{csv_path.stem}.xlsx", options)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Clean CSV data and export it to Excel (.xlsx).",
    )
    parser.add_argument("input", type=Path, help="CSV file or directory containing CSV files.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output .xlsx file or output directory. Defaults next to the input.",
    )
    parser.add_argument("--delimiter", help="CSV delimiter. Auto-detected when omitted.")
    parser.add_argument("--encoding", help="CSV encoding. Auto-detected when omitted.")
    parser.add_argument(
        "--keep-empty",
        action="store_true",
        help="Keep fully empty rows. Empty columns are kept by default.",
    )
    parser.add_argument(
        "--drop-empty-columns",
        action="store_true",
        help="Drop columns where all values are empty.",
    )
    parser.add_argument(
        "--no-normalize-headers",
        action="store_true",
        help="Keep original column names instead of converting underscores to readable headers.",
    )
    parser.add_argument(
        "--dedupe",
        action="store_true",
        help="Remove duplicate rows after cleaning.",
    )
    parser.add_argument(
        "--infer-types",
        action="store_true",
        help="Infer numeric and date columns. By default, values are preserved as text.",
    )
    parser.add_argument(
        "--combine",
        action="store_true",
        help="When input is a directory, write all CSV files to one workbook with one sheet per file.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    input_path = args.input.resolve()

    try:
        csv_files = resolve_csv_files(input_path)
        output_path = resolve_output_path(input_path, csv_files, args).resolve()
        options = ConvertOptions(
            output=output_path,
            delimiter=args.delimiter,
            encoding=args.encoding,
            normalize_headers=not args.no_normalize_headers,
            keep_empty=args.keep_empty,
            drop_empty_columns=args.drop_empty_columns,
            dedupe=args.dedupe,
            infer_types=args.infer_types,
            combine=args.combine,
        )

        if len(csv_files) == 1 and input_path.is_file():
            convert_one(csv_files[0], output_path, options)
        else:
            convert_many(csv_files, options)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
