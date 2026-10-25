"""Read M2M pivot metrics and update the existing editable M2M presentation."""

import argparse
import json
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from .constants import ACHIEVEMENT_MAX_DAYS
from .dates import parse_report_month
from .presentation import ChartData, _metric_value, _month_key, _pivot_rows, _rounded, _section


EXPECTED_SERIES = {
    1: ("LEADTIME FAB-CASH", "LEADTIME CASH-COMPLETE"),
    2: ("INCOMING FAB", "MRC+OTC"),
    3: ("M2M SERVICE",),
    4: ("COMPLETION ORDER", "MRC+OTC"),
}
NUMBER_FORMATS = {1: ("#,##0", "#,##0"), 2: ("#,##0", "#,##0.00"),
                  3: ("#,##0",), 4: ("#,##0", "#,##0.00")}


def extract_m2m_charts(workbook: str | Path, report_date: date, report_week: int):
    if not 1 <= report_week <= 53:
        raise ValueError("Report week must be between 1 and 53.")
    source = load_workbook(workbook, read_only=True, data_only=True)
    try:
        if "PIVOT PERFORMANCE" not in source:
            raise ValueError("M2M workbook has no PIVOT PERFORMANCE sheet.")
        sheet = source["PIVOT PERFORMANCE"]
        revenue = _section(sheet, 1, "INCOMING FAB (MRC)")
        quantity = _section(sheet, 1, "INCOMING FAB (QTY)")
        fab = _section(sheet, 1, "LEAD TIME FAB - CASH")
        cash = _section(sheet, 1, f"M2M SERVICE (TARGET = {ACHIEVEMENT_MAX_DAYS['M2M']} DAYS)")
        monthly_cash = _section(sheet, 10, "LEAD TIME CASH-COMPLETE")
        rev_month, _ = _pivot_rows(sheet, revenue, 1, (2,))
        qty_month, _ = _pivot_rows(sheet, quantity, 1, (2,))
        cash_month, cash_total = _pivot_rows(sheet, monthly_cash, 10, (11,))
        rev_week, _ = _pivot_rows(sheet, revenue, 10, (11,))
        qty_week, _ = _pivot_rows(sheet, quantity, 10, (11,))
        fab_week, fab_total = _pivot_rows(sheet, fab, 10, (11,))
        cash_week, cash_week_total = _pivot_rows(sheet, cash, 1, (2,))
        labels = tuple(f"W{week}" for week in range(max(1, report_week - 3), report_week + 1))
        months = sorted({key for rows in (qty_month, rev_month, cash_month) for label in rows
                         if (key := _month_key(label)) and key[0] == report_date.year
                         and key[1] <= report_date.month} | {(report_date.year, report_date.month)})
        # Derive the year from the selected period, not stale labels in the PPT.
        month_labels = tuple(date(year, month, 1).strftime("%b-%y").upper() for year, month in months)

        def values(rows, keys, *, missing=None):
            return tuple(_metric_value(rows[key][0]) if key in rows and rows[key][0] is not None
                         else missing for key in keys)

        def monthly(rows, *, missing=None):
            return values({_month_key(label): value for label, value in rows.items()}, months, missing=missing)

        def average(total):
            return _rounded(total[0]) if total and total[0] is not None else None

        # No orders means zero count/revenue, but an undefined mean stays a chart gap.
        return {
            1: ChartData(labels, values(fab_week, labels), values(cash_week, labels)),
            2: ChartData(labels, values(qty_week, labels, missing=0), values(rev_week, labels, missing=0)),
            3: ChartData(month_labels, monthly(cash_month), ()),
            4: ChartData(month_labels, monthly(qty_month, missing=0), monthly(rev_month, missing=0)),
        }, {"m2m_fab": average(fab_total), "m2m_cash": average(cash_week_total),
            "monthly_m2m_cash": average(cash_total)}
    finally:
        source.close()


def write_m2m_presentation(*, m2m_workbook: str | Path, template_pptx: str | Path,
                           output_pptx: str | Path, report_date: date, report_week: int,
                           report_month: date | None = None) -> Path:
    from .presentation import write_weekly_presentation
    return write_weekly_presentation(mobile_workbook=m2m_workbook, template_pptx=template_pptx,
        output_pptx=output_pptx, report_date=report_date, report_week=report_week,
        report_month=report_month, report_type="m2m")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--m2m-workbook", required=True, type=Path)
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--report-date", required=True, type=date.fromisoformat)
    parser.add_argument("--report-month", type=parse_report_month)
    parser.add_argument("--report-week", required=True, type=int)
    args = parser.parse_args()
    path = write_m2m_presentation(
        m2m_workbook=args.m2m_workbook, template_pptx=args.template, output_pptx=args.output,
        report_date=args.report_date, report_month=args.report_month, report_week=args.report_week,
    )
    print(json.dumps({"output": str(path.resolve())}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
