"""Stable schemas and business constants for the weekly reports."""

from __future__ import annotations


MASTER_SHEET_NAME = "ALL DATA"
MASTER_TABLE_NAME = "Table1"
PIVOT_PERFORMANCE_SHEET_NAME = "PIVOT PERFORMANCE"
MOBILE_ONLY_REMOVED_PIVOTS = frozenset(
    {"PivotTable25", "PivotTable18", "PivotTable8", "PivotTable6", "PivotTable13"}
)

DATE_COLUMNS = frozenset(
    {
        "fab_upload",
        "cash_dt",
        "order created",
        "order_completion_date",
        "ACTIVATION DATE",
    }
)

INTEGER_COLUMNS = frozenset(
    {
        "Count FAB",
        "Count Cash",
        "Count Completion",
        "LEADTIME FAB-CASH",
        "LEADTIME CASH-COMPLETE",
    }
)

DECIMAL_COLUMNS = frozenset({"No.", "otc", "mrc", "OTC+MRC", "OTC+MRC (inMio)"})

MASTER_COLUMNS: tuple[str, ...] = (
    "No.",
    "Month of FAB Upload",
    "Week FAB",
    "YEAR FAB Upload",
    "quote_num",
    "ca_name",
    "service_group",
    "order_status",
    "order_status2",
    "order_status (Confirmed by MIDO Dept.)",
    "fab_upload",
    "Count FAB",
    "cash_dt",
    "Count Cash",
    "order number",
    "order created",
    "order_completion_date",
    "Count Completion",
    "Week Completion",
    "Month Completion",
    "Year Completion",
    "otc",
    "mrc",
    "OTC+MRC",
    "OTC+MRC (inMio)",
    "LEADTIME FAB-CASH",
    "LEADTIME CASH-COMPLETE",
    "ACHIEVEMENT",
    "ADJUSTMENT",
    "ACTIVATION DATE",
    "SALES NAME",
    "DEPT SALES",
    "DIVISION SALES",
    "GROUP SALES",
)

SOURCE_COLUMNS: tuple[str, ...] = (
    "quote_num",
    "ca_name",
    "service_group",
    "order_number",
    "order_status",
    "quote_crea",
    "FAB_Upload",
    "mld_dt",
    "cash_date",
    "order_crea",
    "order_completion_date",
    "sales_name",
    "otc",
    "mrc",
    "order_type",
    "remark",
    "sourcefile",
    "sales_stage",
    "ca_number",
    "addr",
    "addr_b",
    "addr_line_2",
    "addr_line_2_b",
    "building",
    "building_b",
    "city",
    "city_b",
    "sub_service",
    "quote_name",
    "product",
    "dept",
    "div",
    "grp",
)

MASTER_SERVICE_GROUPS = frozenset({"Bulk Service", "Mobile", "M2M"})
MOBILE_SERVICE_GROUPS = frozenset({"Bulk Service", "Mobile"})
M2M_SERVICE_GROUPS = frozenset({"M2M"})

SERVICE_ALIASES: dict[str, str] = {
    "bulk service": "Bulk Service",
    "bulkservice": "Bulk Service",
    "m2m": "M2M",
    "mobile": "Mobile",
}

ORDER_STATUS_ALIASES: dict[str, str] = {
    "cancel": "CANCEL",
    "cancelled": "CANCEL",
    "canceled": "CANCEL",
    "complete": "COMPLETE",
    "completed": "COMPLETE",
    "in progress": "In Progress",
    "open": "In Progress",
}

ACHIEVEMENT_MAX_DAYS: dict[str, int] = {
    "Mobile": 3,
    "Bulk Service": 5,
    "M2M": 19,
}
