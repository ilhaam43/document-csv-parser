"""Workbook differences that do not change the shared Master business rules."""

from dataclasses import dataclass

from .constants import MASTER_SERVICE_GROUPS, MOBILE_SERVICE_GROUPS, M2M_SERVICE_GROUPS, MOBILE_ONLY_REMOVED_PIVOTS


@dataclass(frozen=True)
class ReportProfile:
    name: str
    services: frozenset[str]
    weekly_window: bool = False
    summary_label_column: int = 19
    removed_pivots: frozenset[str] = frozenset()
    preserve_extra_columns: bool = False


PROFILES = {
    "master": ReportProfile("master", MASTER_SERVICE_GROUPS),
    "mobile": ReportProfile("mobile", MOBILE_SERVICE_GROUPS, True,
                            removed_pivots=MOBILE_ONLY_REMOVED_PIVOTS),
    "m2m": ReportProfile("m2m", M2M_SERVICE_GROUPS, True, 17, preserve_extra_columns=True),
}


def report_profile(report_type: str | None = None, *, mobile: bool = False) -> ReportProfile:
    name = report_type or ("mobile" if mobile else "master")
    if name not in PROFILES or (mobile and name != "mobile"):
        raise ValueError("report_type must be master, mobile, or m2m, without conflicting mobile mode.")
    return PROFILES[name]
