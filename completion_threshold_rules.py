"""Shared weekly completion-indicator business rules."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date


STANDARD_PROFILE = "standard"
ONGOING_PROFILE = "ongoing"
WEEKLY_MODE = "weekly"
TARGET_AFTER_MODE = "target_after"


@dataclass(frozen=True)
class CompletionThresholds:
    week: int
    green_percent: int
    yellow_percent: int
    red_percent: int


_THRESHOLD_MATRIX: dict[str, dict[str, dict[int, tuple[int, int]]]] = {
    STANDARD_PROFILE: {
        WEEKLY_MODE: {
            1: (30, 20),
            2: (40, 30),
            3: (50, 40),
            4: (60, 50),
            5: (60, 50),
        },
        TARGET_AFTER_MODE: {
            1: (10, 5),
            2: (20, 10),
            3: (30, 20),
            4: (40, 30),
            5: (50, 40),
        },
    },
    ONGOING_PROFILE: {
        WEEKLY_MODE: {
            1: (30, 20),
            2: (40, 30),
            3: (50, 40),
            4: (99, 98),
            5: (99, 98),
        },
        TARGET_AFTER_MODE: {
            1: (10, 5),
            2: (20, 10),
            3: (30, 20),
            4: (99, 98),
            5: (99, 98),
        },
    },
}


def report_week(report_date: date) -> int:
    """Return the stakeholder week bucket for a monthly report date."""
    day = report_date.day
    if day <= 6:
        return 1
    if day <= 13:
        return 2
    if day <= 20:
        return 3
    if day <= 28:
        return 4
    return 5


def completion_thresholds_for(
    report_date: date,
    *,
    profile: str = STANDARD_PROFILE,
    mode: str = WEEKLY_MODE,
) -> CompletionThresholds:
    """Resolve the upper/lower traffic-light thresholds for one report block."""
    try:
        profile_matrix = _THRESHOLD_MATRIX[profile]
    except KeyError as exc:
        raise ValueError(f"Unknown completion threshold profile: {profile}") from exc
    try:
        mode_matrix = profile_matrix[mode]
    except KeyError as exc:
        raise ValueError(f"Unknown completion threshold mode: {mode}") from exc

    week = report_week(report_date)
    green_percent, lower_percent = mode_matrix[week]
    return CompletionThresholds(
        week=week,
        green_percent=green_percent,
        yellow_percent=lower_percent,
        red_percent=lower_percent,
    )


def completion_legend_values_for(
    report_date: date,
    *,
    profile: str = STANDARD_PROFILE,
    mode: str = WEEKLY_MODE,
) -> tuple[str, str, str]:
    thresholds = completion_thresholds_for(
        report_date,
        profile=profile,
        mode=mode,
    )
    return (
        f"Green : >{thresholds.green_percent}%",
        f"Yellow : >={thresholds.yellow_percent}%",
        f"Red : <{thresholds.red_percent}%",
    )
