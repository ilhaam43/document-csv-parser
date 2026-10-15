"""Typed values shared by weekly pipeline stages."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any


@dataclass(frozen=True)
class RecordKey:
    quote_num: str
    order_number: str


@dataclass(frozen=True)
class CatalystOverride:
    key: RecordKey
    order_status: str


@dataclass(frozen=True)
class ReviewDecision:
    order_status: str | None = None
    completion_date: datetime | None = None
    leadtime_fab_cash: int | None = None
    pm_approved: bool | None = None
    reviewed_by: str = ""
    review_reason: str = ""


@dataclass(frozen=True)
class CheckpointRow:
    fab_upload: datetime
    leadtime_fab_cash: int | None


@dataclass(frozen=True)
class ReviewLoadResult:
    decisions: dict[int, ReviewDecision]
    count: int
    checkpoint_rows: dict[int, CheckpointRow] = field(default_factory=dict)
    unmatched_new_rows: int = 0


@dataclass(frozen=True)
class ValidationIssue:
    code: str
    message: str
    key: RecordKey | None = None
    row_number: int | None = None


@dataclass
class SourceSelectionResult:
    selected_rows: list[dict[str, Any]] = field(default_factory=list)
    rejected_reasons: Counter[str] = field(default_factory=Counter)
    issues: list[ValidationIssue] = field(default_factory=list)


@dataclass(frozen=True)
class MatchResult:
    row_number: int | None
    strategy: str
    issue: ValidationIssue | None = None


@dataclass
class WeeklyAudit:
    checks: dict[str, Any] = field(default_factory=dict)
    failures: list[str] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return not self.failures

    def check(self, name: str, actual: Any, expected: Any) -> None:
        self.checks[name] = {"actual": actual, "expected": expected}
        if actual != expected:
            self.failures.append(f"{name}: expected {expected!r}, got {actual!r}")

    def to_dict(self) -> dict[str, Any]:
        return {
            "passed": self.passed,
            "checks": self.checks,
            "failures": list(self.failures),
        }
