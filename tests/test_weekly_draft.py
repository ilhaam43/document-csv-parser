import csv
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

from weekly.finalize import finalize_weekly_data
from weekly.prepare import prepare_weekly_inputs


class WeeklyDraftTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        self.previous_rows = [
            {
                "No.": 102.5, "quote_num": "Q-1", "order number": None, "order_status": "#N/A",
                "order_status2": "In Progress",
                "order_status (Confirmed by MIDO Dept.)": "In Progress",
                "service_group": "Mobile", "fab_upload": datetime(2026, 9, 1),
                "otc": 0, "mrc": 0,
            },
            {
                "No.": 108.75, "quote_num": "Q-2", "order number": "O-2", "order_status": "Complete",
                "order_status2": "COMPLETE",
                "order_status (Confirmed by MIDO Dept.)": "COMPLETE",
                "service_group": "M2M", "fab_upload": datetime(2026, 9, 2),
                "otc": 0, "mrc": 0,
            },
        ]
        self.bq_rows = [
            {
                "quote_num": "Q-1", "order_number": "O-1", "order_status": "#N/A",
                "service_group": "Mobile", "FAB_Upload": "2026-09-01",
            },
            {
                "quote_num": "Q-2", "order_number": "O-2", "order_status": "Open",
                "service_group": "M2M", "FAB_Upload": "2026-09-02",
            },
            {
                "quote_num": "Q-3", "order_number": "O-3", "order_status": "#N/A",
                "service_group": "Mobile", "FAB_Upload": "2026-09-03",
            },
        ]

    def _inputs(self, module: str):
        return (
            patch(f"{module}.iter_master_rows", return_value=iter(self.previous_rows)),
            patch(f"{module}.iter_bq_rows", return_value=iter(self.bq_rows)),
            patch(f"{module}.load_holidays", return_value=set()),
        )

    @staticmethod
    def _read_csv(path: Path) -> list[dict[str, str]]:
        with path.open(encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))

    def test_review_queue_only_contains_unresolved_draft_rows(self) -> None:
        patches = self._inputs("weekly.prepare")
        with patches[0], patches[1], patches[2]:
            audit = prepare_weekly_inputs(
                bq_csv="bq.csv", previous_master="old.xlsx", holiday_workbook="holiday.xlsx",
                report_date=date(2026, 9, 24), output_dir=self.directory / "prepare",
            )
        self.assertEqual(audit["draft_master_rows"], 3)
        self.assertEqual(audit["new_rows_without_fab_seconds"], 1)
        self.assertEqual(audit["pending_catalyst_rows"], 2)
        draft = self._read_csv(self.directory / "prepare" / "draft-master-data.csv")
        self.assertEqual([row["No."] for row in draft], ["1", "2", "3"])
        self.assertEqual(draft[0]["order number"], "O-1")
        self.assertEqual(draft[0]["order_status"], "#N/A")
        self.assertEqual(draft[0]["order_status2"], "In Progress")
        self.assertEqual(draft[1]["order_status"], "Complete")
        with (self.directory / "prepare" / "catalyst-review.csv").open(encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            self.assertEqual(reader.fieldnames, [
                "master_row", "quote_num", "order_number", "ca_name",
                "service_group", "fab_upload", "current_status", "reviewed_status",
            ])
            review = list(reader)
        self.assertEqual([(row["master_row"], row["order_number"]) for row in review], [("2", "O-1"), ("4", "O-3")])

    def test_finalize_uses_draft_identity_and_leaves_blank_review_pending(self) -> None:
        patches = self._inputs("weekly.prepare")
        with patches[0], patches[1], patches[2]:
            prepare_weekly_inputs(
                bq_csv="bq.csv", previous_master="old.xlsx", holiday_workbook="holiday.xlsx",
                report_date=date(2026, 9, 24), output_dir=self.directory / "prepare",
            )
        review_path = self.directory / "prepare" / "catalyst-review.csv"
        with review_path.open(encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            fieldnames = reader.fieldnames
            rows = list(reader)
        rows[0]["reviewed_status"] = "Complete"
        with review_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        patches = self._inputs("weekly.finalize")
        with patches[0], patches[1], patches[2]:
            audit = finalize_weekly_data(
                bq_csv="bq.csv", previous_master="old.xlsx", catalyst_review=review_path,
                holiday_workbook="holiday.xlsx", report_date=date(2026, 9, 24),
                output_dir=self.directory / "final",
            )
        self.assertEqual(audit["catalyst_decisions"], 1)
        self.assertEqual(audit["new_rows_without_fab_seconds"], 1)
        self.assertEqual(audit["pending_catalyst_rows"], 1)
        decisions = self._read_csv(self.directory / "final" / "review-decisions.csv")
        self.assertEqual(decisions[0]["reviewed_status"], "Complete")
        self.assertEqual(decisions[0]["reviewed_by"], "")
        self.assertEqual(decisions[0]["review_reason"], "")
        final = self._read_csv(self.directory / "final" / "master-final-data.csv")
        self.assertEqual([row["No."] for row in final], ["1", "2", "3"])
        self.assertEqual(final[0]["order number"], "O-1")
        self.assertEqual(final[0]["order_status"], "Complete")
        self.assertEqual(final[2]["order_status"], "#N/A")

    def test_matched_empty_bq_lifecycle_is_resolved_before_catalyst_review(self) -> None:
        self.bq_rows[0]["order_status"] = ""
        patches = self._inputs("weekly.prepare")
        with patches[0], patches[1], patches[2]:
            audit = prepare_weekly_inputs(
                bq_csv="bq.csv", previous_master="old.xlsx", holiday_workbook="holiday.xlsx",
                report_date=date(2026, 9, 24), output_dir=self.directory / "prepare",
            )
        self.assertEqual(audit["pending_catalyst_rows"], 1)
        draft = self._read_csv(self.directory / "prepare" / "draft-master-data.csv")
        self.assertEqual(draft[0]["order_status"], "In Progress")
        self.assertEqual(draft[0]["order_status2"], "In Progress")
        self.assertEqual(draft[0]["order_status (Confirmed by MIDO Dept.)"], "In Progress")
        review = self._read_csv(self.directory / "prepare" / "catalyst-review.csv")
        self.assertEqual([row["quote_num"] for row in review], ["Q-3"])

    def test_prepare_preserves_timestamp_seconds_but_renumbers_source_no(self) -> None:
        self.bq_rows[2]["FAB_Upload"] = "2026-09-03 10:02:29"
        self.bq_rows[2]["No."] = "54881.7412032706"
        patches = self._inputs("weekly.prepare")
        with patches[0], patches[1], patches[2]:
            audit = prepare_weekly_inputs(
                bq_csv="bq.csv", previous_master="old.xlsx", holiday_workbook="holiday.xlsx",
                report_date=date(2026, 9, 24), output_dir=self.directory / "prepare",
            )
        self.assertEqual(audit["new_rows_without_fab_seconds"], 0)
        rows = self._read_csv(self.directory / "prepare" / "new-master-rows.csv")
        self.assertEqual(rows[0]["No."], "3")
        self.assertEqual(rows[0]["fab_upload"], "2026-09-03 10:02:29")

    def test_terminal_new_order_is_preserved_and_optional_override_is_audited(self) -> None:
        self.bq_rows[2]["order_status"] = "Open"
        self.bq_rows.append({
            "quote_num": "Q-3", "order_number": "O-4", "order_status": "Complete",
            "service_group": "Mobile", "FAB_Upload": "2026-09-03",
        })
        patches = self._inputs("weekly.prepare")
        with patches[0], patches[1], patches[2]:
            audit = prepare_weekly_inputs(
                bq_csv="bq.csv", previous_master="old.xlsx", holiday_workbook="holiday.xlsx",
                report_date=date(2026, 9, 24), output_dir=self.directory / "prepare",
            )
        self.assertEqual(audit["source_conflict_rows"], 1)
        patches = self._inputs("weekly.finalize")
        with patches[0], patches[1], patches[2]:
            final_audit = finalize_weekly_data(
                bq_csv="bq.csv", previous_master="old.xlsx",
                catalyst_review=self.directory / "prepare" / "catalyst-review.csv",
                holiday_workbook="holiday.xlsx", report_date=date(2026, 9, 24),
                output_dir=self.directory / "final",
            )
        self.assertEqual(final_audit["source_conflict_rows"], 1)
        self.assertEqual(final_audit["source_auto_order_rows"], 1)
        final = self._read_csv(self.directory / "final" / "master-final-data.csv")
        self.assertEqual(final[-1]["order_status"], "Complete")
        source_decisions = self._read_csv(self.directory / "final" / "bq-source-decisions.csv")
        self.assertEqual(source_decisions[0]["reviewed_source"], "order")
        self.assertEqual(source_decisions[0]["decision_method"], "automatic_terminal_order")
        source_review = self.directory / "prepare" / "source-review.csv"
        override = self._read_csv(source_review)
        override[0]["reviewed_source"] = "quote"
        with source_review.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=override[0].keys())
            writer.writeheader()
            writer.writerows(override)
        patches = self._inputs("weekly.finalize")
        with patches[0], patches[1], patches[2]:
            override_audit = finalize_weekly_data(
                bq_csv="bq.csv", previous_master="old.xlsx",
                catalyst_review=self.directory / "prepare" / "catalyst-review.csv",
                source_review=source_review,
                holiday_workbook="holiday.xlsx", report_date=date(2026, 9, 24),
                output_dir=self.directory / "final-override",
            )
        self.assertEqual(override_audit["source_quote_choices"], 1)
        overridden = self._read_csv(self.directory / "final-override" / "master-final-data.csv")
        self.assertEqual(overridden[-1]["order_status"], "Open")
        decision = self._read_csv(self.directory / "final-override" / "bq-source-decisions.csv")[0]
        self.assertEqual(decision["decision_method"], "explicit_override")

    def test_active_existing_and_new_rows_refresh_by_quote_without_changing_order_identity(self) -> None:
        self.previous_rows[0].update({"order number": "O-1", "order_status": "Open"})
        self.bq_rows.insert(0, {
            "quote_num": "Q-1", "order_number": "OTHER", "order_status": "Open",
            "service_group": "Mobile", "FAB_Upload": "2026-09-01",
        })
        self.bq_rows[1].update({"order_status": "Complete", "order_completion_date": "2026-09-18"})
        self.bq_rows[3].update({"order_status": "Complete", "order_completion_date": "2026-09-18"})
        self.bq_rows.append({
            "quote_num": "Q-3", "order_number": "O-4", "order_status": "Open",
            "service_group": "Mobile", "FAB_Upload": "2026-09-03",
            "cash_date": "2026-09-17 19:15:45 UTC",
        })
        patches = self._inputs("weekly.prepare")
        with patches[0], patches[1], patches[2]:
            prepare_weekly_inputs(
                bq_csv="bq.csv", previous_master="old.xlsx", holiday_workbook="holiday.xlsx",
                report_date=date(2026, 9, 24), output_dir=self.directory / "prepare",
            )
        draft = self._read_csv(self.directory / "prepare" / "draft-master-data.csv")
        self.assertEqual(draft[0]["order_status"], "Open")
        self.assertEqual(draft[0]["order number"], "O-1")
        self.assertEqual(draft[0]["order_completion_date"], "")
        self.assertEqual(draft[-1]["order_status"], "Complete")
        self.assertEqual(draft[-1]["order number"], "O-4")
        self.assertEqual(draft[-1]["cash_dt"], "2026-09-17 00:00:00")
        patches = self._inputs("weekly.finalize")
        with patches[0], patches[1], patches[2]:
            audit = finalize_weekly_data(
                bq_csv="bq.csv", previous_master="old.xlsx",
                catalyst_review=self.directory / "prepare" / "catalyst-review.csv",
                holiday_workbook="holiday.xlsx", report_date=date(2026, 9, 24),
                output_dir=self.directory / "final",
            )
        final = self._read_csv(self.directory / "final" / "master-final-data.csv")
        self.assertEqual(final, draft)
        self.assertEqual(audit["source_auto_quote_rows"], 2)
        methods = self._read_csv(self.directory / "final" / "bq-source-decisions.csv")
        self.assertEqual({r["decision_method"] for r in methods}, {"automatic_active_quote"})


if __name__ == "__main__":
    unittest.main()
