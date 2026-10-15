import unittest
from datetime import date, datetime
from decimal import Decimal

from weekly.rules import achievement_for, build_master_record, inclusive_networkdays


class WeeklyRulesTests(unittest.TestCase):
    def test_raw_bq_lifecycle_dates_are_date_only_but_fab_keeps_timestamp(self) -> None:
        result = build_master_record({
            "FAB_Upload": "2026-10-01 15:40:42 UTC",
            "cash_date": "2026-10-02 19:15:30 UTC",
            "order_crea": "2026-10-03 10:02:13 UTC",
            "order_completion_date": "2026-10-05 22:10:12 UTC",
            "service_group": "Mobile", "order_status": "Complete",
        }, row_number=1)
        self.assertEqual(result["fab_upload"], datetime(2026, 10, 1, 15, 40, 42))
        for field, day in (("cash_dt", 2), ("order created", 3), ("order_completion_date", 5)):
            self.assertEqual(result[field], datetime(2026, 10, day))

    def test_zero_rules_are_column_specific_not_a_global_replacement(self) -> None:
        result = build_master_record({
            "FAB_Upload": "2026-09-20 09:00:00",
            "cash_date": "2026-09-20", "order_completion_date": "2026-09-20",
            "service_group": "Mobile", "order_status": "Complete",
        }, row_number=1)
        self.assertEqual(result["LEADTIME FAB-CASH"], 0)
        self.assertEqual(result["LEADTIME CASH-COMPLETE"], 1)

    def test_networkdays_is_inclusive_and_excludes_holidays(self) -> None:
        self.assertEqual(
            inclusive_networkdays(
                date(2026, 9, 17),
                date(2026, 9, 21),
                {date(2026, 9, 18)},
            ),
            2,
        )

    def test_networkdays_keeps_manual_zero_and_supports_explicit_normalization(self) -> None:
        self.assertEqual(inclusive_networkdays(date(2026, 9, 19), date(2026, 9, 20)), 0)
        self.assertEqual(inclusive_networkdays(
            date(2026, 9, 19), date(2026, 9, 20), normalize_zero=True
        ), 1)

    def test_networkdays_keeps_negative_values_visible(self) -> None:
        self.assertEqual(inclusive_networkdays(date(2026, 9, 21), date(2026, 9, 18)), -2)

    def test_achievement_rules_and_explicit_pm_approval(self) -> None:
        self.assertEqual(
            achievement_for(
                service_group="Mobile",
                completion_date=date(2026, 9, 21),
                leadtime_cash_complete=3,
            ),
            "ACHIEVED",
        )
        self.assertIsNone(
            achievement_for(
                service_group="Mobile",
                completion_date=date(2026, 9, 21),
                leadtime_cash_complete=-1,
            )
        )
        self.assertEqual(
            achievement_for(
                service_group="Bulk Service",
                completion_date=date(2026, 9, 21),
                leadtime_cash_complete=6,
            ),
            "NOT ACHIEVED",
        )
        self.assertEqual(
            achievement_for(
                service_group="M2M",
                completion_date=date(2026, 9, 21),
                leadtime_cash_complete=18,
            ),
            "ACHIEVED",
        )
        self.assertEqual(
            achievement_for(
                service_group="M2M",
                completion_date=date(2026, 9, 21),
                leadtime_cash_complete=19,
            ),
            "ACHIEVED",
        )
        self.assertEqual(
            achievement_for(
                service_group="M2M",
                completion_date=date(2026, 9, 21),
                leadtime_cash_complete=20,
            ),
            "NOT ACHIEVED",
        )
        self.assertEqual(
            achievement_for(
                service_group="Mobile",
                completion_date=date(2026, 9, 21),
                leadtime_cash_complete=99,
                pm_approved=True,
            ),
            "ACHIEVED",
        )
        self.assertEqual(
            achievement_for(
                service_group="Mobile",
                completion_date=date(2026, 9, 21),
                leadtime_cash_complete=99,
            ),
            "NOT ACHIEVED",
        )

    def test_build_master_record_maps_reviewed_fields(self) -> None:
        source = {
            "quote_num": "Q-1",
            "ca_name": "Customer",
            "service_group": "mobile",
            "order_number": "O-1",
            "order_status": "Complete",
            "FAB_Upload": "9/17/2026 10:30",
            "cash_date": "9/18/2026",
            "order_crea": "9/18/2026",
            "order_completion_date": "9/21/2026",
            "mld_dt": "",
            "sales_name": "Sales",
            "otc": "1000",
            "mrc": "2000",
            "dept": "must not be copied",
            "div": "must not be copied",
            "grp": "must not be copied",
        }
        record = build_master_record(source, row_number=1)
        self.assertEqual(record["Week FAB"], "W38")
        self.assertEqual(record["order_status2"], "COMPLETE")
        self.assertEqual(record["OTC+MRC"], Decimal("3000"))
        self.assertEqual(record["OTC+MRC (inMio)"], Decimal("0.003"))
        self.assertEqual(record["LEADTIME FAB-CASH"], 2)
        self.assertEqual(record["LEADTIME CASH-COMPLETE"], 2)
        self.assertEqual(record["ACHIEVEMENT"], "ACHIEVED")
        self.assertIsNone(record["DEPT SALES"])

    def test_new_record_preserves_source_name_spacing(self) -> None:
        record = build_master_record(
            {
                "quote_num": "Q-1",
                "ca_name": "PT. AT&T  GLOBAL NETWORK SERVICES INDONE",
                "sales_name": "R  NINDYA MUSTIKA",
                "service_group": "Mobile",
                "FAB_Upload": "9/20/2026 22:43",
            },
            row_number=2,
        )
        self.assertEqual(record["ca_name"], "PT. AT&T  GLOBAL NETWORK SERVICES INDONE")
        self.assertEqual(record["SALES NAME"], "R  NINDYA MUSTIKA")

    def test_upstream_precision_is_preserved_but_number_is_sequential(self) -> None:
        record = build_master_record(
            {
                "quote_num": "Q-1", "service_group": "Mobile",
                "FAB_Upload": "9/30/2026 10:02:29", "No.": "54881.7412032706",
            },
            row_number=53593,
        )
        self.assertEqual(record["No."], 53593)
        self.assertEqual(record["fab_upload"], datetime(2026, 9, 30, 10, 2, 29))


if __name__ == "__main__":
    unittest.main()
