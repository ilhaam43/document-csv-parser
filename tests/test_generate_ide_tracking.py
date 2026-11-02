import unittest
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter

from generate_ide_tracking import (
    ACTUAL_RFS_DATE_HEADER,
    DEPT_HEADER,
    NEW_RFS_INITIAL_HEADER,
    QUOTE_ID_HEADER,
    TargetPivotPageFilter,
    build_all_order,
    apply_date_formats,
    apply_month_complete_formula,
    apply_na_errors,
    apply_previous_field_fallbacks,
    apply_text_formats,
    completion_thresholds,
    date_from_filename,
    financial_value_or_zero,
    capture_delivery_pivot_filters,
    configure_delivery_pivot_filters,
    delivery_pivot_policy,
    has_expected_all_order_freeze_panes,
    hidden_row_ranges_between_sections,
    identify_target_complete_pivots,
    is_missing,
    load_collabs_lookup,
    load_source_new_rfs_lookup,
    normalize_dashboard_text_dates,
    parse_dashboard_field_datetime,
    normalize_pivot_section_row_visibility,
    parse_excel_datetime,
    planned_rfs_with_completion_year,
    preserve_pivot_refresh_layout,
    read_raw_dataframe,
    read_previous_dataframe,
    recover_source_quote_date_years,
    reconcile_date_column,
    reconciled_date_value,
    resolve_financial_values,
    restore_target_pivot_filters,
    set_pivot_page_field_position,
    target_period,
    target_value,
    target_value_with_previous,
    validate_output,
    validate_raw_headers,
)


class IdeDynamicColumnsTests(unittest.TestCase):
    RAW_HEADERS = (
        "NO", "QUOTE ID", "PROJECT NAME", "QUOTE NAME", "QUOTE NAME 2",
        "CUSTOMER NAME", "PROJECT TYPE", "BRAND", "PARTNERS",
        "STATUS Catalog / non Catalog", "NOMOR PO", "SO", "CID", "QUOTE DATE",
        "FAB SIGN DATE", "FAB UPLOAD DATE", "NEW RFS INITIAL", "CASH DATE",
        "SO DATE", "SERVICE READY", "ACTUAL RFS DATE", "AOS SENT DATE",
        "AOS SIGN DATE", "SO COMPLETION DATE", "STATUS DELIVERY", "SD Note",
        "SALES", "SEGMENT SALES", "PRODUCT2", "ADMIN", "APM", "PM", "DEPT.",
        "PIC CUSTOMER", "CUSTOMER ADDRESS", "Project Category", "REVENUE",
        "REVENUE UPSIDE", "PENALTY BC", "CoS BC", "RISK BC", "PENALTY ACTUAL",
        "COS ACTUAL (PO)", "RISK ACTUAL", "COST PLANNED", "ACTUAL COST",
        "COST EFFICIENCY INDEX (CEI)", "STATUS HOTO D2", "OTD",
    )

    def setUp(self) -> None:
        raw_row = dict.fromkeys(self.RAW_HEADERS)
        raw_row.update({
            "NO": 1, "QUOTE ID": "Q-1", "QUOTE NAME": "Example",
            "FAB UPLOAD DATE": datetime(2026, 10, 1),
            "NEW RFS INITIAL": datetime(2026, 10, 12),
            "ACTUAL RFS DATE": datetime(2026, 10, 12),
            "STATUS DELIVERY": "SO Complete",
        })
        self.raw = pd.DataFrame([raw_row])
        self.previous = pd.DataFrame([{
            "QUOTE ID": "Q-1", "ORDER TYPE": "New Registration",
            "OTC": 100, "MRC": 200, "STATUS DELIVERY": "SO Complete",
            "Pre-Installation Start Date": datetime(2026, 10, 1),
        }])

    def _build(self, raw, *, report_date=date(2026, 10, 12), previous_path=Path("Daily Tracking IDE 9 October 2026.xlsx")):
        with patch("generate_ide_tracking.load_collabs_lookup", return_value={}):
            return build_all_order(
                raw, self.previous, Path("collabs.csv"), report_date, previous_path,
            ).dataframe

    def test_raw_additions_keep_their_values_and_existing_business_columns(self) -> None:
        baseline = self._build(self.raw)
        self.assertEqual(len(baseline.columns), 62)
        self.assertEqual(baseline.columns.get_loc("MONTH COMPLETE"), 32)
        for position in (6, len(self.raw.columns)):
            with self.subTest(position=position):
                extended = self.raw.copy()
                extended.insert(position, "MONTH COMPLETE", ["Okt"])
                extended["SOURCE NOTES"] = ["Confirmed"]
                result = self._build(extended)
                self.assertEqual(len(result.columns), 63)
                self.assertTrue(result["MONTH COMPLETE"].isna().all())
                self.assertEqual(result["SOURCE NOTES"].tolist(), ["Confirmed"])
                self.assertEqual(result.columns.get_loc("MONTH COMPLETE"), 32)
                pd.testing.assert_frame_equal(
                    result.drop(columns=["MONTH COMPLETE", "SOURCE NOTES"]),
                    baseline.drop(columns=["MONTH COMPLETE"]),
                )

    def test_missing_month_has_one_calculated_column_after_completion_date(self) -> None:
        with patch("generate_ide_tracking.load_collabs_lookup", return_value={}):
            build = build_all_order(
                self.raw, self.previous, Path("collabs.csv"), date(2026, 10, 13),
                Path("Daily Tracking IDE 9 October 2026.xlsx"),
            )
        self.assertEqual(build.dataframe.columns.tolist().count("MONTH COMPLETE"), 1)
        self.assertEqual(build.dataframe.columns.get_loc("MONTH COMPLETE"), 32)
        self.assertTrue(build.dataframe["MONTH COMPLETE"].isna().all())

    def test_source_month_is_discarded_and_both_dashboard_formats_build_the_same_data(self) -> None:
        raw = self.raw.copy()
        raw["MONTH COMPLETE"] = ["Okt"]
        result = self._build(raw)
        pd.testing.assert_frame_equal(result, self._build(self.raw))
        self.assertEqual(result.columns.tolist().count("MONTH COMPLETE"), 1)

    def test_new_period_keeps_current_dates_per_duplicate_row(self) -> None:
        raw = pd.concat([self.raw, self.raw], ignore_index=True)
        raw[NEW_RFS_INITIAL_HEADER] = [datetime(2026, 4, 28)] * 2
        raw[ACTUAL_RFS_DATE_HEADER] = [datetime(2026, 3, 12), "12/03/2026"]
        result = self._build(raw, report_date=date(2026, 11, 2), previous_path=Path("Daily Tracking IDE 30 October 2026.xlsx"))
        self.assertEqual(result[ACTUAL_RFS_DATE_HEADER].tolist(), [
            datetime(2026, 3, 12), datetime(2026, 12, 3),
        ])

    def test_continuation_actual_rfs_uses_first_quote_match_but_preserves_blank(self) -> None:
        raw = pd.concat([self.raw] * 3, ignore_index=True)
        raw[ACTUAL_RFS_DATE_HEADER] = [datetime(2026, 4, 15), datetime(2026, 4, 18), None]
        result = self._build(raw)
        self.assertEqual(result[ACTUAL_RFS_DATE_HEADER].tolist()[:2], [datetime(2026, 4, 15)] * 2)
        self.assertTrue(is_missing(result.at[2, ACTUAL_RFS_DATE_HEADER]))

    def test_year_carry_distinguishes_text_from_typed_dates_in_new_period(self) -> None:
        for value in (datetime(2025, 4, 15), "15/04/2025"):
            raw = self.raw.copy()
            raw[NEW_RFS_INITIAL_HEADER] = [value]
            self.previous[NEW_RFS_INITIAL_HEADER] = [datetime(2026, 4, 15)]
            continued = self._build(raw)
            refreshed = self._build(raw, report_date=date(2026, 11, 2),
                                    previous_path=Path("Daily Tracking IDE 30 October 2026.xlsx"))
            self.assertEqual(continued.at[0, NEW_RFS_INITIAL_HEADER], datetime(2026, 4, 15))
            expected = datetime(2026 if isinstance(value, str) else 2025, 4, 15)
            self.assertEqual(refreshed.at[0, NEW_RFS_INITIAL_HEADER], expected)

    def test_planned_year_rule_only_runs_on_period_refresh_and_records_evidence(self) -> None:
        raw = self.raw.copy()
        raw[NEW_RFS_INITIAL_HEADER] = ["20/06/2025"]
        raw["SO COMPLETION DATE"] = [datetime(2026, 6, 26)]
        for report, previous, expected in (
            (date(2026, 10, 13), Path("Daily Tracking IDE 12 October 2026.xlsx"), 2025),
            (date(2026, 11, 2), Path("Daily Tracking IDE 30 October 2026.xlsx"), 2026),
        ):
            built = build_all_order(raw, self.previous, Path("unused.csv"), report, previous)
            self.assertEqual(built.dataframe.at[0, NEW_RFS_INITIAL_HEADER], datetime(expected, 6, 20))
            self.assertEqual(len(built.planned_rfs_year_rollovers), int(expected == 2026))
            if expected == 2026:
                self.assertEqual(built.planned_rfs_year_rollovers[0]["evidence"], "SO COMPLETION DATE")

    def test_baseline_status_and_header_roll_forward_independently(self) -> None:
        self.previous["STATUS DELIVERY 8 October 2026"] = ["#N/A"]
        for report, previous, expected_header in (
            (date(2026, 10, 13), Path("Daily Tracking IDE 12 October 2026.xlsx"), "STATUS DELIVERY 8 October 2026"),
            (date(2026, 11, 2), Path("Daily Tracking IDE 30 October 2026.xlsx"), "STATUS DELIVERY 30 October 2026"),
        ):
            data = self._build(self.raw, report_date=report, previous_path=previous)
            self.assertEqual(data.at[0, expected_header], "#N/A")
            self.assertEqual(data.at[0, "STATUS DELIVERY"], "SO Complete")

    def test_daily_placeholder_is_refreshed_without_resetting_genuine_history(self) -> None:
        self.raw["STATUS DELIVERY"] = [None]
        self.previous["STATUS DELIVERY"] = [None]
        self.previous["Pre-Installation Start Date"] = [datetime(2026, 10, 11)]
        row = self._build(self.raw, previous_path=Path("Daily Tracking IDE 12 October 2026.xlsx")).iloc[0]
        self.assertEqual(row["Pre-Installation Start Date"], datetime(2026, 10, 11))
        row = self._build(self.raw, report_date=date(2026, 10, 13),
                          previous_path=Path("Daily Tracking IDE 12 October 2026.xlsx")).iloc[0]
        self.assertEqual(row["Pre-Installation Start Date"], datetime(2026, 10, 12))
        self.previous["Pre-Installation Start Date"] = [datetime(2026, 10, 1)]
        row = self._build(self.raw).iloc[0]
        self.assertEqual(row["Pre-Installation Start Date"], datetime(2026, 10, 1))

    def test_continuation_financial_fallback_is_only_for_quotes_absent_from_snapshot(self) -> None:
        self.previous["OTC"] = [None]
        with patch("generate_ide_tracking.load_collabs_lookup", return_value={"NEW": (500, 700)}) as load_csv:
            raw = pd.concat([self.raw] * 2, ignore_index=True)
            raw[QUOTE_ID_HEADER] = ["Q-1", "NEW"]
            built = build_all_order(raw, self.previous, Path("collabs.csv"), date(2026, 10, 13),
                                    Path("Daily Tracking IDE 12 October 2026.xlsx"))
        load_csv.assert_called_once_with(Path("collabs.csv"), {"NEW"})
        self.assertEqual(built.dataframe["OTC"].tolist(), [0, 500])
        self.assertEqual(built.dataframe["MRC"].tolist(), [200, 700])

    def test_po_numeric_compatibility_preserves_leading_zero_and_long_identifiers(self) -> None:
        raw = pd.concat([self.raw] * 4, ignore_index=True)
        raw["NOMOR PO"] = ["1234", "001234", "1234567890123456", "0"]
        self.assertEqual(self._build(raw)["NOMOR PO"].tolist(), [1234, "001234", "1234567890123456", 0])

    def test_refresh_keeps_ancillary_native_and_text_representations(self) -> None:
        raw = pd.concat([self.raw] * 3, ignore_index=True)
        raw["CASH DATE"] = ["26/06/2026", datetime(2926, 6, 26), 46028]
        result = self._build(raw, report_date=date(2026, 11, 2),
                             previous_path=Path("Daily Tracking IDE 30 October 2026.xlsx"))
        self.assertEqual(result["CASH DATE"].tolist(), [
            "26/06/2026", datetime(2926, 6, 26), datetime(2026, 1, 6),
        ])

    def test_refresh_retains_source_month_separately_but_primary_stays_calculated(self) -> None:
        self.raw["MONTH COMPLETE"] = ["Okt"]
        built = build_all_order(self.raw, self.previous, Path("unused.csv"), date(2026, 11, 2),
                                Path("Daily Tracking IDE 30 October 2026.xlsx"))
        self.assertTrue(is_missing(built.dataframe.at[0, "MONTH COMPLETE"]))
        self.assertEqual(built.dataframe.at[0, "MONTH COMPLETE2"], "Okt")
        self.assertEqual(built.month_complete_format, "MMM-YY")
        self.assertEqual(built.dataframe.columns.get_loc("MONTH COMPLETE"), 32)

    def test_abbreviated_month_format_is_retained_on_later_days(self) -> None:
        self.previous["MONTH COMPLETE"] = ["Nov-26"]
        built = build_all_order(self.raw, self.previous, Path("unused.csv"), date(2026, 11, 5),
                                Path("Daily Tracking IDE 4 November 2026.xlsx"))
        self.assertEqual(built.month_complete_format, "MMM-YY")

    def test_rfs_text_reproduces_approved_legacy_conversion_without_per_quote_rules(self) -> None:
        for text, expected in (
            ("12/02/2026", datetime(2026, 12, 2)),
            ("03/04/2027", datetime(2027, 3, 4)),
            ("26/02/2026", datetime(2026, 2, 26)),
        ):
            with self.subTest(text=text):
                raw = self.raw.copy()
                raw[NEW_RFS_INITIAL_HEADER] = [text]
                raw[ACTUAL_RFS_DATE_HEADER] = [text]
                result = self._build(raw)
                self.assertEqual(result.at[0, NEW_RFS_INITIAL_HEADER], expected)
                self.assertEqual(result.at[0, ACTUAL_RFS_DATE_HEADER], expected)

    def test_legacy_rfs_conversion_is_applied_consistently_in_metrics(self) -> None:
        raw = self.raw.copy()
        raw[QUOTE_ID_HEADER] = ["NEW"]
        raw[NEW_RFS_INITIAL_HEADER] = ["02/26/2026"]
        raw[ACTUAL_RFS_DATE_HEADER] = ["02/26/2026"]
        row = self._build(raw).iloc[0]
        self.assertEqual(row[NEW_RFS_INITIAL_HEADER], datetime(2026, 2, 26))
        self.assertEqual(row[ACTUAL_RFS_DATE_HEADER], datetime(2026, 2, 26))
        self.assertEqual(row["TARGET  Detemined as 1 Oct 26"], "Before October")
        self.assertEqual(row["Aging Of RFS"], (date(2026, 10, 12) - date(2026, 2, 26)).days)

    def test_reviewed_typed_input_is_not_flipped_to_match_legacy_text(self) -> None:
        raw = self.raw.copy()
        raw[NEW_RFS_INITIAL_HEADER] = [datetime(2026, 2, 12)]
        typed = self._build(raw).iloc[0]
        self.assertEqual(typed[NEW_RFS_INITIAL_HEADER], datetime(2026, 2, 12))
        self.assertEqual(typed["TARGET  Detemined as 1 Oct 26"], "Before October")
        raw[NEW_RFS_INITIAL_HEADER] = ["12/02/2026"]
        text = self._build(raw).iloc[0]
        self.assertEqual(text[NEW_RFS_INITIAL_HEADER], datetime(2026, 12, 2))
        self.assertEqual(text["TARGET  Detemined as 1 Oct 26"], "After October")

    def test_reference_financial_snapshot_uses_exact_keys_and_zero_without_csv_override(self) -> None:
        for quote in ("Q-1 ", "NEW"):
            with self.subTest(quote=quote):
                raw = self.raw.copy()
                raw[QUOTE_ID_HEADER] = [quote]
                with patch("generate_ide_tracking.load_collabs_lookup") as load_csv:
                    result = build_all_order(
                        raw, self.previous, Path("collabs.csv"), date(2026, 11, 2),
                        Path("Daily Tracking IDE 30 October 2026.xlsx"),
                    )
                load_csv.assert_not_called()
                self.assertEqual(result.dataframe["OTC"].tolist(), [0])
                self.assertEqual(result.dataframe["MRC"].tolist(), [0])
                self.assertEqual(result.financial_source_policy, "reference-snapshot")

    def test_collabs_can_supply_financials_when_reference_has_no_financial_snapshot(self) -> None:
        previous = self.previous.drop(columns=["OTC", "MRC"])
        with patch("generate_ide_tracking.load_collabs_lookup", return_value={"Q-1": (500, 700)}):
            result = build_all_order(
                self.raw, previous, Path("collabs.csv"), date(2026, 10, 13),
                Path("Daily Tracking IDE 12 October 2026.xlsx"),
            )
        self.assertEqual(result.dataframe["OTC"].tolist(), [500])
        self.assertEqual(result.dataframe["MRC"].tolist(), [700])
        self.assertEqual(result.financial_source_policy, "collabs-fallback")

    def test_month_formula_uses_named_column_positions_and_calculates_before_pivot_refresh(self) -> None:
        from unittest.mock import Mock
        worksheet = Mock()
        dataframe = pd.DataFrame({
            "SO COMPLETION DATE": [datetime(2026, 10, 13), None],
            "MONTH COMPLETE": [None, None],
        })
        apply_month_complete_formula(worksheet, dataframe)
        worksheet.Cells.assert_any_call(2, 2)
        worksheet.Cells.assert_any_call(3, 2)
        month_range = worksheet.Range.return_value
        self.assertEqual(month_range.NumberFormat, "General")
        self.assertEqual(month_range.FormulaR1C1, '=TEXT(RC[-1],"[$-409]mmmm")')
        month_range.Calculate.assert_called_once()

    def test_abbreviated_month_formula_does_not_turn_missing_completion_into_january(self) -> None:
        from unittest.mock import Mock
        worksheet = Mock()
        data = pd.DataFrame({"SO COMPLETION DATE": [None], "MONTH COMPLETE": [None]})
        apply_month_complete_formula(worksheet, data, date_format="MMM-YY")
        self.assertEqual(worksheet.Range.return_value.FormulaR1C1,
                         '=IF(ISNUMBER(RC[-1]),TEXT(RC[-1],"[$-409]mmm-yy"),"")')
        worksheet.Range.return_value.Calculate.assert_called_once()

    def test_na_errors_use_native_formulas_without_the_windows_clipboard(self) -> None:
        from unittest.mock import Mock
        excel = Mock()
        excel.AutoCorrect = SimpleNamespace(AutoFillFormulasInLists=True)
        sheet = Mock()
        data = pd.DataFrame({"STATUS": ["#N/A", "SO Complete", "#N/A"]})
        apply_na_errors(excel, sheet, data)
        target = sheet.Range.return_value
        self.assertEqual(target.Formula, "=NA()")
        self.assertEqual(target.Calculate.call_count, 2)
        target.Copy.assert_not_called()
        target.PasteSpecial.assert_not_called()
        self.assertTrue(excel.AutoCorrect.AutoFillFormulasInLists)

    def _delivery_pivot(self, column, fields):
        from unittest.mock import Mock
        pivot = Mock()
        pivot.Name = column
        pivot.ColumnFields = [SimpleNamespace(Name=column)]
        pivot.PivotFields.return_value = fields
        pivot.PageFields.Count = 4
        return pivot

    def _delivery_field(self, name, items, selected=None):
        from unittest.mock import Mock
        field = Mock()
        field.Name = name
        field.Position = 1
        field.Orientation = 3
        field.CurrentPage = "(All)"
        field.EnableMultiplePageItems = True
        values = [SimpleNamespace(Name=item, Visible=selected is None or item in selected) for item in items]
        field.PivotItems.return_value = values
        field.ClearAllFilters.side_effect = lambda: [setattr(item, "Visible", True) for item in values]
        return field

    def test_ontime_cumulative_clears_both_month_and_status_checkboxes(self) -> None:
        for month_names in (("January", "October"), ("Okt",), ("November",), ()):
            with self.subTest(months=month_names):
                field = self._delivery_field("MONTH  COMPLETE", month_names, ["October"])
                status = self._delivery_field("STATUS DELIVERY", ["SO Complete", "06-Installation"], ["SO Complete"])
                pivot = self._delivery_pivot("ON TIME DELIVERY", [status, field])
                worksheet = FakePivotWorksheet([pivot])
                configure_delivery_pivot_filters(worksheet, (), "cumulative-active", date(2027, 3, 2))
                self.assertEqual(field.Position, 3)
                self.assertTrue(all(item.Visible for item in field.PivotItems()))
                self.assertTrue(all(item.Visible for item in status.PivotItems()))
                pivot.RefreshTable.assert_called_once()

    def test_existing_period_keeps_effective_checkbox_filters_not_just_all_caption(self) -> None:
        month = self._delivery_field("MONTH COMPLETE", ["February", "March"], ["March"])
        status = self._delivery_field("STATUS DELIVERY", ["SO Complete", "06-Installation"], ["SO Complete"])
        pivot = self._delivery_pivot("ON TIME DELIVERY", [month, status])
        sheet = FakePivotWorksheet([pivot])
        states = capture_delivery_pivot_filters(sheet)
        self.assertEqual(states[0].selected_items, ("March",))
        self.assertFalse(states[0].show_all)
        for item in month.PivotItems():
            item.Visible = True
        configure_delivery_pivot_filters(sheet, states, "existing-period", date(2027, 3, 2))
        self.assertEqual([item.Name for item in month.PivotItems() if item.Visible], ["March"])
        self.assertEqual([item.Name for item in status.PivotItems() if item.Visible], ["SO Complete"])

    def test_single_current_page_is_captured_even_if_all_items_report_visible(self) -> None:
        month = self._delivery_field("MONTH COMPLETE", ["February", "March"])
        month.CurrentPage = "March"
        month.EnableMultiplePageItems = False
        status = self._delivery_field("STATUS DELIVERY", ["SO Complete"])
        sheet = FakePivotWorksheet([self._delivery_pivot("ON TIME DELIVERY", [month, status])])
        self.assertEqual(capture_delivery_pivot_filters(sheet)[0].selected_items, ("March",))

    def test_single_all_page_does_not_misread_excel_hidden_item_flags(self) -> None:
        month = self._delivery_field("MONTH COMPLETE", ["February", "March"], [])
        month.EnableMultiplePageItems = False
        status = self._delivery_field("STATUS DELIVERY", ["SO Complete"])
        sheet = FakePivotWorksheet([self._delivery_pivot("ON TIME DELIVERY", [month, status])])
        state = capture_delivery_pivot_filters(sheet)[0]
        self.assertTrue(state.show_all)
        self.assertEqual(state.selected_items, ("February", "March"))

    def test_new_period_aging_includes_waiting_aos_but_not_closed_orders(self) -> None:
        status = self._delivery_field("STATUS DELIVERY", [
            "04-Pre-Installation", "07-Waiting AOS/UAT On Hold", "SO Complete", "Cancel", "(blank)",
        ], ["04-Pre-Installation"])
        pivot = self._delivery_pivot("STATUS ORDER", [status])
        configure_delivery_pivot_filters(FakePivotWorksheet([pivot]), (), "cumulative-active", date(2027, 3, 2))
        self.assertEqual([item.Name for item in status.PivotItems() if item.Visible], [
            "04-Pre-Installation", "07-Waiting AOS/UAT On Hold", "(blank)",
        ])

    def test_policy_uses_report_period_not_missing_month_items(self) -> None:
        self.assertEqual(delivery_pivot_policy(date(2027, 2, 26), date(2027, 2, 27)), "existing-period")
        self.assertEqual(delivery_pivot_policy(date(2027, 2, 28), date(2027, 3, 1)), "existing-period")
        self.assertEqual(delivery_pivot_policy(date(2027, 2, 28), date(2027, 3, 2)), "cumulative-active")
        self.assertEqual(delivery_pivot_policy(date(2026, 12, 31), date(2027, 1, 2)), "cumulative-active")
        self.assertEqual(delivery_pivot_policy(None, date(2027, 1, 2)), "cumulative-active")

    def test_ontime_pivot_missing_month_field_remains_an_error(self) -> None:
        from unittest.mock import Mock
        pivot = Mock()
        pivot.Name = "OnTime"
        pivot.ColumnFields = [SimpleNamespace(Name="ON TIME DELIVERY")]
        pivot.PivotFields.return_value = []
        with self.assertRaisesRegex(RuntimeError, "OnTime.*MONTH COMPLETE"):
            configure_delivery_pivot_filters(FakePivotWorksheet([pivot]), (), "cumulative-active", date(2027, 3, 2))

    def test_legacy_template_without_month_page_uses_report_month_when_available(self) -> None:
        for month_names, expected in ((["January", "March"], ["March"]), (["January"], ["January"])):
            with self.subTest(months=month_names):
                month = self._delivery_field("MONTH COMPLETE", month_names)
                month.Orientation = 0
                status = self._delivery_field("STATUS DELIVERY", ["SO Complete"])
                sheet = FakePivotWorksheet([self._delivery_pivot("ON TIME DELIVERY", [month, status])])
                states = capture_delivery_pivot_filters(sheet)
                self.assertFalse(states[0].is_page_field)
                configure_delivery_pivot_filters(sheet, states, "existing-period", date(2027, 3, 8))
                self.assertEqual([item.Name for item in month.PivotItems() if item.Visible], expected)

    def test_preinstallation_does_not_invent_start_for_new_unknown_or_closed_status(self) -> None:
        for status in (None, "SO Complete", "Cancel", "Waiting AOS"):
            with self.subTest(status=status):
                raw = self.raw.copy()
                raw[QUOTE_ID_HEADER] = ["NEW"]
                raw["STATUS DELIVERY"] = [status]
                row = self._build(raw, report_date=date(2026, 11, 2), previous_path=Path("Daily Tracking IDE 30 October 2026.xlsx")).iloc[0]
                self.assertTrue(is_missing(row["Pre-Installation Start Date"]))
                self.assertTrue(is_missing(row["Pre-Installation Aging"]))
                self.assertTrue(is_missing(row["Range Aging On Pre-Installation"]))

    def test_preinstallation_starts_for_observed_preinstallation_or_installation(self) -> None:
        for status in ("Pre-Installation", "Installation"):
            with self.subTest(status=status):
                raw = self.raw.copy()
                raw[QUOTE_ID_HEADER] = ["NEW"]
                raw["STATUS DELIVERY"] = [status]
                row = self._build(raw).iloc[0]
                self.assertEqual(row["Pre-Installation Start Date"], datetime(2026, 10, 11))
                self.assertEqual(row["Pre-Installation Aging"], 1)

    def test_unknown_status_does_not_erase_known_historical_aging(self) -> None:
        raw = self.raw.copy()
        raw["STATUS DELIVERY"] = [None]
        self.previous["STATUS DELIVERY"] = [None]
        row = self._build(raw).iloc[0]
        self.assertEqual(row["Pre-Installation Start Date"], datetime(2026, 10, 1))
        self.assertEqual(row["Pre-Installation Aging"], 11)
        self.assertEqual(row["Range Aging On Pre-Installation"], "1-7 Days")

    def test_preinstallation_history_uses_original_quote_key_like_excel(self) -> None:
        for quote in ("Q-1 ", "Q-1\n"):
            with self.subTest(quote=quote):
                raw = self.raw.copy()
                raw[QUOTE_ID_HEADER] = [quote]
                row = self._build(raw, report_date=date(2026, 11, 2), previous_path=Path("Daily Tracking IDE 30 October 2026.xlsx")).iloc[0]
                self.assertEqual(row[QUOTE_ID_HEADER], quote)
                self.assertTrue(is_missing(row["Pre-Installation Start Date"]))
                self.assertTrue(is_missing(row["Pre-Installation Aging"]))
        raw = self.raw.copy()
        raw[QUOTE_ID_HEADER] = ["q-1"]
        self.assertEqual(self._build(raw, report_date=date(2026, 11, 2), previous_path=Path("Daily Tracking IDE 30 October 2026.xlsx")).iloc[0]["Pre-Installation Start Date"], datetime(2026, 10, 1))

    def test_all_shared_cache_pivots_keep_formatting_and_do_not_reset_column_widths(self) -> None:
        pivots = [SimpleNamespace(PreserveFormatting=False, HasAutoFormat=True) for _ in range(3)]
        preserve_pivot_refresh_layout(FakePivotWorksheet(pivots))
        for pivot in pivots:
            self.assertTrue(pivot.PreserveFormatting)
            self.assertFalse(pivot.HasAutoFormat)

    def test_missing_required_column_has_a_specific_error(self) -> None:
        with self.assertRaisesRegex(ValueError, "missing required columns: PARTNERS"):
            self._build(self.raw.drop(columns=["PARTNERS"]))

    def test_duplicate_normalized_headers_are_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "duplicate columns: QUOTE ID"):
            validate_raw_headers([*self.RAW_HEADERS, "  QUOTE   ID  "])

    def test_extra_column_preserves_date_like_text_and_formats_typed_dates(self) -> None:
        ranges = {}

        class Worksheet:
            def Range(self, address):
                return ranges.setdefault(address, SimpleNamespace(NumberFormat="General"))

        dataframe = pd.DataFrame({
            "MONTH COMPLETE": ["4/30/2026", datetime(2026, 4, 30), None, 7],
            "SOURCE NOTES": ["A", "B", None, "3/4/2026"],
            "COUNT": [0, 1, None, 2],
        }, dtype=object)
        worksheet = Worksheet()
        apply_text_formats(worksheet, dataframe)
        apply_date_formats(worksheet, dataframe)
        self.assertEqual(ranges["A2:A2"].NumberFormat, "@")
        self.assertEqual(ranges["A3:A3"].NumberFormat, "mm/dd/yy")
        self.assertEqual(ranges["B2:B3"].NumberFormat, "@")
        self.assertEqual(ranges["B5:B5"].NumberFormat, "@")
        self.assertFalse(any(address.startswith("C") for address in ranges))

    def test_rfs_display_is_day_first_without_reformatting_other_date_columns(self) -> None:
        ranges = {}

        class Worksheet:
            def Range(self, address):
                return ranges.setdefault(address, SimpleNamespace(NumberFormat="General"))

        data = pd.DataFrame({
            NEW_RFS_INITIAL_HEADER: [datetime(2026, 2, 12), datetime(2026, 12, 2, 8, 30)],
            ACTUAL_RFS_DATE_HEADER: [datetime(2026, 2, 26), datetime(2026, 3, 2, 9, 15)],
            "FAB UPLOAD DATE": [datetime(2026, 2, 10), datetime(2026, 2, 10, 7, 30)],
        })
        apply_date_formats(Worksheet(), data)
        self.assertEqual(ranges["A2:A2"].NumberFormat, "dd/mm/yyyy")
        self.assertEqual(ranges["A3:A3"].NumberFormat, "dd/mm/yyyy h:mm AM/PM")
        self.assertEqual(ranges["B2:B2"].NumberFormat, "dd/mm/yyyy")
        self.assertEqual(ranges["B3:B3"].NumberFormat, "dd/mm/yyyy h:mm AM/PM")
        self.assertEqual(ranges["C2:C2"].NumberFormat, "mm/dd/yy")
        self.assertEqual(ranges["C3:C3"].NumberFormat, "mm/dd/yy h:mm AM/PM")

    def _output_workbook(self, directory):
        headers = ["QUOTE ID", "OTC", "MRC", "YEAR FAB UPLOAD"]
        headers.extend(f"Source field {index}" for index in range(1, 58))
        headers.append("MONTH COMPLETE")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "ALL ORDER"
        sheet.append(headers)
        sheet.append(["Q-1", 100, 200, "2026", *([None] * 57), "Okt"])
        sheet.cell(2, 4).number_format = "@"
        sheet.freeze_panes = "C1"
        sheet.add_table(Table(displayName="Table47", ref=f"A1:{get_column_letter(len(headers))}2"))
        workbook.create_sheet("PIVOT")
        return workbook, Path(directory) / "output.xlsx", headers

    def test_output_accepts_source_column_growth_and_ignores_unused_cell_formatting(self) -> None:
        with TemporaryDirectory() as directory:
            workbook, path, headers = self._output_workbook(directory)
            workbook["ALL ORDER"].cell(2, 70).number_format = "#,##0"
            workbook.save(path)
            workbook.close()
            result = validate_output(path, ["Q-1"], 0, expected_headers=headers)
            self.assertEqual(result["columns"], 62)
            self.assertEqual(result["table_ref"], "A1:BJ2")

    def test_output_rejects_changed_header_even_when_column_count_matches(self) -> None:
        with TemporaryDirectory() as directory:
            workbook, path, headers = self._output_workbook(directory)
            workbook["ALL ORDER"].cell(1, len(headers)).value = "OTHER FIELD"
            workbook.save(path)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "column names or order"):
                validate_output(path, ["Q-1"], 0, expected_headers=headers)

    def test_output_table_must_include_the_added_source_column(self) -> None:
        with TemporaryDirectory() as directory:
            workbook, path, headers = self._output_workbook(directory)
            workbook["ALL ORDER"].tables["Table47"].ref = "A1:BI2"
            workbook.save(path)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "Table47 range"):
                validate_output(path, ["Q-1"], 0, expected_headers=headers)

    def test_output_rejects_a_pivot_cache_that_did_not_pick_up_new_columns(self) -> None:
        with TemporaryDirectory() as directory:
            workbook, path, headers = self._output_workbook(directory)
            workbook["PIVOT"]._pivots = [SimpleNamespace(
                name="StalePivot",
                cache=SimpleNamespace(cacheFields=[
                    SimpleNamespace(name=header) for header in headers[:-1]
                ]),
            )]
            with patch("generate_ide_tracking.load_workbook", return_value=workbook):
                with self.assertRaisesRegex(ValueError, "StalePivot.*source columns"):
                    validate_output(path, ["Q-1"], 1, expected_headers=headers)


class IdeRawDateSourceTests(unittest.TestCase):
    def test_invalid_quote_year_uses_agreeing_source_without_changing_time(self) -> None:
        raw = pd.DataFrame({QUOTE_ID_HEADER: ["A"], "QUOTE DATE": [datetime(3036, 8, 30)]})
        source = SimpleNamespace(iter_rows=lambda **_: iter([
            ("quote_num", "quote_created"), ("A", "2026-08-30 00:07:05"),
            ("A", "2026-08-30 00:07:05"),
        ]))
        changes = recover_source_quote_date_years(raw, source)
        self.assertEqual(raw.at[0, "QUOTE DATE"], datetime(2026, 8, 30))
        self.assertEqual(len(changes), 1)
        self.assertEqual(changes[0]["evidence"], "source.quote_created")

    def test_quote_year_recovery_does_not_guess_from_report_year_or_conflicting_evidence(self) -> None:
        for original, source_values in (
            (datetime(2025, 8, 30), ["2026-08-30"]),
            (datetime(3029, 8, 30), ["2029-08-29"]),
            (datetime(3029, 8, 30), ["2029-08-30", "2028-08-30"]),
            (datetime(3029, 8, 30), ["3029-08-30"]),
            (datetime(3029, 8, 30), []),
        ):
            with self.subTest(original=original, source_values=source_values):
                raw = pd.DataFrame({QUOTE_ID_HEADER: ["A"], "QUOTE DATE": [original]})
                source = SimpleNamespace(iter_rows=lambda **_: iter([
                    ("quote_num", "quote_created"), *[("A", value) for value in source_values],
                ]))
                self.assertEqual(recover_source_quote_date_years(raw, source), ())
                self.assertEqual(raw.at[0, "QUOTE DATE"], original)

    def _read(self, values, source_values, *, previous=None, continuation=True, extras=()):
        with TemporaryDirectory() as directory:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "ICT ORDER 2026"
            headers = (*IdeDynamicColumnsTests.RAW_HEADERS, *extras)
            sheet.append(headers)
            sheet.append([None] * len(headers))
            for quote_id, value in values:
                row = {QUOTE_ID_HEADER: quote_id, NEW_RFS_INITIAL_HEADER: value, None: "scratch"}
                sheet.append([row.get(header) for header in headers])
            if source_values is not None:
                source = workbook.create_sheet("source")
                for row_number, (quote_id, value) in enumerate(source_values, start=1):
                    source.cell(row_number, 4, quote_id)
                    source.cell(row_number, 38, value)
            path = Path(directory) / "dashboard.xlsx"
            workbook.save(path)
            workbook.close()
            return read_raw_dataframe(path, previous=previous, reference_continuation=continuation)

    def test_error_recovery_prioritizes_occurrence_history_only_for_continuation(self) -> None:
        previous = pd.DataFrame({QUOTE_ID_HEADER: ["A", "A"],
                                 NEW_RFS_INITIAL_HEADER: [datetime(2026, 4, 30), datetime(2026, 4, 29)]})
        for continuation, expected in ((True, datetime(2026, 4, 29)), (False, datetime(2026, 4, 30))):
            result = self._read([("A", "30/04/2026"), ("A", "#REF!")], None,
                                previous=previous, continuation=continuation)
            self.assertEqual(result.at[1, NEW_RFS_INITIAL_HEADER], expected)
            self.assertIn(("a", 1), result.attrs["rfs_original_text_rows"][NEW_RFS_INITIAL_HEADER])

    def test_unlabelled_scratch_tail_is_dropped_without_removing_named_empty_column(self) -> None:
        result = self._read([("A", None)], None, extras=("SOURCE NOTE", None))
        self.assertIn("SOURCE NOTE", result)
        self.assertTrue(is_missing(result.at[0, "SOURCE NOTE"]))
        self.assertFalse(any(str(header).startswith("Unnamed:") for header in result))

    def test_previous_native_na_errors_are_not_collapsed_to_missing(self) -> None:
        with TemporaryDirectory() as directory:
            book = Workbook()
            book.active.title = "ALL ORDER"
            book.active.append([QUOTE_ID_HEADER, "STATUS DELIVERY 1 March 2027"])
            book.active.append(["A", "#N/A"])
            path = Path(directory) / "previous.xlsx"
            book.save(path)
            book.close()
            self.assertEqual(read_previous_dataframe(path).iat[0, 1], "#N/A")

    def test_valid_text_typed_dates_blanks_and_zero_are_not_replaced_by_source(self) -> None:
        values = [
            ("A", "04/05/2026"), ("B", "19//6/2026"),
            ("C", datetime(2025, 4, 4)), ("D", None), ("E", "-"), ("F", 0),
        ]
        result = self._read(values, [(quote, datetime(2026, 12, 1)) for quote, _ in values])
        dates = result[NEW_RFS_INITIAL_HEADER].tolist()
        self.assertEqual(dates[:3], ["04/05/2026", datetime(2026, 6, 19), datetime(2025, 4, 4)])
        self.assertTrue(is_missing(dates[3]))
        self.assertEqual(dates[4:], ["-", 0])

    def test_original_quote_key_is_preserved_while_empty_quotes_are_filtered(self) -> None:
        result = self._read([("Q-1 ", None), ("Q-2\n", None), ("   ", None)], None)
        self.assertEqual(result[QUOTE_ID_HEADER].tolist(), ["Q-1 ", "Q-2\n"])

    def test_unreadable_dates_are_recovered_from_source_when_it_has_a_valid_date(self) -> None:
        result = self._read(
            [("A", "#REF!"), ("B", "24/26/2026")],
            [("A", "08-06-2026 20:19:55"), ("B", datetime(2026, 6, 19))],
        )
        self.assertEqual(result[NEW_RFS_INITIAL_HEADER].tolist(), [
            datetime(2026, 8, 6), datetime(2026, 6, 19),
        ])

    def test_invalid_source_does_not_erase_the_error_or_use_a_later_duplicate(self) -> None:
        result = self._read(
            [("A", "#REF!")], [("A", None), ("A", datetime(2026, 8, 6))],
        )
        self.assertEqual(result[NEW_RFS_INITIAL_HEADER].tolist(), ["#REF!"])

    def test_broken_quote_lookup_prefers_agreeing_current_dates_before_source(self) -> None:
        result = self._read(
            [("A", "30/04/2026"), ("A", "#REF!"), ("A", datetime(2026, 4, 30))],
            [("A", datetime(2026, 4, 29))],
        )
        self.assertEqual(result[NEW_RFS_INITIAL_HEADER].tolist(), [
            "30/04/2026", datetime(2026, 4, 30), datetime(2026, 4, 30),
        ])

    def test_conflicting_current_dates_do_not_invent_a_first_match_recovery(self) -> None:
        result = self._read(
            [("A", "15/04/2026"), ("A", "30/04/2026"), ("A", "#REF!")], [],
        )
        self.assertEqual(result[NEW_RFS_INITIAL_HEADER].tolist(), [
            "15/04/2026", "30/04/2026", "#REF!",
        ])

    def test_agreeing_current_dates_can_recover_a_lookup_without_a_source_sheet(self) -> None:
        result = self._read([("A", "30/04/2026"), ("A", "#REF!")], None)
        self.assertEqual(result[NEW_RFS_INITIAL_HEADER].tolist(), [
            "30/04/2026", datetime(2026, 4, 30),
        ])


class IdePlannedRfsYearTests(unittest.TestCase):
    def test_text_year_uses_completion_year_before_report_cutoff(self) -> None:
        value, changed = planned_rfs_with_completion_year(
            "20/06/2025", datetime(2026, 6, 26), date(2026, 11, 2), is_text_source=True,
        )
        self.assertEqual(value, datetime(2026, 6, 20))
        self.assertTrue(changed)

    def test_typed_dates_future_candidates_missing_evidence_and_leap_days_are_unchanged(self) -> None:
        for initial, complete, is_text in (
            (datetime(2025, 6, 20), datetime(2026, 6, 26), False),
            ("23/12/2025", datetime(2026, 6, 26), True),
            ("20/06/2025", None, True),
            ("20/06/2027", datetime(2026, 6, 26), True),
            ("29/02/2024", datetime(2025, 6, 26), True),
        ):
            with self.subTest(initial=initial, complete=complete, is_text=is_text):
                self.assertEqual(planned_rfs_with_completion_year(
                    initial, complete, date(2026, 11, 2), is_text_source=is_text,
                ), (initial, False))

    def test_cutoff_includes_report_day_and_is_not_hardcoded_to_a_year(self) -> None:
        initial = "02/03/2028"
        self.assertEqual(planned_rfs_with_completion_year(
            initial, datetime(2029, 1, 15), date(2029, 2, 3), is_text_source=True,
        ), (datetime(2029, 2, 3), True))
        self.assertEqual(planned_rfs_with_completion_year(
            initial, datetime(2029, 1, 15), date(2029, 2, 2), is_text_source=True,
        ), (initial, False))


class FakeRangeDimension:
    def __init__(self, count):
        self.Count = count


class FakeTableRange:
    def __init__(self, row, column, rows, columns):
        self.Row = row
        self.Column = column
        self.Rows = FakeRangeDimension(rows)
        self.Columns = FakeRangeDimension(columns)


class FakePageField:
    def __init__(self, pivot, name, position):
        self.pivot = pivot
        self.Name = name
        self.Orientation = 3
        self._position = position
        self.position_changes = []
        self.filters_cleared = False

    @property
    def Position(self):
        return self._position

    @Position.setter
    def Position(self, value):
        if not 1 <= value <= self.pivot.PageFields.Count:
            raise ValueError("Unable to set the Position property of the PivotField class")
        self.position_changes.append(value)
        self._position = value

    def ClearAllFilters(self):
        self.filters_cleared = True


class IdePivotPagePositionTests(unittest.TestCase):
    def test_saved_position_is_bounded_by_the_refreshed_page_field_count(self) -> None:
        pivot = SimpleNamespace(Name="PivotTable2", PageFields=SimpleNamespace(Count=3))
        field = FakePageField(pivot, "TARGET", 1)
        set_pivot_page_field_position(pivot, field, 4)
        self.assertEqual(field.Position, 3)
        self.assertEqual(field.position_changes, [3])

    def test_already_valid_position_does_not_write_to_excel_again(self) -> None:
        pivot = SimpleNamespace(Name="PivotTable2", PageFields=SimpleNamespace(Count=3))
        field = FakePageField(pivot, "TARGET", 3)
        set_pivot_page_field_position(pivot, field, 4)
        self.assertEqual(field.position_changes, [])

    def test_existing_position_is_preserved_when_all_filters_survive_refresh(self) -> None:
        pivot = SimpleNamespace(Name="PivotTable2", PageFields=SimpleNamespace(Count=4))
        field = FakePageField(pivot, "TARGET", 2)
        set_pivot_page_field_position(pivot, field, 4)
        self.assertEqual(field.Position, 4)

    def test_target_selection_is_restored_after_month_header_normalization_removes_a_filter(self) -> None:
        from unittest.mock import Mock
        pivot = Mock()
        pivot.Name = "PivotTable2"
        pivot.PageFields.Count = 3
        field = FakePageField(pivot, "TARGET  Detemined as 1 Oct 26", 3)
        pivot.PivotFields.return_value = FakePivotCollection([field])
        field.PivotItems = lambda: FakePivotCollection([SimpleNamespace(Name="Target October")])
        worksheet = Mock()
        worksheet.PivotTables.return_value = pivot
        state = TargetPivotPageFilter("PivotTable2", 4, ("target",), False)

        restore_target_pivot_filters(worksheet, [state], date(2026, 10, 14))

        self.assertEqual(field.Position, 3)
        self.assertEqual(field.position_changes, [])
        self.assertTrue(field.filters_cleared)
        self.assertEqual(field.CurrentPage, "Target October")
        pivot.RefreshTable.assert_called_once()

    def test_empty_page_field_collection_is_reported(self) -> None:
        pivot = SimpleNamespace(Name="PivotTable2", PageFields=SimpleNamespace(Count=0))
        field = FakePageField(pivot, "TARGET", 1)
        with self.assertRaisesRegex(RuntimeError, "PivotTable2.*no page filters"):
            set_pivot_page_field_position(pivot, field, 4)

    def test_position_failure_keeps_pivot_and_field_context(self) -> None:
        class FailingField(FakePageField):
            @FakePageField.Position.setter
            def Position(self, value):
                raise ValueError("Excel rejected the change")

        pivot = SimpleNamespace(Name="PivotTable2", PageFields=SimpleNamespace(Count=4))
        field = FailingField(pivot, "MONTH COMPLETE", 4)
        with self.assertRaisesRegex(RuntimeError, "MONTH COMPLETE.*PivotTable2.*3 of 4"):
            set_pivot_page_field_position(pivot, field, 3)


class FakePivot:
    def __init__(self, name, row, column, rows, columns, row_range=None):
        self.Name = name
        self.TableRange2 = FakeTableRange(row, column, rows, columns)
        start, count = row_range if row_range is not None else (row, rows)
        self.RowRange = FakeTableRange(start, column, count, 1)


class FakePivotCollection:
    def __init__(self, pivots):
        self._pivots = pivots
        self.Count = len(pivots)

    def __call__(self, index):
        return self._pivots[index - 1]

    def __iter__(self):
        return iter(self._pivots)


class FakePivotWorksheet:
    def __init__(self, pivots):
        self._pivots = FakePivotCollection(pivots)

    def PivotTables(self):
        return self._pivots


class FakeUsedRange:
    def __init__(self, values, row=1):
        self.Value2 = values
        self.Row = row


class FakeHiddenRows:
    def __init__(self, worksheet, start_row, end_row):
        self.worksheet = worksheet
        self.start_row = start_row
        self.end_row = end_row

    @property
    def EntireRow(self):
        return self

    @property
    def Hidden(self):
        return None

    @Hidden.setter
    def Hidden(self, value):
        self.worksheet.visibility_changes.append(
            (self.start_row, self.end_row, bool(value))
        )


class FakeVisibilityWorksheet(FakePivotWorksheet):
    def __init__(self, pivots, values):
        super().__init__(pivots)
        self.UsedRange = FakeUsedRange(values)
        self.visibility_changes = []

    def Cells(self, row, column):
        return row, column

    def Range(self, start_cell, end_cell):
        return FakeHiddenRows(self, start_cell[0], end_cell[0])

    def Rows(self, row):
        return FakeHiddenRows(self, row, row)


class IdeDateParsingTests(unittest.TestCase):
    def test_refreshed_pivot_rows_are_unhidden_before_gap_rows(self) -> None:
        values = tuple(
            (value,)
            for value in [
                "TARGET COMPLETE",
                None,
                "header",
                "row 1",
                "row 2",
                "row 3",
                "row 4",
                "Grand Total",
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                "ORDER AGING START FROM PRE-INSTALLATION STATUS",
            ]
        )
        worksheet = FakeVisibilityWorksheet(
            [FakePivot("target-complete", 3, 1, 6, 4)],
            values,
        )

        normalize_pivot_section_row_visibility(worksheet)

        self.assertIn((3, 8, False), worksheet.visibility_changes)
        self.assertIn((1, 1, False), worksheet.visibility_changes)
        self.assertIn((20, 20, False), worksheet.visibility_changes)
        self.assertIn((10, 17, True), worksheet.visibility_changes)

    def test_pivot_filter_rows_remain_hidden(self) -> None:
        values = tuple(
            (value,)
            for value in [
                "TARGET COMPLETE",
                None,
                "New Registration",
                None,
                "TARGET  Detemined as 1 Sep 26",
                "YEAR FAB UPLOAD",
                "YEAR FAB UPLOAD",
                "Count of QUOTE ID",
                "Div./Dept.",
                "Order",
                "Grand Total",
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                "ORDER AGING START FROM PRE-INSTALLATION STATUS",
            ]
        )
        worksheet = FakeVisibilityWorksheet(
            [FakePivot("target-complete", 5, 4, 7, 7, row_range=(9, 3))],
            values,
        )

        normalize_pivot_section_row_visibility(worksheet)

        self.assertIn((9, 11, False), worksheet.visibility_changes)
        self.assertFalse(
            any(
                start <= row <= end and not hidden
                for start, end, hidden in worksheet.visibility_changes
                for row in range(5, 9)
            )
        )

    def test_freeze_pane_validation_ignores_horizontal_scroll_position(self) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.freeze_panes = "C1"
        worksheet.sheet_view.pane.topLeftCell = "N1"
        worksheet.sheet_view.pane.state = "frozenSplit"

        self.assertTrue(has_expected_all_order_freeze_panes(worksheet))

        worksheet.sheet_view.pane.xSplit = 2.5
        self.assertFalse(has_expected_all_order_freeze_panes(worksheet))

        worksheet.freeze_panes = "N1"
        self.assertFalse(has_expected_all_order_freeze_panes(worksheet))

    def test_hidden_row_ranges_follow_dynamic_section_boundaries(self) -> None:
        self.assertEqual(
            hidden_row_ranges_between_sections(25, 113),
            [(27, 110)],
        )
        self.assertEqual(
            hidden_row_ranges_between_sections(137, 166, {144}),
            [(139, 142), (145, 163)],
        )

    def test_target_complete_pivots_allow_dynamic_start_rows(self) -> None:
        summary = FakePivot("summary", 7, 1, 45, 3)
        status = FakePivot("status", 6, 4, 46, 7)
        later_left = FakePivot("later-left", 68, 1, 20, 5)
        later_right = FakePivot("later-right", 65, 8, 18, 4)
        worksheet = FakePivotWorksheet(
            [later_right, status, later_left, summary]
        )

        actual_summary, actual_status = identify_target_complete_pivots(worksheet)

        self.assertIs(actual_summary, summary)
        self.assertIs(actual_status, status)

    def test_target_complete_pivots_reject_unrelated_later_section(self) -> None:
        summary = FakePivot("summary", 7, 1, 45, 3)
        unrelated_status = FakePivot("later-right", 65, 8, 18, 4)
        worksheet = FakePivotWorksheet([summary, unrelated_status])

        with self.assertRaisesRegex(RuntimeError, "TARGET COMPLETE"):
            identify_target_complete_pivots(worksheet)

    def test_report_filename_accepts_english_month_with_any_case(self) -> None:
        expected = date(2026, 8, 18)

        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 August 2026.xlsx")), expected)
        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 august 2026.xlsx")), expected)
        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 AUGUST 2026.xlsx")), expected)

    def test_report_filename_accepts_indonesian_month_with_any_case(self) -> None:
        expected = date(2026, 8, 18)

        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 Agustus 2026.xlsx")), expected)
        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 agustus 2026.xlsx")), expected)
        self.assertEqual(date_from_filename(Path("IDE DASHBOARD 18 AGUSTUS 2026.xlsx")), expected)

    def test_report_filename_accepts_english_and_indonesian_month_abbreviations(self) -> None:
        cases = {
            "IDE DASHBOARD 18 Aug 2026.xlsx": date(2026, 8, 18),
            "IDE DASHBOARD 18 AGU 2026.xlsx": date(2026, 8, 18),
            "IDE DASHBOARD 18 Agt 2026.xlsx": date(2026, 8, 18),
            "IDE DASHBOARD 18 Sep 2026.xlsx": date(2026, 9, 18),
            "IDE DASHBOARD 18 Sept 2026.xlsx": date(2026, 9, 18),
            "IDE DASHBOARD 18 Oct 2026.xlsx": date(2026, 10, 18),
            "IDE DASHBOARD 18 Okt 2026.xlsx": date(2026, 10, 18),
            "IDE DASHBOARD 18 Dec 2026.xlsx": date(2026, 12, 18),
            "IDE DASHBOARD 18 Des 2026.xlsx": date(2026, 12, 18),
        }

        for filename, expected in cases.items():
            with self.subTest(filename=filename):
                self.assertEqual(date_from_filename(Path(filename)), expected)

    def test_slash_dates_follow_dashboard_day_first_format(self) -> None:
        self.assertEqual(
            parse_excel_datetime("10/01/2026"),
            datetime(2026, 1, 10),
        )
        self.assertEqual(
            parse_excel_datetime("04/05/2026"),
            datetime(2026, 5, 4),
        )

    def test_iso_dates_are_unchanged(self) -> None:
        self.assertEqual(
            parse_excel_datetime("2026-08-13 07:30:00"),
            datetime(2026, 8, 13, 7, 30),
        )

    def test_month_first_can_be_selected_for_region_sensitive_fields(self) -> None:
        self.assertEqual(
            parse_excel_datetime("08/06/2026", dayfirst=False),
            datetime(2026, 8, 6),
        )
        self.assertEqual(
            parse_excel_datetime("22/06/2026", dayfirst=False),
            datetime(2026, 6, 22),
        )

    def test_dashboard_rfs_uses_legacy_reference_profile_and_fab_keeps_its_policy(self) -> None:
        dataframe = pd.DataFrame(
            {
                NEW_RFS_INITIAL_HEADER: ["08/06/2026"],
                ACTUAL_RFS_DATE_HEADER: ["11/05/2026"],
                "FAB UPLOAD DATE": ["5/20/2026"],
            }
        )

        normalize_dashboard_text_dates(dataframe)

        self.assertEqual(dataframe.at[0, NEW_RFS_INITIAL_HEADER], datetime(2026, 8, 6))
        self.assertEqual(dataframe.at[0, ACTUAL_RFS_DATE_HEADER], datetime(2026, 11, 5))
        self.assertEqual(dataframe.at[0, "FAB UPLOAD DATE"], datetime(2026, 5, 20))

    def test_strict_rfs_contract_covers_day_boundary_and_preserves_typed_dates(self) -> None:
        for header in (NEW_RFS_INITIAL_HEADER, ACTUAL_RFS_DATE_HEADER):
            for value, expected in (
                ("11/02/2026", datetime(2026, 2, 11)),
                ("12/02/2026", datetime(2026, 2, 12)),
                ("13/02/2026", datetime(2026, 2, 13)),
                ("26/02/2026", datetime(2026, 2, 26)),
                ("12-02-2026 07:30:00", datetime(2026, 2, 12, 7, 30)),
                ("29/02/2028", datetime(2028, 2, 29)),
                ("29/02/2026", None),
                ("02/26/2026", None),
                ("02/26/26", None),
                ("2026-12-02 07:30:00", datetime(2026, 12, 2, 7, 30)),
                ("2026-12-02T07:30:00.123", datetime(2026, 12, 2, 7, 30, 0, 123000)),
                ("2026/12/02", datetime(2026, 12, 2)),
                ("2026-26-02", None),
                (datetime(2026, 12, 2), datetime(2026, 12, 2)),
                (None, None),
            ):
                with self.subTest(header=header, value=value):
                    self.assertEqual(
                        parse_dashboard_field_datetime(value, header, rfs_text_policy="day-first"),
                        expected,
                    )

    def test_rfs_profiles_do_not_guess_from_kpis_and_never_flip_typed_dates(self) -> None:
        self.assertEqual(
            parse_dashboard_field_datetime("11/05/2026", NEW_RFS_INITIAL_HEADER),
            datetime(2026, 11, 5),
        )
        self.assertEqual(
            parse_dashboard_field_datetime("11/05/2026", NEW_RFS_INITIAL_HEADER, rfs_text_policy="day-first"),
            datetime(2026, 5, 11),
        )
        for policy in ("legacy-excel", "day-first"):
            self.assertEqual(
                parse_dashboard_field_datetime(datetime(2026, 2, 12), NEW_RFS_INITIAL_HEADER, rfs_text_policy=policy),
                datetime(2026, 2, 12),
            )
        with self.assertRaisesRegex(ValueError, "Unknown RFS text date policy"):
            parse_dashboard_field_datetime("11/05/2026", NEW_RFS_INITIAL_HEADER, rfs_text_policy="guess")

    def test_duplicate_dates_keep_their_own_values_and_do_not_fill_empty_rows(self) -> None:
        dataframe = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["A", "A", "B", "B"],
                NEW_RFS_INITIAL_HEADER: [
                    datetime(2026, 4, 15),
                    datetime(2026, 4, 30),
                    datetime(2026, 6, 8),
                    pd.NA,
                ],
                ACTUAL_RFS_DATE_HEADER: [
                    datetime(2026, 4, 15),
                    datetime(2026, 4, 30),
                    datetime(2026, 7, 1),
                    pd.NA,
                ],
            }
        )

        previous = dataframe.copy()
        previous[NEW_RFS_INITIAL_HEADER] = datetime(2026, 1, 1)
        previous[ACTUAL_RFS_DATE_HEADER] = datetime(2026, 1, 1)
        for header in (NEW_RFS_INITIAL_HEADER, ACTUAL_RFS_DATE_HEADER):
            values, count = reconcile_date_column(dataframe, previous, header, date(2026, 11, 2))
            pd.testing.assert_series_equal(values, dataframe[header], check_names=False)
            self.assertEqual(count, 0)

    def test_source_new_rfs_lookup_uses_first_match_and_legacy_date_parsing(self) -> None:
        class WorksheetStub:
            def iter_rows(self, **_kwargs):
                return iter(
                    [
                        ("A", *([None] * 33), "08-06-2026 20:19:55"),
                        ("A", *([None] * 33), "09-06-2026 20:19:55"),
                    ]
                )

        lookup = load_source_new_rfs_lookup(WorksheetStub(), {"A"})

        self.assertEqual(lookup["A"], datetime(2026, 8, 6))

    def test_missing_current_date_is_not_filled_from_previous(self) -> None:
        value, used_fallback = reconciled_date_value(
            pd.NA,
            datetime(2026, 7, 1),
            date(2026, 8, 14),
        )

        self.assertTrue(is_missing(value))
        self.assertFalse(used_fallback)

    def test_reference_year_is_retained_only_for_a_matching_historical_occurrence(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1", "Q-1"],
                ACTUAL_RFS_DATE_HEADER: [
                    datetime(2025, 12, 16),
                    datetime(2025, 12, 16),
                ],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                ACTUAL_RFS_DATE_HEADER: [datetime(2026, 12, 16)],
            }
        )

        values, fallback_count = reconcile_date_column(
            current,
            previous,
            ACTUAL_RFS_DATE_HEADER,
            date(2026, 8, 14),
        )

        self.assertEqual(list(values), [datetime(2026, 12, 16), datetime(2025, 12, 16)])
        self.assertEqual(fallback_count, 1)

    def test_month_opening_does_not_override_valid_current_rfs_dates(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                NEW_RFS_INITIAL_HEADER: [datetime(2026, 8, 9)],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                NEW_RFS_INITIAL_HEADER: [datetime(2026, 9, 8)],
            }
        )

        values, fallback_count = reconcile_date_column(
            current,
            previous,
            NEW_RFS_INITIAL_HEADER,
            date(2026, 9, 1),
        )

        self.assertEqual(list(values), [datetime(2026, 8, 9)])
        self.assertEqual(fallback_count, 0)

    def test_non_opening_day_keeps_current_rfs_date(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                NEW_RFS_INITIAL_HEADER: [datetime(2026, 8, 9)],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1"],
                NEW_RFS_INITIAL_HEADER: [datetime(2026, 9, 8)],
            }
        )

        values, fallback_count = reconcile_date_column(
            current,
            previous,
            NEW_RFS_INITIAL_HEADER,
            date(2026, 9, 2),
        )

        self.assertEqual(list(values), [datetime(2026, 8, 9)])
        self.assertEqual(fallback_count, 0)

    def test_month_opening_preserves_each_duplicate_current_actual_rfs(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1", "Q-1"],
                ACTUAL_RFS_DATE_HEADER: [
                    datetime(2026, 4, 15),
                    datetime(2026, 4, 15),
                ],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["Q-1", "Q-1"],
                ACTUAL_RFS_DATE_HEADER: [
                    datetime(2026, 4, 15),
                    datetime(2026, 4, 30),
                ],
            }
        )

        values, fallback_count = reconcile_date_column(
            current,
            previous,
            ACTUAL_RFS_DATE_HEADER,
            date(2026, 9, 1),
        )

        self.assertEqual(
            list(values),
            [datetime(2026, 4, 15), datetime(2026, 4, 15)],
        )
        self.assertEqual(fallback_count, 0)

    def test_unreadable_duplicate_dates_fall_back_by_occurrence_not_first_quote(self) -> None:
        current = pd.DataFrame({
            QUOTE_ID_HEADER: ["Q-1", "Q-1"], ACTUAL_RFS_DATE_HEADER: ["#REF!", "#REF!"],
        })
        previous = pd.DataFrame({
            QUOTE_ID_HEADER: ["Q-1", "Q-1"],
            ACTUAL_RFS_DATE_HEADER: [datetime(2026, 4, 15), datetime(2026, 4, 30)],
        })
        values, count = reconcile_date_column(current, previous, ACTUAL_RFS_DATE_HEADER, date(2026, 11, 2))
        self.assertEqual(values.tolist(), previous[ACTUAL_RFS_DATE_HEADER].tolist())
        self.assertEqual(count, 2)

    def test_unreadable_date_without_valid_history_is_reported_not_invented(self) -> None:
        value, used_fallback = reconciled_date_value("24/26/2026", None, date(2026, 11, 2))
        self.assertEqual(value, "24/26/2026")
        self.assertFalse(used_fallback)

    def test_previous_field_fallback_uses_first_quote_match_and_preserves_na_error(self) -> None:
        current = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["A", "A", "B"],
                DEPT_HEADER: [pd.NA, pd.NA, "#N/A"],
            }
        )
        previous = pd.DataFrame(
            {
                QUOTE_ID_HEADER: ["A", "A"],
                DEPT_HEADER: ["First Dept", "Second Dept"],
            }
        )

        fallback_count = apply_previous_field_fallbacks(current, previous, [DEPT_HEADER])

        self.assertEqual(fallback_count, 2)
        self.assertEqual(current[DEPT_HEADER].iloc[:2].tolist(), ["First Dept", "First Dept"])
        self.assertEqual(current.at[2, DEPT_HEADER], "#N/A")

    def test_invalid_date_is_not_coerced(self) -> None:
        self.assertIsNone(parse_excel_datetime("24/26/2026"))

    def test_target_bucket_uses_parsed_calendar_month(self) -> None:
        report_date = date(2026, 8, 13)
        self.assertEqual(target_value(datetime(2026, 1, 10), report_date), "Before August")
        self.assertEqual(target_value("13/08/2026", report_date), "Target August")
        self.assertEqual(target_value(datetime(2026, 9, 1), report_date), "After August")
        self.assertEqual(target_value("10/01/2026", report_date), "After August")
        self.assertEqual(target_value("01/09/2026", report_date), "Before August")

    def test_first_calendar_day_closes_the_previous_target_month(self) -> None:
        self.assertEqual(target_period(date(2026, 9, 1)), date(2026, 8, 31))
        self.assertEqual(target_period(date(2026, 1, 1)), date(2025, 12, 31))

    def test_later_report_days_use_the_current_target_month(self) -> None:
        self.assertEqual(target_period(date(2026, 9, 2)), date(2026, 9, 2))

    def test_missing_current_target_carries_the_previous_bucket_in_same_period(self) -> None:
        self.assertEqual(
            target_value_with_previous(
                pd.NA,
                "Target August",
                date(2026, 8, 31),
            ),
            "Target August",
        )

    def test_current_target_date_takes_precedence_over_previous(self) -> None:
        self.assertEqual(
            target_value_with_previous(
                datetime(2026, 9, 15),
                "Target Not Yet Inputted",
                date(2026, 8, 31),
            ),
            "After August",
        )

    def test_existing_target_is_carried_within_the_same_period(self) -> None:
        self.assertEqual(
            target_value_with_previous(
                datetime(2026, 8, 9),
                "After August",
                date(2026, 8, 31),
            ),
            "After August",
        )

    def test_existing_target_is_reclassified_when_period_changes(self) -> None:
        self.assertEqual(
            target_value_with_previous(
                pd.NA,
                "Target August",
                date(2026, 9, 2),
            ),
            "Target Not Yet Inputted",
        )

    def test_missing_previous_financial_value_falls_back_to_zero(self) -> None:
        self.assertEqual(financial_value_or_zero("1,250,000"), 1_250_000)
        self.assertEqual(financial_value_or_zero("#N/A"), 0)
        self.assertEqual(financial_value_or_zero(None), 0)

    def test_collabs_lookup_reads_only_requested_quote_ids(self) -> None:
        with TemporaryDirectory() as temp_dir:
            collabs_path = Path(temp_dir) / "collabs.csv"
            pd.DataFrame(
                {
                    "quote_num": ["Q-1", "Q-2"],
                    "otc": ["1,250", "9,999"],
                    "mrc": ["250", "999"],
                    "unused": ["x", "y"],
                }
            ).to_csv(collabs_path, index=False)

            lookup = load_collabs_lookup(collabs_path, {"Q-1"})

        self.assertEqual(lookup, {"Q-1": (1_250, 250)})

    def test_financial_fallback_uses_previous_then_collabs_then_zero(self) -> None:
        self.assertEqual(
            resolve_financial_values(100, 200, (300, 400)),
            (100, 200, False, False),
        )
        self.assertEqual(
            resolve_financial_values("#N/A", None, (300, 400)),
            (300, 400, True, False),
        )
        self.assertEqual(
            resolve_financial_values(None, None, (None, "#N/A")),
            (0, 0, False, True),
        )

    def test_completion_thresholds_follow_the_five_business_week_buckets(self) -> None:
        expectations = {
            1: (1, 30, 20, 20),
            6: (1, 30, 20, 20),
            7: (2, 40, 30, 30),
            13: (2, 40, 30, 30),
            14: (3, 50, 40, 40),
            20: (3, 50, 40, 40),
            21: (4, 60, 50, 50),
            28: (4, 60, 50, 50),
            29: (5, 60, 50, 50),
            31: (5, 60, 50, 50),
        }

        for day, expected in expectations.items():
            with self.subTest(day=day):
                thresholds = completion_thresholds(date(2026, 8, day))
                self.assertEqual(
                    (
                        thresholds.week,
                        thresholds.green_percent,
                        thresholds.yellow_percent,
                        thresholds.red_percent,
                    ),
                    expected,
                )


if __name__ == "__main__":
    unittest.main()
