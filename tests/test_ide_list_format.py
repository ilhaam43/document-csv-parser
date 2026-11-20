import unittest
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

from generate_ide_tracking import InputFiles, determine_report_date, generate_ide_tracking, target_header
from ide_list_format import (
    REQUIRED, ColumnPresentation, apply_column_alignment, build_rows, column_presentations,
    generate, is_list_export, normalize_status, presentation_key, read_column_presentations,
    read_records, target_bucket, fit_unwrapped_text_columns, is_target_after_filter,
    percentage_helper_formula, _restore_percentages,
)


class IdeListPercentageTests(unittest.TestCase):
    def test_target_after_filter_follows_bucket_not_month_or_pivot_name(self):
        for period in (date(2026, 11, 17), date(2026, 12, 1), date(2027, 1, 1)):
            with self.subTest(period=period):
                self.assertTrue(is_target_after_filter({target_header(period): {"After " + period.strftime("%B")}}))

    def test_other_or_mixed_filters_do_not_change_formula_policy(self):
        name = target_header(date(2026, 11, 17))
        for fields in ({}, {name: None}, {name: set()}, {name: {"Target November"}},
                       {name: {"Before November"}}, {name: {"Target Not Yet Inputted"}},
                       {name: {"After November", "Target November"}},
                       {"Department": {"After November"}}):
            with self.subTest(fields=fields):
                self.assertFalse(is_target_after_filter(fields))

    def test_target_after_uses_manual_precedence(self):
        self.assertEqual(percentage_helper_formula({"5.SO Complete": ["$I$356"],
                                                   "6.Cancel": ["$K$356"]}, "$C$356", target_after=True),
                         "=IF($C$356=0,0,SUM($I$356)+SUM($K$356)/$C$356)")

    def test_missing_status_is_zero_not_a_neighbouring_column(self):
        self.assertEqual(percentage_helper_formula({"5.SO Complete": ["$I$355"]}, "$C$355", target_after=True),
                         "=IF($C$355=0,0,SUM($I$355)+0/$C$355)")
        self.assertEqual(percentage_helper_formula({"6.Cancel": ["$G$42"]}, "$C$42", target_after=True),
                         "=IF($C$42=0,0,0+SUM($G$42)/$C$42)")
        self.assertEqual(percentage_helper_formula({}, "$D$18", target_after=True),
                         "=IF($D$18=0,0,0+0/$D$18)")

    def test_split_status_cells_are_summed_before_division(self):
        self.assertEqual(percentage_helper_formula({"5.SO Complete": ["$Q$8", "$R$8"],
                                                   "6.Cancel": ["$S$8", "$T$8"]}, "$F$8", target_after=True),
                         "=IF($F$8=0,0,SUM($Q$8,$R$8)+SUM($S$8,$T$8)/$F$8)")

    def test_other_completion_blocks_keep_the_existing_ratio(self):
        self.assertEqual(percentage_helper_formula({"5.SO Complete": ["$I$31"]}, "$C$31"),
                         "=IF($C$31=0,0,($I$31)/$C$31)")
        self.assertEqual(percentage_helper_formula({"5.SO Complete": ["$I$31"],
                                                   "6.Cancel": ["$K$31"]}, "$C$31"),
                         "=IF($C$31=0,0,($I$31+$K$31)/$C$31)")

    def test_restore_chooses_policy_from_filter_and_matches_summary_row_identity(self):
        class Cells(list):
            def __call__(self, row, column):
                return self[column - 1]

        class Items(list):
            @property
            def Count(self):
                return len(self)

            def __call__(self, index):
                return self[index - 1]

        def phase_cell(address, status=None):
            items = Items([SimpleNamespace(Name=status)] if status else [])
            from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
            column = column_index_from_string(coordinate_from_string(address)[0])
            return SimpleNamespace(Address=address, Column=column, identity="department", PivotCell=SimpleNamespace(ColumnItems=items))

        phase_row = SimpleNamespace(Row=48, Cells=Cells([
            phase_cell("$G$48", "6.Cancel"), phase_cell("$H$48", "5.SO Complete"), phase_cell("$I$48"),
        ]))
        count_cell = SimpleNamespace(Address="$C$42", identity="department",
                                     PivotCell=SimpleNamespace(DataField=SimpleNamespace(SourceName="Quote ID")))
        fields = [SimpleNamespace(SourceName="Department")]
        summary = SimpleNamespace(Name="summary", RowFields=fields,
                                  DataFields=[SimpleNamespace(SourceName="Quote ID", Function=-4112)],
                                  TableRange1=SimpleNamespace(Column=1, Row=40),
                                  DataBodyRange=SimpleNamespace(Rows=[SimpleNamespace(Cells=[count_cell])]))
        phase = SimpleNamespace(Name="arbitrary-name", RowFields=fields,
                                TableRange1=SimpleNamespace(Column=7, Row=45, Columns=SimpleNamespace(Count=3)),
                                DataBodyRange=SimpleNamespace(Rows=[phase_row]))
        output = SimpleNamespace()
        sheet = SimpleNamespace(PivotTables=lambda name=None: phase if name else [summary, phase],
                                Cells=lambda row, column: output)
        name = target_header(date(2026, 11, 17))
        for selection, expected in [("After November", "SUM($H$48)+SUM($G$48)/$C$42"),
                                    ("Target November", "($H$48+$G$48)/$C$42")]:
            with self.subTest(selection=selection), \
                 patch("generate_ide_tracking.restore_pivot_percentage_blocks"), \
                 patch("ide_list_format._capture_filters", return_value={phase.Name: {name: {selection}}}), \
                 patch("ide_list_format.pivot_row_identity", side_effect=lambda cell, **kwargs: cell.identity):
                _restore_percentages(sheet, None, [SimpleNamespace(pivot_name=phase.Name)])
                self.assertEqual(output.Formula, f"=IF($C$42=0,0,{expected})")
                self.assertEqual(output.NumberFormat, "0.00%")


class IdeListRulesTests(unittest.TestCase):
    def raw(self, **values):
        row = dict.fromkeys(REQUIRED)
        row.update({"Key": "IDE-1", "Quote ID": "Q1", "Quote Name": "New connection",
                    "Status": "SO Complete", "Target RFS": datetime(2026, 11, 30),
                    "Actual RFS Date": datetime(2026, 11, 29), "Revenue": 0})
        row.update(values)
        return row

    def reference(self, **values):
        return {"Quote ID": "Q1", "Status": "3.Delivery", "ORDER TYPE": "New Registration", **values}

    def build(self, rows, previous=None, day=date(2026, 11, 17)):
        return build_rows(rows, previous or [self.reference()], day)[0]

    def test_first_quote_history_preserves_every_unique_key_and_current_status(self):
        rows = [self.raw(), self.raw(Key="IDE-2", Status="Cancel")]
        frame = self.build(rows, [self.reference(), self.reference(Status="5.SO Complete", **{"ORDER TYPE": "Renewal"})])
        self.assertEqual(frame["Key"].tolist(), ["IDE-1", "IDE-2"])
        self.assertEqual(frame["Yesterday Status "].tolist(), ["3.Delivery", "3.Delivery"])
        self.assertEqual(frame["Status"].tolist(), ["5.SO Complete", "6.Cancel"])
        self.assertEqual(frame["ORDER TYPE"].tolist(), ["New Registration"] * 2)

    def test_new_renewal_quote_does_not_require_a_history_match(self):
        for title in ("Renewal product", "RENEWEL LICENCE ZOOM", "Renew contract"):
            with self.subTest(title=title):
                row = self.build([self.raw(**{"Quote ID": "NEW", "Quote Name": title})]).iloc[0]
                self.assertEqual(row["ORDER TYPE"], "Renewal")
                self.assertEqual(row["PROCESS"], "Non New Reg")
                self.assertEqual(row["Yesterday Status "], "#N/A")

    def test_raw_revenue_and_missing_dates_do_not_fall_back_to_history(self):
        frame = self.build([self.raw(Revenue=None, **{"SO Completion Date": None})],
                           [self.reference(Revenue=999, **{"SO Completion Date": datetime(2026, 11, 1)})])
        self.assertIsNone(frame.iloc[0]["Revenue"])
        self.assertEqual(frame.iloc[0]["MONTH COMPLETE"], "Jan-00")
        self.assertEqual(frame.iloc[0]["YEAR FAB UPLOAD"], "1900")
        self.assertEqual(frame.iloc[0]["Status"], "5.SO Complete")

    def test_aging_uses_report_date_and_target_not_actual(self):
        for difference, status in [(-6, "On Track"), (-5, "Potential Delay"), (-1, "Potential Delay"), (0, "Delay"), (1, "Delay")]:
            from datetime import timedelta
            row = self.build([self.raw(**{"Target RFS": datetime(2026, 11, 17) - timedelta(days=difference)})]).iloc[0]
            self.assertEqual(row["Aging Of RFS"], difference)
            self.assertEqual(row["STATUS ORDER"], status)

    def test_valid_target_is_not_erased_by_missing_fab_or_actual(self):
        row = self.build([self.raw(**{"FAB Upload Date": None, "Actual RFS Date": None})]).iloc[0]
        self.assertEqual(row[target_header(date(2026, 11, 17))], "Target November")
        self.assertIsNone(row["ON TIME DELIVERY"])

    def test_source_month_is_separate_from_generated_month_and_new_fields_survive(self):
        row = self.build([self.raw(**{"Month Complete": "raw-label", "SO Completion Date": datetime(2026, 10, 31),
                                     "Future Field": "kept"})]).iloc[0]
        self.assertEqual(row["Month Complete2"], "raw-label")
        self.assertEqual(row["MONTH COMPLETE"], "Oct-26")
        self.assertEqual(row["Future Field"], "kept")

    def test_raw_order_type_and_sd_note_append_without_changing_report_fields(self):
        baseline = self.build([self.raw(), self.raw(Key="IDE-2")])
        for raw_header in ("order type", "ORDER TYPE", "Order Type"):
            with self.subTest(raw_header=raw_header):
                frame = self.build([
                    self.raw(**{raw_header: "New-reg", "SD Note": "Raw note"}),
                    self.raw(Key="IDE-2", **{raw_header: None, "SD Note": None}),
                ])
                self.assertEqual(list(frame.columns), list(baseline.columns) + ["Order Type (Raw)", "SD Note"])
                self.assertTrue(frame[baseline.columns].equals(baseline))
                self.assertEqual(frame["Order Type (Raw)"].tolist(), ["New-reg", None])
                self.assertEqual(frame["SD Note"].tolist(), ["Raw note", None])

    def test_duplicate_output_header_fails_before_excel_can_rename_it(self):
        with self.assertRaisesRegex(ValueError, "Order Type \\(Raw\\).*conflicts"):
            self.build([self.raw(**{"order type": "New-reg", "Order Type (Raw)": "different"})])

    def test_unusual_source_year_is_reported_not_guessed(self):
        frame, warnings = build_rows([self.raw(**{"FAB Upload Date": datetime(2926, 7, 28)})],
                                     [self.reference()], date(2026, 11, 17))
        self.assertEqual(frame.iloc[0]["YEAR FAB UPLOAD"], "2926")
        self.assertEqual(warnings[0]["field"], "FAB Upload Date")

    def test_duplicate_key_unknown_status_and_text_dates_fail_explicitly(self):
        for rows, message in [([self.raw(), self.raw()], "unique"),
                              ([self.raw(Status="mystery")], "Unrecognized"),
                              ([self.raw(**{"Target RFS": "11/12/2026"})], "Excel date")]:
            with self.subTest(message=message), self.assertRaisesRegex(ValueError, message):
                self.build(rows)

    def test_targets_follow_calendar_month_including_year_rollover(self):
        self.assertEqual(target_bucket(datetime(2026, 12, 31), date(2027, 1, 1)), "Before January")
        self.assertEqual(target_bucket(datetime(2027, 1, 1), date(2027, 1, 1)), "Target January")
        self.assertEqual(target_bucket(None, date(2027, 1, 1)), "Target Not Yet Inputted")

    def test_all_six_statuses_are_idempotent(self):
        for original, expected in [("Quotation", "1.Quotation"), ("Pra-Delivery", "2.Pra-Delivery"),
                                   ("Delivery", "3.Delivery"), ("Waiting AOS", "4.Waiting AOS"),
                                   ("SO Complete", "5.SO Complete"), ("Cancel", "6.Cancel")]:
            self.assertEqual(normalize_status(original), expected)
            self.assertEqual(normalize_status(expected), expected)

    def test_iso_filename_and_dispatch_without_collabs(self):
        self.assertEqual(determine_report_date(None, Path("IDE-list-2026-11-17.xlsx")), date(2026, 11, 17))
        files = InputFiles(Path("raw.xlsx"), Path("reference.xlsx"))
        with patch("ide_list_format.is_list_export", return_value=True), \
             patch("ide_list_format.generate", return_value={"format_profile": "ide-list"}) as generate:
            self.assertEqual(generate_ide_tracking(files, Path("output.xlsx"), date(2026, 11, 17))["format_profile"], "ide-list")
            generate.assert_called_once()
        with patch("ide_list_format.is_list_export", return_value=False), \
             self.assertRaisesRegex(ValueError, "requires.*CSV"):
            generate_ide_tracking(files, Path("output.xlsx"), date(2026, 11, 17))

    def test_detects_content_and_preserves_first_data_row(self):
        row = self.raw()
        sheet = SimpleNamespace(iter_rows=lambda **_: iter([tuple(row), tuple(row.values())]))
        class Book:
            sheetnames = ["List"]
            def __getitem__(self, name):
                return sheet
            def close(self):
                pass
        with patch("ide_list_format.load_workbook", side_effect=lambda *a, **k: Book()):
            self.assertTrue(is_list_export(Path("renamed.xlsx")))
            self.assertEqual(read_records(Path("renamed.xlsx"), "List"), [row])

    def test_generation_failure_does_not_publish_partial_output(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            reference, output = root / "reference.xlsx", root / "output.xlsx"
            reference.write_bytes(b"template")
            output.write_bytes(b"existing-output")
            with patch("ide_list_format.read_records", side_effect=[[self.raw()], [self.reference()]]), \
                 patch("ide_list_format._write_native", side_effect=RuntimeError("native refresh failed")), \
                 self.assertRaisesRegex(RuntimeError, "native refresh failed"):
                generate(root / "raw.xlsx", reference, output, date(2026, 11, 17))
            self.assertEqual(output.read_bytes(), b"existing-output")
            self.assertEqual(reference.read_bytes(), b"template")
            self.assertFalse(list(root.glob("ide_list_*.xlsx")))


class IdeListPresentationTests(unittest.TestCase):
    def test_raw_order_type_presentation_is_separate_from_report_order_type(self):
        raw_key = presentation_key("order type", raw=True)
        source = {raw_key: ColumnPresentation("@")}
        models = column_presentations(["ORDER TYPE", "Order Type (Raw)"], {}, source)
        self.assertEqual(models["ORDER TYPE"].number_format, "General")
        self.assertEqual(models["Order Type (Raw)"].number_format, "@")

    def test_number_formats_change_display_not_values(self):
        frame, _ = build_rows([IdeListRulesTests().raw(Progress=.5, **{"Revenue": 1234.5})],
                              [IdeListRulesTests().reference()], date(2026, 11, 17))
        formats = column_presentations(frame.columns, {}, {})
        self.assertEqual(formats["Progress"].number_format, "0.0%")
        self.assertEqual(formats["Aging Of RFS"].number_format, "0")
        self.assertIn("#,##0.00", formats["Revenue"].number_format)
        self.assertEqual(frame.iloc[0]["Progress"], .5)
        self.assertEqual(frame.iloc[0]["Revenue"], 1234.5)

    def test_dates_keep_timestamps_but_use_manual_display_format(self):
        models = column_presentations(["Updated", "SO Completion Date"], {
            "updated": ColumnPresentation("dd/mm/yyyy hh:mm"),
        }, {})
        self.assertEqual(models["Updated"].number_format, r"d\ mmm\ yyyy")
        self.assertEqual(models["SO Completion Date"].number_format, r"d\ mmm\ yyyy")

    def test_reference_alignment_is_bound_by_header_and_history_alias(self):
        alignment = ("center", "top", False, False, 0, 0)
        model = ColumnPresentation("General", alignment)
        models = column_presentations(["Progress status", "Yesterday Status "],
                                       {"status": model, "progress status": model}, {})
        self.assertEqual(models["Progress status"].alignment, alignment)
        self.assertEqual(models["Yesterday Status "].alignment, alignment)
        self.assertEqual(presentation_key("Month Complete3"), presentation_key("Month Complete2"))
        self.assertEqual(presentation_key("Month Complete", raw=True), presentation_key("Month Complete2"))
        self.assertNotEqual(presentation_key("MONTH COMPLETE"), presentation_key("Month Complete2"))
        self.assertEqual(presentation_key("TARGET  Detemined as 1 Nov 26"),
                         presentation_key("TARGET  Detemined as 1 Dec 26"))

    def test_unknown_columns_inherit_raw_presentation(self):
        source = ColumnPresentation("0.00%", ("right", "top", True, False, 0, 0))
        model = column_presentations(["Future Rate"], {}, {"future rate": source})["Future Rate"]
        self.assertEqual(model, source)

    def test_sparse_progress_ignores_unformatted_blank_first_rows(self):
        alignment = SimpleNamespace(horizontal=None, vertical="top", wrapText=None,
                                    shrinkToFit=None, indent=0, textRotation=0)
        def cell(value, fmt="General"):
            return SimpleNamespace(value=value, number_format=fmt, alignment=alignment)
        sheet = SimpleNamespace(iter_rows=lambda: iter([
            [cell("Progress")], [cell(None)], [cell(.5, "0.0%")], [cell(.75, "0.0%")],
        ]))
        class Book:
            def __getitem__(self, name):
                return sheet
            def close(self):
                pass
        with patch("ide_list_format.load_workbook", return_value=Book()):
            model = read_column_presentations(Path("reference.xlsx"), "ALL ORDER")["progress"]
        self.assertEqual(model.number_format, "0.0%")

    def test_financial_columns_use_consistent_reference_accounting_format(self):
        financial = ColumnPresentation('_(* #,##0.00_);_(@_)')
        models = column_presentations(["Revenue", "Revenue (Juta)", "Actual Cost", "CEI"],
                                       {"revenue": financial, "revenue (juta)": ColumnPresentation()}, {})
        self.assertEqual(models["Revenue (Juta)"].number_format, financial.number_format)
        self.assertEqual(models["Actual Cost"].number_format, financial.number_format)
        self.assertEqual(models["CEI"].number_format, "General")

    def test_native_alignment_maps_excel_constants(self):
        target = SimpleNamespace()
        apply_column_alignment(target, ("center", "top", True, False, 0, 0))
        self.assertEqual(target.HorizontalAlignment, -4108)
        self.assertEqual(target.VerticalAlignment, -4160)
        self.assertTrue(target.WrapText)

    def test_correct_alignment_does_not_trigger_repeated_excel_layout_updates(self):
        class Target:
            def __init__(self):
                object.__setattr__(self, "writes", [])
                for name, value in {"HorizontalAlignment": 1, "VerticalAlignment": -4160,
                                    "WrapText": True, "ShrinkToFit": False, "IndentLevel": 0,
                                    "Orientation": 0}.items():
                    object.__setattr__(self, name, value)
            def __setattr__(self, name, value):
                self.writes.append(name)
                object.__setattr__(self, name, value)
        target = Target()
        apply_column_alignment(target, (None, "top", True, False, 0, 0))
        self.assertEqual(target.writes, [])

    def test_generated_helper_wrap_matches_both_manual_samples(self):
        wrapped = ColumnPresentation(alignment=(None, "top", True, False, 0, 0))
        models = column_presentations(["Aging Of RFS", "STATUS ORDER", "ON TIME DELIVERY",
                                       "MONTH COMPLETE", "ORDER TYPE", "PROCESS"],
                                      {presentation_key(name): wrapped for name in
                                       ("Aging Of RFS", "STATUS ORDER", "ON TIME DELIVERY", "MONTH COMPLETE")}, {})
        for name in ("Aging Of RFS", "STATUS ORDER", "ON TIME DELIVERY", "MONTH COMPLETE"):
            self.assertFalse(models[name].alignment[2])
        self.assertTrue(models["ORDER TYPE"].alignment[2])
        self.assertTrue(models["PROCESS"].alignment[2])

    def test_long_unwrapped_names_expand_without_changing_wrapped_columns(self):
        import pandas as pd
        columns = {1: SimpleNamespace(ColumnWidth=12), 2: SimpleNamespace(ColumnWidth=20)}
        sheet = SimpleNamespace(Columns=lambda index: columns[index])
        body = SimpleNamespace(Columns=lambda index: SimpleNamespace(
            AutoFit=lambda: setattr(columns[index], "ColumnWidth", 26)))
        frame = pd.DataFrame({"PM": ["Muhammad Mahdi Ramadhan"], "Quote Name": ["a" * 100]})
        models = {"PM": ColumnPresentation(),
                  "Quote Name": ColumnPresentation(alignment=(None, "top", True, False, 0, 0))}
        fit_unwrapped_text_columns(sheet, body, frame, models)
        self.assertEqual(columns[1].ColumnWidth, 26)
        self.assertEqual(columns[2].ColumnWidth, 20)

    def test_autofit_never_shrinks_reference_width(self):
        import pandas as pd
        column = SimpleNamespace(ColumnWidth=20)
        sheet = SimpleNamespace(Columns=lambda index: column)
        body = SimpleNamespace(Columns=lambda index: SimpleNamespace(
            AutoFit=lambda: setattr(column, "ColumnWidth", 15)))
        frame = pd.DataFrame({"PM": ["a" * 30]})
        fit_unwrapped_text_columns(sheet, body, frame, {"PM": ColumnPresentation()})
        self.assertEqual(column.ColumnWidth, 20)

    def test_all_order_headers_center_without_changing_body_alignment(self):
        body = ("left", "top", True, False, 0, 0)
        heading = ("left", "top", True, False, 2, 0)
        reference = ColumnPresentation(alignment=body, header_alignment=heading)
        models = column_presentations(["Quote ID", "Future Field"],
                                       {"quote id": reference}, {"future field": reference})
        for model in models.values():
            self.assertEqual(model.header_alignment, ("center", "center", True, False, 0, 0))
            self.assertEqual(model.alignment, body)


if __name__ == "__main__":
    unittest.main()
