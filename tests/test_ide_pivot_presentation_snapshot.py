import unittest
from collections import Counter
from types import SimpleNamespace
from unittest.mock import patch

from ide_pivot_presentation import CellPresentation, PivotPresentation, _row_identity, display_style, header_style, row_role, select_cell_style, style_ranges, validate_centered_header, pivot_row_identity, read_presentation_grid


class PivotPresentationSnapshotTests(unittest.TestCase):
    def setUp(self):
        self.style = CellPresentation("Indosat Regular", 12, "0", ("center", "center", True, False, 0, 0))
        self.model = PivotPresentation(row_fields=("department", "pm", "apm"))
        self.metric = ("measure", ("quote id", -4112, -4143, (("status", "3.Delivery"),)))

    def test_row_role_uses_field_order_not_alphabetic_order(self):
        identity = (("apm", "A"), ("department", "D"), ("pm", "P"))
        self.assertEqual(row_role(identity, self.model.row_fields), "apm")
        self.assertEqual(row_role((), self.model.row_fields), "__total__")

    def test_existing_row_style_survives_row_position_change(self):
        identity = (("department", "Delivery"),)
        self.model.cells[identity, self.metric] = self.style
        self.assertEqual(select_cell_style(self.model, {}, "phase", identity, self.metric), self.style)

    def test_new_status_category_inherits_same_measure_style(self):
        self.model.defaults[("measure", "quote id", -4112, -4143)][self.style] = 4
        self.assertEqual(select_cell_style(self.model, {}, "phase", (("pm", "New PM"),), self.metric), self.style)

    def test_empty_non_new_reference_borrows_apm_style_from_peer_pivot(self):
        label = ("label", 0)
        peers = {("phase", ("apm", label)): Counter({self.style: 3})}
        self.assertEqual(select_cell_style(self.model, peers, "phase", (("pm", "P"), ("apm", "A")), label), self.style)

    def test_unknown_label_can_use_reference_total_style(self):
        label = ("label", 0)
        self.model.defaults[("__total__", label)][self.style] = 1
        self.assertEqual(select_cell_style(self.model, {}, "phase", (("pm", "P"),), label), self.style)

    def test_otd_percent_does_not_inherit_count_number_format(self):
        percent = ("measure", ("quote id", -4112, 8, ()))
        self.model.defaults[("measure", "quote id", -4112, -4143)][self.style] = 1
        self.assertIsNone(select_cell_style(self.model, {}, "phase", (), percent))

    def test_empty_summary_total_has_no_native_row_items(self):
        pivot = SimpleNamespace(Name="Empty", DataBodyRange=SimpleNamespace(
            Rows=SimpleNamespace(Count=1), Columns=SimpleNamespace(Count=2)))
        sheet = SimpleNamespace(Cells=lambda r, c: SimpleNamespace(Value2=None))
        with patch("ide_pivot_presentation.pivot_row_identity", side_effect=RuntimeError("No RowItems")):
            self.assertEqual(_row_identity(sheet, pivot, 43, 2), ())

    def test_sparse_revenue_uses_the_quote_count_cell_for_row_identity(self):
        pivot = SimpleNamespace(Name="Summary", DataBodyRange=SimpleNamespace(
            Rows=SimpleNamespace(Count=2), Columns=SimpleNamespace(Count=2)))
        sheet = SimpleNamespace(Cells=lambda r, c: SimpleNamespace(Value2=None if c == 2 else 1))
        identity = (("pm", "P"),)
        with patch("ide_pivot_presentation.pivot_row_identity", side_effect=[RuntimeError("Blank"), identity]):
            self.assertEqual(_row_identity(sheet, pivot, 10, 2), identity)

    def test_unknown_nonempty_row_is_not_silently_treated_as_grand_total(self):
        pivot = SimpleNamespace(Name="Summary", DataBodyRange=SimpleNamespace(
            Rows=SimpleNamespace(Count=2), Columns=SimpleNamespace(Count=1)))
        sheet = SimpleNamespace(Cells=lambda r, c: SimpleNamespace(Value2=1))
        with patch("ide_pivot_presentation.pivot_row_identity", side_effect=RuntimeError("Unexpected")), \
             self.assertRaisesRegex(RuntimeError, "Cannot identify"):
            _row_identity(sheet, pivot, 10, 2)

    def test_empty_two_level_reference_uses_own_font_and_relative_indent(self):
        self.model.row_fields = ("pm", "apm")
        label = ("label", 0)
        self.model.cells[(), label] = self.style
        peers = {("phase", ("pm", label)): Counter({CellPresentation(
            "Indosat Regular", 11, "General", ("left", None, False, False, 1, 0)): 3})}
        top = select_cell_style(self.model, peers, "phase", (("pm", "P"),), label)
        child = select_cell_style(self.model, peers, "phase", (("pm", "P"), ("apm", "A")), label)
        self.assertEqual(top.font_size, 12)
        self.assertEqual(top.alignment[4], 0)
        self.assertEqual(child.alignment[4], 1)

    def test_summary_measures_are_centered_without_changing_row_labels(self):
        style = CellPresentation("Indosat Regular", 12, "#,##0.00", (None, None, False, False, 0, 0))
        centered = display_style(style, "summary", self.metric)
        self.assertEqual(centered.alignment[:2], ("center", "center"))
        self.assertEqual(centered.number_format, style.number_format)
        self.assertEqual(display_style(style, "summary", ("label", 0)), style)

    def test_style_ranges_group_consecutive_rows_without_including_other_cells(self):
        ranges = list(style_ranges([(10, 2), (11, 2), (11, 2), (13, 2), (10, 3)]))
        self.assertEqual(ranges, ["B10:B11,B13:B13,C10:C10"])

    def test_style_ranges_bound_multi_area_com_arguments(self):
        ranges = list(style_ranges([(row, 2) for row in range(1, 100, 2)], limit=30))
        self.assertTrue(all(len(address) <= 30 for address in ranges))
        self.assertEqual(sum(address.count(",")+1 for address in ranges), 50)

    def test_new_measure_heading_inherits_centered_fallback(self):
        style = CellPresentation("Indosat Regular", 12, "General", ("left", "top", True, False, 0, 0))
        centered = header_style(style, is_new_measure_heading=True)
        self.assertEqual(centered.alignment, ("center", "center", True, False, 0, 0))
        self.assertEqual(centered.font_name, style.font_name)
        self.assertEqual(centered.font_size, style.font_size)

    def test_existing_pivot_headers_and_filters_keep_reference_alignment(self):
        style = CellPresentation("Indosat Regular", 12, "General", ("left", "top", False, False, 0, 0))
        self.assertEqual(header_style(style, is_new_measure_heading=False), style)
        self.assertIsNone(header_style(None, is_new_measure_heading=True))

    def test_saved_all_order_header_alignment_is_validated_before_publication(self):
        cell = SimpleNamespace(value="Revenue", coordinate="B9", alignment=SimpleNamespace(
            horizontal="center", vertical="center"))
        validate_centered_header(cell)
        cell.alignment.horizontal = "left"
        with self.assertRaisesRegex(RuntimeError, "B9.*not centered"):
            validate_centered_header(cell)
        cell.value = None
        validate_centered_header(cell)

    def test_cached_row_field_names_are_scoped_by_hierarchy_depth(self):
        parent_reads = []
        class Item:
            def __init__(self, field, name):
                self.field, self.Name = field, name
            @property
            def Parent(self):
                parent_reads.append(self.field)
                return SimpleNamespace(SourceName=self.field)
        class Items(list):
            @property
            def Count(self):
                return len(self)
            def __call__(self, index):
                return self[index-1]
        def cell(*items):
            return SimpleNamespace(PivotCell=SimpleNamespace(RowItems=Items(items)))
        cache = {}
        first = pivot_row_identity(cell(Item("Department", "D1"), Item("PM", "P1")), field_names=cache)
        second = pivot_row_identity(cell(Item("Department", "D2"), Item("PM", "P2")), field_names=cache)
        parent = pivot_row_identity(cell(Item("Department", "D2")), field_names=cache)
        self.assertEqual(first, (("department", "D1"), ("pm", "P1")))
        self.assertEqual(second, (("department", "D2"), ("pm", "P2")))
        self.assertEqual(parent, (("department", "D2"),))
        self.assertEqual(parent_reads, ["Department", "PM", "Department"])
        self.assertEqual(pivot_row_identity(cell(), field_names=cache), ())

    def test_successful_identity_does_not_read_values_or_refetch_body_geometry(self):
        pivot = SimpleNamespace(Name="P")
        cell = SimpleNamespace()
        identity = (("pm", "P"),)
        with patch("ide_pivot_presentation.pivot_row_identity", return_value=identity):
            self.assertEqual(_row_identity(None, pivot, 3, 2, data_columns=1, data_rows=2,
                                           cell_at=lambda row, column: cell), identity)

    def test_streamed_styles_keep_empty_styled_cells_and_explicit_row_geometry(self):
        from tests.test_ide_workbook_package import package
        from ide_workbook_package import S
        from xml.etree import ElementTree as ET
        alignment = SimpleNamespace(horizontal="center", vertical="top", wrapText=False,
                                    shrinkToFit=False, indent=0, textRotation=0)
        def cell(row, column, value):
            return SimpleNamespace(row=row, column=column, value=value, has_style=True,
                                   font=SimpleNamespace(name="Indosat Regular", sz=12),
                                   number_format="0.00%", alignment=alignment)
        class Sheet:
            max_row = 3
            def __iter__(self):
                return iter([[cell(1, 1, "Header"), SimpleNamespace(value=None)],
                             [SimpleNamespace(value=None), cell(3, 2, None)]])
        class Book:
            def __getitem__(self, name):
                return Sheet()
            def close(self):
                pass
        def alter(parts):
            root = parts["xl/worksheets/pivot.xml"]
            ET.SubElement(root, S+"sheetFormatPr", defaultRowHeight="19.5")
            rows = ET.SubElement(root, S+"sheetData")
            ET.SubElement(rows, S+"row", r="2", ht="30", hidden="1")
        with package(alter) as archive, patch("ide_pivot_presentation.ZipFile", return_value=archive), \
                patch("ide_pivot_presentation.load_workbook", return_value=Book()) as read:
            grid, rows = read_presentation_grid("reference.xlsx")
        self.assertEqual(set(grid), {(1, 1), (3, 2)})
        self.assertEqual(grid[3, 2].number_format, "0.00%")
        self.assertEqual(rows, {1: (19.5, False), 2: (30, True), 3: (19.5, False)})
        read.assert_called_once_with("reference.xlsx", read_only=True, data_only=True)


if __name__ == "__main__":
    unittest.main()
