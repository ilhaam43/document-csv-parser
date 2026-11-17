import io
import unittest
from types import SimpleNamespace
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from ide_list_format import read_records, validate_progress_bar
from ide_workbook_package import S, R, conditional_formats, relationship_parts, worksheet_part, validate_table_and_caches


P = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def element(tag, **attributes):
    return ET.Element(S + tag, attributes)


def relationships(items, kind=None):
    root = ET.Element(P + "Relationships")
    for identifier, target in items:
        item = ET.SubElement(root, P + "Relationship", Id=identifier, Target=target)
        if kind:
            item.set("Type", R[1:-1] + "/" + kind)
    return root


def package(alter=None):
    workbook = element("workbook")
    sheets = ET.SubElement(workbook, S + "sheets")
    for name, identifier in (("ALL ORDER", "data"), ("PIVOT", "pivot")):
        ET.SubElement(sheets, S + "sheet", {"name": name, R + "id": identifier})
    caches = ET.SubElement(workbook, S + "pivotCaches")
    ET.SubElement(caches, S + "pivotCache", {"cacheId": "7", R + "id": "cache"})
    data = element("worksheet")
    tables = ET.SubElement(data, S + "tableParts")
    ET.SubElement(tables, S + "tablePart", {R + "id": "table"})
    pivot = element("worksheet")
    pivots = ET.SubElement(pivot, S + "pivotTableParts")
    for identifier in ("one", "two"):
        ET.SubElement(pivots, S + "pivotTablePart", {R + "id": identifier})
    cache = element("pivotCacheDefinition", recordCount="2")
    source = ET.SubElement(cache, S + "cacheSource")
    ET.SubElement(source, S + "worksheetSource", name="Table1")
    parts = {
        "xl/workbook.xml": workbook,
        "xl/_rels/workbook.xml.rels": relationships([
            ("data", "worksheets/data.xml"), ("pivot", "/xl/worksheets/pivot.xml"), ("cache", "pivotCache/cache.xml")]),
        "xl/worksheets/data.xml": data,
        "xl/worksheets/_rels/data.xml.rels": relationships([("table", "../tables/table1.xml")]),
        "xl/worksheets/pivot.xml": pivot,
        "xl/worksheets/_rels/pivot.xml.rels": relationships([
            ("one", "../pivotTables/one.xml"), ("two", "../pivotTables/two.xml")], kind="pivotTable"),
        "xl/pivotTables/one.xml": element("pivotTableDefinition", name="One", cacheId="7"),
        "xl/pivotTables/two.xml": element("pivotTableDefinition", name="Two", cacheId="7"),
        "xl/tables/table1.xml": element("table", name="Table1", ref="A1:C3"),
        "xl/pivotCache/cache.xml": cache,
    }
    if alter:
        alter(parts)
    output = io.BytesIO()
    with ZipFile(output, "w") as archive:
        for name, root in parts.items():
            archive.writestr(name, ET.tostring(root))
    output.seek(0)
    return ZipFile(output)


class IdePackageTests(unittest.TestCase):
    def validate(self, archive, count=2):
        data_part = worksheet_part(archive, "ALL ORDER")
        validate_table_and_caches(archive, ET.fromstring(archive.read(data_part)), data_part, "A1:C3", 2, count)

    def test_resolve_relative_and_absolute_sheet_parts(self):
        with package() as archive:
            self.assertEqual(worksheet_part(archive, "ALL ORDER"), "xl/worksheets/data.xml")
            self.assertEqual(worksheet_part(archive, "PIVOT"), "xl/worksheets/pivot.xml")
            self.assertEqual(relationship_parts(archive, "xl/worksheets/data.xml")["table"], "xl/tables/table1.xml")

    def test_missing_sheet_fails_explicitly(self):
        with package() as archive, self.assertRaisesRegex(ValueError, "Missing"):
            worksheet_part(archive, "Missing")

    def test_table_and_each_pivot_cache_are_validated_without_cache_records(self):
        with package() as archive, patch.object(archive, "read", wraps=archive.read) as read:
            self.validate(archive)
            self.assertEqual(sum(call.args[0] == "xl/pivotCache/cache.xml" for call in read.call_args_list), 1)
            self.assertTrue(all("pivotCacheRecords" not in call.args[0] for call in read.call_args_list))

    def test_stale_table_range_or_missing_table_name_fails(self):
        for attribute, value in (("ref", "A1:C4"), ("name", "Table2")):
            with self.subTest(attribute=attribute), package(
                    lambda parts: parts["xl/tables/table1.xml"].set(attribute, value)) as archive, \
                    self.assertRaisesRegex(RuntimeError, "table range"):
                self.validate(archive)

    def test_missing_pivot_or_wrong_pivot_count_fails(self):
        with package() as archive, self.assertRaisesRegex(RuntimeError, "pivots were lost"):
            self.validate(archive, count=3)

    def test_excel_links_pivots_through_relationships_without_a_sheet_manifest(self):
        def alter(parts):
            root = parts["xl/worksheets/pivot.xml"]
            root.remove(root.find(S + "pivotTableParts"))
            ET.SubElement(parts["xl/worksheets/_rels/pivot.xml.rels"], P + "Relationship",
                          Id="printer", Target="../printerSettings/printer.bin", Type=R[1:-1]+"/printerSettings")
        with package(alter) as archive:
            self.validate(archive)

    def test_wrong_cache_source_or_record_count_fails(self):
        for source, count in (("OldTable", "2"), ("Table1", "1")):
            def alter(parts):
                cache = parts["xl/pivotCache/cache.xml"]
                cache.set("recordCount", count)
                cache.find(f"{S}cacheSource/{S}worksheetSource").set("name", source)
            with self.subTest(source=source, count=count), package(alter) as archive, \
                    self.assertRaisesRegex(RuntimeError, "not bound to current"):
                self.validate(archive)

    def test_unregistered_cache_on_second_pivot_is_not_skipped(self):
        with package(lambda parts: parts["xl/pivotTables/two.xml"].set("cacheId", "8")) as archive, \
                self.assertRaisesRegex(RuntimeError, "Two has no registered cache"):
            self.validate(archive)

    def test_each_distinct_cache_is_checked(self):
        def alter(parts):
            parts["xl/pivotTables/two.xml"].set("cacheId", "8")
            ET.SubElement(parts["xl/workbook.xml"].find(S + "pivotCaches"), S + "pivotCache",
                          {"cacheId": "8", R + "id": "other"})
            ET.SubElement(parts["xl/_rels/workbook.xml.rels"], P + "Relationship",
                          Id="other", Target="pivotCache/other.xml")
            parts["xl/pivotCache/other.xml"] = element("pivotCacheDefinition", recordCount="1")
        with package(alter) as archive, self.assertRaisesRegex(RuntimeError, "Two is not bound"):
            self.validate(archive)

    def test_saved_conditional_formats_keep_progress_rule_validation(self):
        root = element("worksheet")
        area = ET.SubElement(root, S + "conditionalFormatting", sqref="J2:J2076")
        rule = ET.SubElement(area, S + "cfRule", type="dataBar", priority="1")
        data_bar = ET.SubElement(rule, S + "dataBar")
        for value in ("0", "1"):
            ET.SubElement(data_bar, S + "cfvo", type="num", val=value)
        ET.SubElement(data_bar, S + "color", rgb="FF7FD3CA")
        validate_progress_bar(SimpleNamespace(conditional_formatting=conditional_formats(root)), 10, 2075)
        area.set("sqref", "AT2:AT2076")
        with self.assertRaisesRegex(RuntimeError, "must cover"):
            validate_progress_bar(SimpleNamespace(conditional_formatting=conditional_formats(root)), 10, 2075)

    def test_combined_input_reader_populates_records_and_presentation_once(self):
        expected = [{"Quote ID": "Q1", "Progress": .81}]
        styles = {"progress": object()}
        def read(path, sheet, *, records):
            records.extend(expected)
            return styles
        presentations = {}
        with patch("ide_list_format.read_column_presentations", side_effect=read) as reader:
            self.assertEqual(read_records("raw.xlsx", "List", presentations=presentations), expected)
            reader.assert_called_once()
        self.assertEqual(presentations, styles)


if __name__ == "__main__":
    unittest.main()
