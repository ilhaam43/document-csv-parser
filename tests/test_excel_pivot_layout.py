import unittest
from datetime import date

from types import SimpleNamespace
from unittest.mock import Mock, patch

from excel_pivot_layout import (
    align_completion_threshold_legends,
    apply_completion_icon_thresholds,
    completion_legend_values,
    completion_mode_for_row,
    completion_section_modes,
    completion_percentage_headers,
    percentage_block_end_row,
    percentage_format_row_spans,
    section_banner_end_columns,
    update_dynamic_pivot_section_titles,
)
from completion_threshold_rules import completion_thresholds_for


class FakeIconCondition:
    def __init__(self, kind=6, icon_set=4):
        self.Type = kind
        self.IconSet = SimpleNamespace(ID=icon_set)
        self.criteria = {i: SimpleNamespace(Type=None, Value=None, Operator=None) for i in (2, 3)}
        self.IconCriteria = self.criteria.__getitem__


class FakeFormatConditions:
    def __init__(self, *conditions):
        self.conditions = list(conditions)
        self.created = 0

    @property
    def Count(self):
        return len(self.conditions)

    def __call__(self, index):
        return self.conditions[index - 1]

    def AddIconSetCondition(self):
        self.created += 1
        result = FakeIconCondition(icon_set=1)
        self.conditions.append(result)
        return result


class CompletionPercentageFormattingTests(unittest.TestCase):
    def target(self, *conditions):
        return SimpleNamespace(
            FormatConditions=FakeFormatConditions(*conditions),
            Worksheet=SimpleNamespace(Parent=SimpleNamespace(IconSets=lambda key: SimpleNamespace(ID=key))),
        )

    def test_existing_traffic_light_is_updated_without_adding_duplicates(self):
        rule = FakeIconCondition()
        target = self.target(rule)
        for day, yellow, green in ((2, 0.2, 0.3), (14, 0.4, 0.5)):
            self.assertEqual(apply_completion_icon_thresholds(
                target, completion_thresholds_for(date(2027, 3, day)), create_if_missing=True,
            ), 1)
            self.assertEqual(rule.criteria[2].Value, yellow)
            self.assertEqual(rule.criteria[3].Value, green)
            self.assertEqual(rule.criteria[2].Type, 0)
            self.assertEqual(rule.criteria[3].Type, 0)
            self.assertEqual(rule.criteria[2].Operator, 7)
            self.assertEqual(rule.criteria[3].Operator, 5)
        self.assertEqual(target.FormatConditions.created, 0)

    def test_missing_rule_recovery_is_explicit_and_preserves_other_rules(self):
        other_rule = FakeIconCondition(kind=1)
        arrows = FakeIconCondition(icon_set=1)
        target = self.target(other_rule, arrows)
        thresholds = completion_thresholds_for(date(2027, 3, 2), mode='target_after')
        self.assertEqual(apply_completion_icon_thresholds(target, thresholds), 0)
        self.assertEqual(apply_completion_icon_thresholds(target, thresholds, create_if_missing=True), 1)
        self.assertEqual(target.FormatConditions.conditions[:2], [other_rule, arrows])
        rule = target.FormatConditions.conditions[-1]
        self.assertEqual(rule.IconSet.ID, 4)
        self.assertFalse(rule.ReverseOrder)
        self.assertFalse(rule.ShowIconOnly)
        self.assertEqual(rule.criteria[2].Value, 0.05)
        self.assertEqual(rule.criteria[3].Value, 0.1)
        apply_completion_icon_thresholds(target, thresholds, create_if_missing=True)
        self.assertEqual(target.FormatConditions.created, 1)
        self.assertIsNone(arrows.criteria[2].Value)

    def test_com_failure_to_create_or_update_a_rule_is_not_silently_ignored(self):
        thresholds = completion_thresholds_for(date(2027, 3, 2))
        target = self.target()
        target.FormatConditions.AddIconSetCondition = Mock(side_effect=RuntimeError('cannot create'))
        with self.assertRaisesRegex(RuntimeError, 'cannot create'):
            apply_completion_icon_thresholds(target, thresholds, create_if_missing=True)
        rule = FakeIconCondition()
        rule.IconCriteria = Mock(side_effect=RuntimeError('cannot update'))
        target = self.target(rule)
        with self.assertRaisesRegex(RuntimeError, 'cannot update'):
            apply_completion_icon_thresholds(target, thresholds, create_if_missing=True)
        self.assertEqual(target.FormatConditions.created, 0)

    def test_fragmented_existing_rules_all_receive_the_same_business_thresholds(self):
        target = self.target(FakeIconCondition(), FakeIconCondition())
        self.assertEqual(apply_completion_icon_thresholds(
            target, completion_thresholds_for(date(2027, 3, 2)), create_if_missing=True,
        ), 2)
        self.assertEqual([r.criteria[2].Value for r in target.FormatConditions.conditions], [0.2, 0.2])
        self.assertEqual(target.FormatConditions.created, 0)

    def test_shifted_source_headers_keep_section_order_and_column_offsets(self):
        header = 'Percentage of Completion (SO Complete, Cancel & Change Target)'
        sheet = SimpleNamespace(UsedRange=SimpleNamespace(
            Row=8, Column=5, Value2=((None, header), (None, None),
                                    (header, None), ('Percentage of Order On Going', None),
                                    ('Percentage\nof Completion', None)),
        ))
        self.assertEqual(completion_percentage_headers(sheet), [(8, 6), (10, 5), (12, 5)])

    def test_empty_source_has_no_completion_headers(self):
        self.assertEqual(completion_percentage_headers(SimpleNamespace(
            UsedRange=SimpleNamespace(Value2=None, Row=1, Column=1),
        )), [])

    def test_empty_percentage_block_does_not_capture_a_later_table(self):
        sheet = SimpleNamespace(Cells=lambda row, column: SimpleNamespace(
            Formula='', Value2=None if row < 20 else 'another section',
        ))
        self.assertEqual(percentage_block_end_row(sheet, 10, 8, 30), 10)

    def test_block_includes_formula_without_cache_and_zero_total_but_stops_at_gap(self):
        values = {11: ('=F11/C11', None), 12: ('=0', 0), 14: ('=0', 0)}
        sheet = SimpleNamespace(Cells=lambda row, column: SimpleNamespace(
            Formula=values.get(row, ('', None))[0], Value2=values.get(row, ('', None))[1],
        ))
        self.assertEqual(percentage_block_end_row(sheet, 10, 8, 30), 12)

    def test_grand_total_only_never_overwrites_header_or_formats_a_phantom_detail(self):
        self.assertEqual(percentage_format_row_spans(114, 115), [
            ('header', 114, 114), ('grand_total', 115, 115),
        ])
        self.assertEqual(percentage_format_row_spans(114, 114), [('header', 114, 114)])

    def test_detail_roles_cover_each_row_once_up_to_the_actual_total(self):
        for count in range(1, 9):
            spans = percentage_format_row_spans(40, 40 + count)
            self.assertEqual([r for _, start, end in spans for r in range(start, end + 1)],
                             list(range(40, 41 + count)))
            self.assertEqual(spans[-1][0], 'grand_total')

    def test_legend_alignment_recovers_missing_rule_at_the_dynamic_block(self):
        values = (('TARGET COMPLETE MARCH 2027', None), ('Percentage of Completion', None), (0, None))
        sheet = Mock()
        sheet.UsedRange = SimpleNamespace(Value2=values, Row=10, Column=8, Rows=SimpleNamespace(Count=3))
        sheet.Cells.side_effect = lambda row, col: SimpleNamespace(
            Formula='=0' if (row, col) == (12, 8) else '',
            Value2=0 if (row, col) == (12, 8) else None, Address=f'R{row}C{col}',
        )
        with patch('excel_pivot_layout.apply_completion_icon_thresholds', return_value=1) as apply:
            align_completion_threshold_legends(sheet, date(2027, 3, 2), [(13, 1, 'weekly')])
        self.assertEqual(apply.call_count, 1)
        self.assertEqual(apply.call_args.kwargs, {'create_if_missing': True})
        self.assertEqual(apply.call_args.args[1].green_percent, 30)


class SectionBannerEndColumnsTests(unittest.TestCase):
    def test_completion_legend_uses_the_week_one_threshold_band(self):
        self.assertEqual(
            completion_legend_values(date(2026, 9, 1)),
            ("Green : >30%", "Yellow : >=20%", "Red : <20%"),
        )

    def test_completion_legend_keeps_week_five_for_target_after(self):
        self.assertEqual(
            completion_legend_values(date(2026, 9, 29), "target_after"),
            ("Green : >50%", "Yellow : >=40%", "Red : <40%"),
        )

    def test_target_after_mode_is_detected_from_the_nearest_section(self):
        values = (
            ("TARGET COMPLETE SEPTEMBER 2026",),
            (None,),
            ("Percentage of Completion",),
            ("TARGET AFTER SEPTEMBER 2026",),
            (None,),
            ("Percentage of Completion",),
        )
        section_modes = completion_section_modes(values, 10)

        self.assertEqual(completion_mode_for_row(12, section_modes), "weekly")
        self.assertEqual(
            completion_mode_for_row(15, section_modes),
            "target_after",
        )

    def test_updates_month_titles_without_fixed_cell_addresses(self):
        values = (
            ("TARGET COMPLETE AUGUST 2026", None),
            (None, "ALL TARGET AUGUST"),
        )
        cells = {}

        class Cell:
            def __init__(self, value=None):
                self.Value = value

        class Sheet:
            UsedRange = SimpleNamespace(Value=values, Row=7, Column=3)

            def Cells(self, row, column):
                return cells.setdefault((row, column), Cell())

        sheet = Sheet()
        updated = update_dynamic_pivot_section_titles(
            sheet,
            {
                "target complete": "TARGET COMPLETE SEPTEMBER 2026",
                "all target": "ALL TARGET SEPTEMBER",
            },
        )

        self.assertEqual(updated, 2)
        self.assertEqual(cells[(7, 3)].Value, "TARGET COMPLETE SEPTEMBER 2026")
        self.assertEqual(cells[(8, 4)].Value, "ALL TARGET SEPTEMBER")

    def test_separates_titles_that_share_the_same_row(self):
        titles = [(1, 1), (1, 19), (159, 1)]
        pivot_ranges = [
            (9, 1, 3),
            (8, 5, 10),
            (8, 19, 23),
            (177, 1, 3),
            (176, 5, 11),
        ]
        side_headers = [(9, 11), (176, 12)]

        self.assertEqual(
            section_banner_end_columns(titles, pivot_ranges, side_headers),
            {
                (1, 1): 11,
                (1, 19): 23,
                (159, 1): 12,
            },
        )

    def test_uses_pivot_edge_when_section_has_no_side_calculation(self):
        self.assertEqual(
            section_banner_end_columns(
                titles=[(92, 1), (158, 1)],
                pivot_ranges=[(110, 1, 3), (109, 5, 6), (167, 1, 3), (166, 5, 10)],
                side_headers=[(167, 11)],
            ),
            {
                (92, 1): 6,
                (158, 1): 11,
            },
        )


if __name__ == "__main__":
    unittest.main()
