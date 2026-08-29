import unittest

from csv_to_excel_on_going import (
    DEFAULT_ON_PROGRESS_PIVOT_STYLE,
    resolve_on_progress_pivot_style,
)


class FakePivot:
    def __init__(self, style_name):
        self.TableStyle2 = style_name


class FakeProgressSheet:
    def __init__(self, pivots):
        self._pivots = pivots

    def PivotTables(self, name):
        return self._pivots[name]


class ResolveOnProgressPivotStyleTests(unittest.TestCase):
    def test_uses_first_available_template_pivot_style(self):
        sheet = FakeProgressSheet(
            {
                "missing-style": FakePivot(""),
                "template-pivot": FakePivot("PivotStyleLight16"),
            }
        )

        self.assertEqual(
            resolve_on_progress_pivot_style(sheet, ["missing-style", "template-pivot"]),
            "PivotStyleLight16",
        )

    def test_uses_light_style_fallback_when_template_style_is_unavailable(self):
        sheet = FakeProgressSheet({})

        self.assertEqual(
            resolve_on_progress_pivot_style(sheet, ["missing-pivot"]),
            DEFAULT_ON_PROGRESS_PIVOT_STYLE,
        )


if __name__ == "__main__":
    unittest.main()
