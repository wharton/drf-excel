import datetime as dt
from types import SimpleNamespace

import pytest
from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

from drf_excel.utilities import (
    XLSXStyle,
    get_attribute,
    get_setting,
    sanitize_value,
    set_cell_style,
)


class TestXLSXStyle:
    def test_no_style_dict(self):
        style = XLSXStyle()
        assert style.font is None
        assert style.fill is None
        assert style.alignment is None
        assert style.number_format is None
        assert style.border is None

    def test_with_font_style(self):
        style = XLSXStyle(
            {
                "font": {"name": "Arial", "size": 14, "bold": True},
            }
        )

        assert isinstance(style.font, Font)
        assert style.font.name == "Arial"
        assert style.font.size == 14
        assert style.font.bold is True

    def test_with_fill_style(self):
        style = XLSXStyle(
            {
                "fill": {
                    "fill_type": "solid",
                    "start_color": "FFCCFFCC",
                },
            }
        )

        assert isinstance(style.fill, PatternFill)
        assert style.fill.fill_type == "solid"
        assert isinstance(style.fill.start_color, Color)
        assert style.fill.start_color.value == "FFCCFFCC"
        assert style.fill.start_color.type == "rgb"

    def test_with_alignment_style(self):
        style = XLSXStyle(
            {
                "alignment": {
                    "horizontal": "center",
                    "wrap_text": True,
                    "text_rotation": 20,
                },
            }
        )

        assert isinstance(style.alignment, Alignment)
        assert style.alignment.horizontal == "center"
        assert style.alignment.wrap_text is True
        assert style.alignment.text_rotation == 20

    def test_with_border_style(self):
        style = XLSXStyle(
            {"border_side": {"border_style": "thin", "color": "FF000000"}}
        )

        assert isinstance(style.border, Border)
        assert isinstance(style.border.left, Side)
        assert style.border.left.border_style == "thin"
        assert isinstance(style.border.left.color, Color)
        assert style.border.left.color.value == "FF000000"
        assert style.border.left.color.type == "rgb"
        assert (
            style.border.left
            == style.border.right
            == style.border.top
            == style.border.bottom
        )

    def test_with_number_format_style(self):
        style = XLSXStyle({"format": "0.00E+00"})

        assert style.number_format == "0.00E+00"


class TestGetAttribute:
    def test_existing(self):
        obj = SimpleNamespace(a=1)
        assert get_attribute(obj, "a") == 1

    def test_non_existing(self):
        obj = SimpleNamespace(a=1)
        assert get_attribute(obj, "b") is None

    def test_non_existing_with_default(self):
        obj = SimpleNamespace(a=1)
        assert get_attribute(obj, "b", default="c") == "c"

    def test_from_getter_method(self):
        class Foo:
            def get_b(self):
                return 1

        obj = Foo()
        assert get_attribute(obj, "b") == 1


class TestGetSetting:
    def test_not_defined(self):
        assert get_setting("DUMMY") is None

    def test_not_defined_with_default(self):
        assert get_setting("DUMMY", "default") == "default"

    def test_defined(self, settings):
        settings.DRF_EXCEL_DUMMY = "custom-value"
        assert get_setting("DUMMY") == "custom-value"


@pytest.mark.parametrize(
    ("value", "expected_output"),
    [
        # Regular values without illegal characters
        (None, None),
        ("test", "test"),
        (True, "True"),
        (False, False),  # Bug?
        (1, "1"),
        (dt.date(2020, 1, 1), "2020-01-01"),
        (dt.datetime(2020, 1, 1, 1, 1, 1), "2020-01-01 01:01:01"),
        # With illegal characters
        ("test\000", "test"),
        ("t\001est", "test"),
        ("test\005er", "tester"),
        ("tes\010t", "test"),
        ("test\013 me", "test me"),
        ("test\014", "test"),
        ("test\016", "test"),
        ("test\020", "test"),
        ("test\030", "test"),
        ("test\037", "test"),
        # Escape if starts with these characters
        ("=test", "'=test"),
        ("-example", "'-example"),
        ("+foo", "'+foo"),
        ("@bar", "'@bar"),
        ("\tbaz", "'\tbaz"),
        ("\nqux", "'\nqux"),
        ("\rquux", "'\rquux"),
    ],
)
def test_sanitize_value(value, expected_output):
    assert sanitize_value(value) == expected_output


class TestSetCellStyle:
    @pytest.fixture
    def cell(self, worksheet: Worksheet):
        return Cell(worksheet)

    def test_no_style(self, cell):
        set_cell_style(cell, None)
        assert cell.font is not None

    def test_with_styles(self, cell):
        style = XLSXStyle(
            {
                "font": {"name": "Arial"},
                "fill": {"fill_type": "solid"},
                "alignment": {"horizontal": "center"},
                "border_side": {"border_style": "thin"},
                "format": "0.00",
            },
        )
        set_cell_style(cell, style)
        assert cell.font.name == "Arial"
        assert cell.fill.fill_type == "solid"
        assert cell.alignment.horizontal == "center"
        assert cell.border.left.border_style == "thin"
        assert cell.number_format == "0.00"

    def test_with_partial_styles(self, cell):
        style = XLSXStyle({"font": {"name": "Arial"}})
        set_cell_style(cell, style)
        assert cell.font.name == "Arial"
        assert cell.number_format == "General"
