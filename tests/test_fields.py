import pytest
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet
from rest_framework.fields import CharField

from drf_excel.fields import XLSXField
from drf_excel.utilities import XLSXStyle


@pytest.fixture
def style():
    return XLSXStyle()


class TestXLSXField:
    def test_init(self, style: XLSXStyle):
        f = XLSXField("foo", "bar", CharField(), style, None, style)
        assert f.key == "foo"
        assert f.original_value == "bar"
        assert f.value == "bar"

    def test_cell(self, style: XLSXStyle, worksheet: Worksheet):
        f = XLSXField("foo", "bar", CharField(), style, None, style)
        cell = f.cell(worksheet, 1, 1)
        assert isinstance(cell, Cell)
        assert cell.value == "bar"
