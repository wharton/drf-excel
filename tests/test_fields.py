from decimal import Decimal

import pytest
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet
from rest_framework.fields import CharField, IntegerField, FloatField, DecimalField

from drf_excel.fields import XLSXField, XLSXNumberField
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


class TestXLSXNumberField:
    @pytest.mark.parametrize(
        ("original_value", "cleaned_value"),
        [
            # Conversion worked
            (42, 42),
            ("42", 42),
            (42.1, 42),
            (True, 1),
            (False, 0),
            # Conversion fails
            ("42.0", "42.0"),
            ("foo", "foo"),
            (None, None),
        ],
    )
    def test_init_integer_field(self, style: XLSXStyle, original_value, cleaned_value):
        f = XLSXNumberField(
            key="age",
            value=original_value,
            field=IntegerField(),
            style=style,
            mapping=None,
            cell_style=style,
        )
        assert f.key == "age"
        assert f.original_value == original_value
        assert f.value == cleaned_value

    @pytest.mark.parametrize(
        ("original_value", "cleaned_value"),
        [
            # Conversion worked
            (49, 49),
            ("36", 36),
            (11.4, 11.4),
            ("42.0", 42.0),
            (True, 1),
            (False, 0),
            # Conversion fails
            ("foo", "foo"),
            (None, None),
        ],
    )
    def test_init_float_field(self, style: XLSXStyle, original_value, cleaned_value):
        f = XLSXNumberField(
            key="temperature",
            value=original_value,
            field=FloatField(),
            style=style,
            mapping=None,
            cell_style=style,
        )
        assert f.key == "temperature"
        assert f.original_value == original_value
        assert f.value == cleaned_value

    @pytest.mark.parametrize(
        ("original_value", "cleaned_value"),
        [
            # Conversion worked
            (49, Decimal(49)),
            ("36", Decimal(36)),
            (11.4, Decimal(11.4)),
            ("42.20", Decimal("42.20")),
            (True, Decimal(1)),
            (False, Decimal(0)),
            (Decimal(5), Decimal(5)),
            (Decimal("35.00"), Decimal("35.00")),
            (Decimal("55.9"), Decimal("55.9")),
            (Decimal("123.9999"), Decimal("123.9999")),
            # Conversion fails
            ("foo", "foo"),
            (None, None),
        ],
    )
    def test_init_decimal_field(self, style: XLSXStyle, original_value, cleaned_value):
        f = XLSXNumberField(
            key="price",
            value=original_value,
            field=DecimalField(max_digits=10, decimal_places=2),
            style=style,
            mapping=None,
            cell_style=style,
        )
        assert f.key == "price"
        assert f.original_value == original_value
        assert f.value == cleaned_value

    def test_cell_integer(self, style: XLSXStyle, worksheet: Worksheet):
        f = XLSXNumberField(
            key="age",
            value=42,
            field=IntegerField(),
            style=style,
            mapping=None,
            cell_style=style,
        )
        cell = f.cell(worksheet, 1, 1)
        assert isinstance(cell, Cell)
        assert cell.value == 42
        assert cell.number_format == "0"

    def test_cell_float(self, style: XLSXStyle, worksheet: Worksheet):
        f = XLSXNumberField(
            key="weight",
            value=35.5,
            field=FloatField(),
            style=style,
            mapping=None,
            cell_style=style,
        )
        cell = f.cell(worksheet, 1, 1)
        assert isinstance(cell, Cell)
        assert cell.value == 35.5
        assert cell.number_format == "0.00"
