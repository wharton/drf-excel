from openpyxl.styles import Font, PatternFill, Alignment, Border, Color, Side

from drf_excel.utilities import get_setting, XLSXStyle


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


def test_get_setting_not_found():
    assert get_setting("INTEGER_FORMAT") is None
