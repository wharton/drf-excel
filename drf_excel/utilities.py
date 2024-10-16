from django.conf import settings as django_settings
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ESCAPE_CHARS = ("=", "-", "+", "@", "\t", "\r", "\n")


class XLSXStyle:
    # Class that holds all parts of a style, but without being an actual NamedStyle

    def __init__(self, style_dict=None):
        """
        Make style from dictionary.
        :param style_dict: dictionary with style properties.
            Example: {
                'fill': { 'fill_type': 'solid', 'start_color': 'FFCCFFCC' },
                'alignment': {
                    'horizontal': 'center',
                    'vertical': 'center',
                    'wrapText': True,
                    'shrink_to_fit': True
                },
                'border_side': { 'border_style': 'thin', 'color': 'FF000000' },
                'font': {
                    'name': 'Arial',
                    'size': 14,
                    'bold': True,
                    'color': 'FF000000'
                },
                'format': '0.00E+00',
            }
        :return: XLSXStyle object
        """
        if style_dict is None:
            style_dict = {}
        self.font = Font(**style_dict.get("font")) if "font" in style_dict else None
        self.fill = (
            PatternFill(**style_dict.get("fill")) if "fill" in style_dict else None
        )
        self.alignment = (
            Alignment(**style_dict.get("alignment"))
            if "alignment" in style_dict
            else None
        )
        self.number_format = style_dict.get("format", None)
        side = (
            Side(**style_dict.get("border_side"))
            if "border_side" in style_dict
            else None
        )
        self.border = (
            Border(left=side, right=side, top=side, bottom=side) if side else None
        )


def get_attribute(get_from, prop_name, default=None):
    """
    Get attribute from object with name <prop_name>, or take it from function get_<prop_name>
    :param get_from: instance of object
    :param prop_name: name of attribute (str)
    :param default: what to return if attribute doesn't exist
    :return: value of attribute <prop_name> or default
    """
    prop = getattr(get_from, prop_name, None)
    if not prop:
        prop_func = getattr(get_from, f"get_{prop_name}", None)
        if prop_func:
            prop = prop_func()
    if prop is None:
        prop = default
    return prop


def get_setting(key, default=None):
    return getattr(django_settings, f"DRF_EXCEL_{key}", default)


def sanitize_value(value):
    # prepend ' if value is starting with possible malicious char
    if value:
        str_value = str(value)
        str_value = ILLEGAL_CHARACTERS_RE.sub(
            "", str_value
        )  # remove ILLEGAL_CHARACTERS so it doesn't crash
        return "'" + str_value if str_value.startswith(ESCAPE_CHARS) else str_value
    return value


def set_cell_style(cell: Cell, style: XLSXStyle):
    # We are not applying the whole style directly, otherwise we cannot override any part of it
    if style:
        # Only set properties that are provided
        if style.font:
            cell.font = style.font
        if style.fill:
            cell.fill = style.fill
        if style.alignment:
            cell.alignment = style.alignment
        if style.border:
            cell.border = style.border
        cell.number_format = style.number_format or cell.number_format
