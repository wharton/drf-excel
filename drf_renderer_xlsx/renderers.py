import json

from collections.abc import MutableMapping, Iterable
from django.utils.dateparse import parse_datetime
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, NamedStyle
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.writer.excel import save_virtual_workbook
from rest_framework.fields import Field
from rest_framework.renderers import BaseRenderer
from rest_framework.serializers import Serializer
from rest_framework.utils.serializer_helpers import ReturnDict, ReturnList

ESCAPE_CHARS = ('=', '-', '+', '@', '\t', '\r', '\n',)

def get_style_from_dict(style_dict, style_name):
    """
    Make NamedStyle instance from dictionary
    :param style_dict: dictionary with style properties.
           Example:    {'fill': {'fill_type'='solid',
                                 'start_color'='FFCCFFCC'},
                        'alignment': {'horizontal': 'center',
                                      'vertical': 'center',
                                      'wrapText': True,
                                      'shrink_to_fit': True},
                        'border_side': {'border_style': 'thin',
                                        'color': 'FF000000'},
                        'font': {'name': 'Arial',
                                 'size': 14,
                                 'bold': True,
                                 'color': 'FF000000'}
                        }
    :param style_name: name of created style
    :return: openpyxl.styles.NamedStyle instance
    """
    style = NamedStyle(name=style_name)
    if not style_dict:
        return style
    for key, value in style_dict.items():
        if key == "font":
            style.font = Font(**value)
        elif key == "fill":
            style.fill = PatternFill(**value)
        elif key == "alignment":
            style.alignment = Alignment(**value)
        elif key == "border_side":
            side = Side(**value)
            style.border = Border(left=side, right=side, top=side, bottom=side)

    return style


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
        prop_func = getattr(get_from, "get_{}".format(prop_name), None)
        if prop_func:
            prop = prop_func()
    if prop is None:
        prop = default
    return prop


class XLSXRenderer(BaseRenderer):
    """
    Renderer for Excel spreadsheet open data format (xlsx).
    """

    media_type = "application/xlsx"
    format = "xlsx"
    xlsx_header_dict = {}
    ignore_headers = []
    boolean_labels = None
    date_format_mappings = None
    custom_mappings = None
    sanitize_fields = True  # prepend possibly malicious values with "'"

    def render(self, data, accepted_media_type=None, renderer_context=None):
        """
        Render `data` into XLSX workbook, returning a workbook.
        """
        if not self._check_validatation_data(data):
            return json.dumps(data)

        if data is None:
            return bytes()

        wb = Workbook()
        self.ws = wb.active

        results = data["results"] if "results" in data else data

        # Take header and column_header params from view
        header = get_attribute(renderer_context["view"], "header", {})
        self.ws.title = header.get("tab_title", "Report")
        header_title = header.get("header_title", "Report")
        img_addr = header.get("img")
        if img_addr:
            img = Image(img_addr)
            self.ws.add_image(img, "A1")
        header_style = get_style_from_dict(header.get("style"), "header_style")

        column_header = get_attribute(renderer_context["view"], "column_header", {})
        column_header_style = get_style_from_dict(
            column_header.get("style"), "column_header_style"
        )
        column_count = 0
        row_count = 1
        if header:
            row_count += 1
        # Make column headers
        column_titles = column_header.get("titles", [])

        # If we have results, get view and serializer from context, then flatten field
        # names
        if len(results):
            drf_view = renderer_context.get("view")

            # Set `xlsx_use_labels = True` inside the API View to enable labels.
            use_labels = getattr(drf_view, "xlsx_use_labels", False)

            # A list of header keys to ignore in our export
            self.ignore_headers = getattr(drf_view, "xlsx_ignore_headers", [])

            # Create a mapping dict named `xlsx_boolean_labels` inside the API View.
            # I.e.: xlsx_boolean_labels: {True: "Yes", False: "No"}
            self.boolean_display = getattr(drf_view, "xlsx_boolean_labels", None)

            # set dict named xlsx_date_format_mappings with headers as keys and
            # formatting as value. i.e. { 'created_at': '%d.%m.%Y, %H:%M' }
            self.date_format_mappings = getattr(
                drf_view, "xlsx_date_format_mappings", None
            )

            # Map a specific key to a column (I.e. if the field returns a json) or pass
            # a function to format the value
            # Example with key:
            # {"custom_choice": "display"}, showing 'display' in the
            # 'custom_choice' col
            # Example with function:
            # {"custom_choice": custom_func }, passing the value of 'custom_choice' to
            # 'custom_func', allowing for formatting logic
            self.custom_mappings = getattr(drf_view, "xlsx_custom_mappings", None)

            self.xlsx_header_dict = self._flatten_serializer_keys(
                drf_view.get_serializer(), use_labels=use_labels
            )

            for column_name, column_label in self.xlsx_header_dict.items():
                if column_name == "row_color":
                    continue
                column_count += 1
                if column_count > len(column_titles):
                    column_name_display = column_label
                else:
                    column_name_display = column_titles[column_count - 1]

                self.ws.cell(
                    row=row_count, column=column_count, value=column_name_display
                ).style = column_header_style
            self.ws.row_dimensions[row_count].height = column_header.get("height", 45)

        # Set the header row
        if header:
            last_col_letter = "G"
            if column_count:
                last_col_letter = get_column_letter(column_count)
            self.ws.merge_cells("A1:{}1".format(last_col_letter))

            cell = self.ws.cell(row=1, column=1, value=header_title)
            cell.style = header_style
            self.ws.row_dimensions[1].height = header.get("height", 45)

        # Set column width
        column_width = column_header.get("column_width", 20)
        if isinstance(column_width, list):
            for i, width in enumerate(column_width):
                col_letter = get_column_letter(i + 1)
                self.ws.column_dimensions[col_letter].width = width
        else:
            for self.ws_column in range(1, column_count + 1):
                col_letter = get_column_letter(self.ws_column)
                self.ws.column_dimensions[col_letter].width = column_width

        # Make body
        self.body = get_attribute(renderer_context["view"], "body", {})
        self.body_style = get_style_from_dict(self.body.get("style"), "body_style")
        if isinstance(results, dict):
            self._make_body(results, row_count)
        elif isinstance(results, list):
            for row in results:
                self._make_body(row, row_count)
                row_count += 1

        return save_virtual_workbook(wb)

    def _check_validatation_data(self, data):
        detail_key = "detail"
        if detail_key in data:
            return False
        return True

    def _flatten_serializer_keys(
        self,
        serializer,
        parent_key="",
        parent_label="",
        key_sep=".",
        list_sep=", ",
        label_sep=" > ",
        use_labels=False,
    ):
        """
        Iterate through serializer fields recursively when field is a nested serializer.
        """
        def _get_label(parent_label, label_sep, obj):
            if getattr(v, "label", None):
                if parent_label:
                    return f"{parent_label}{label_sep}{v.label}"
                else:
                    return str(v.label)
            else:
                return False

        _header_dict = {}
        _fields = serializer.get_fields()
        for k, v in _fields.items():
            new_key = f"{parent_key}{key_sep}{k}" if parent_key else k
            # Skip headers we want to ignore
            if new_key in self.ignore_headers:
                continue
            # Iterate through fields if field is a serializer. Check for labels and
            # append if `use_labels` is True. Fallback to using keys.
            if isinstance(v, Serializer):
                if use_labels and getattr(v, "label", None):
                    _header_dict.update(
                        self._flatten_serializer_keys(
                            v,
                            new_key,
                            _get_label(parent_label, label_sep, v),
                            key_sep,
                            list_sep,
                            label_sep,
                            use_labels,
                        )
                    )
                else:
                    _header_dict.update(
                        self._flatten_serializer_keys(
                            v,
                            new_key,
                            key_sep=key_sep,
                            list_sep=list_sep,
                            label_sep=label_sep,
                            use_labels=use_labels,
                        )
                    )
            elif isinstance(v, Field):
                if use_labels and getattr(v, "label", None):
                    _header_dict[new_key] = _get_label(parent_label, label_sep, v)
                else:
                    _header_dict[new_key] = new_key
        return _header_dict

    def _flatten_data(self, data, parent_key="", key_sep=".", list_sep=", "):
        def _append_item(key, value):
            if self.date_format_mappings and key in self.date_format_mappings:
                try:
                    date = parse_datetime(value)
                    items.append((key, date.strftime(self.date_format_mappings[key])))
                    return
                except TypeError:
                    pass
            items.append((key, value))

        items = []
        for k, v in data.items():
            new_key = f"{parent_key}{key_sep}{k}" if parent_key else k
            if self.custom_mappings and new_key in self.custom_mappings:
                custom_mapping = self.custom_mappings[new_key]
                if type(custom_mapping) is str:
                    _append_item(new_key, v.get(custom_mapping))
                elif callable(custom_mapping):
                    _append_item(new_key, custom_mapping(v))
            elif isinstance(v, MutableMapping):
                items.extend(self._flatten_data(v, new_key, key_sep=key_sep).items())
            elif isinstance(v, Iterable) and not isinstance(v, str):
                if len(v) > 0 and isinstance(v[0], Iterable):
                    # array of array; write as json
                    _append_item(new_key, json.dumps(v))
                else:
                    # Flatten the array into a comma separated string to fit
                    # in a single spreadsheet column
                    _append_item(new_key, list_sep.join(v))
            elif self.boolean_display and type(v) is bool:
                _append_item(new_key, str(self.boolean_display.get(v, v)))
            else:
                _append_item(new_key, v)
        return dict(items)

    def _sanitize_value(self, raw_value):
        # prepend ' if raw_value is starting with possible malicious char
        if self.sanitize_fields and raw_value:
            str_value = str(raw_value)
            str_value = ILLEGAL_CHARACTERS_RE.sub('', str_value)   # remove ILLEGAL_CHARACTERS so it doesn't crash
            return "'" + str_value if str_value.startswith(ESCAPE_CHARS) else str_value
        return raw_value

    def _make_body(self, row, row_count):
        column_count = 0
        row_count += 1
        flattened_row = self._flatten_data(row)
        for header_key in self.xlsx_header_dict:
            if header_key == "row_color":
                continue
            column_count += 1
            sanitized_value = self._sanitize_value(flattened_row.get(header_key))
            cell = self.ws.cell(
                row=row_count, column=column_count, value=sanitized_value
            )
            cell.style = self.body_style
        self.ws.row_dimensions[row_count].height = self.body.get("height", 40)
        if "row_color" in row:
            last_letter = get_column_letter(column_count)
            cell_range = self.ws[
                "A{}".format(row_count) : "{}{}".format(last_letter, row_count)
            ]
            fill = PatternFill(fill_type="solid", start_color=row["row_color"])
            for r in cell_range:
                for c in r:
                    c.fill = fill
