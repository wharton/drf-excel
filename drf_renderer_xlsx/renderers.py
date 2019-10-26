import json

from collections.abc import MutableMapping, Iterable
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, NamedStyle
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.writer.excel import save_virtual_workbook
from rest_framework.renderers import BaseRenderer
from rest_framework.utils.serializer_helpers import ReturnDict, ReturnList


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

    def render(self, data, accepted_media_type=None, renderer_context=None):
        """
        Render `data` into XLSX workbook, returning a workbook.
        """
        if not self._check_validatation_data(data):
            return self._json_format_response(data)

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

        # If we have results, pull the columns names from the keys of the first row
        if len(results):
            if isinstance(results, ReturnDict):
                column_names_first_row = results
            elif isinstance(results, ReturnList):
                column_names_first_row = self._flatten(results[0])
            for column_name in column_names_first_row.keys():
                if column_name == "row_color":
                    continue
                column_count += 1
                if column_count > len(column_titles):
                    column_name_display = column_name
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
        if isinstance(results, ReturnDict):
            self._make_body(results, row_count)
        elif isinstance(results, ReturnList):
            for row in results:
                self._make_body(row, row_count)
                row_count += 1

        return save_virtual_workbook(wb)

    def _check_validatation_data(self, data):
        detail_key = "detail"
        if detail_key in data:
            return False
        return True

    def _flatten(self, data, parent_key='', key_sep='.', list_sep=', '):
        items = []
        for k, v in data.items():
            new_key = f"{parent_key}{key_sep}{k}" if parent_key else k
            if isinstance(v, MutableMapping):
                items.extend(self._flatten(v, new_key, key_sep=key_sep).items())
            elif isinstance(v, Iterable) and not isinstance(v, str):
                if isinstance(v[0], Iterable):
                    # array of array; write as json
                    items.append((new_key, json.dumps(v)))
                else:
                    # Flatten the array into a comma separated string to fit
                    # in a single spreadsheet column
                    items.append((new_key, list_sep.join(v)))
            else:
                items.append((new_key, v))
        return dict(items)

    def _json_format_response(self, response_data):
        return json.dumps(response_data)

    def _make_body(self, row, row_count):
        column_count = 0
        row_count += 1
        flatten_row = self._flatten(row)
        for column_name, value in flatten_row.items():
            if column_name == "row_color":
                continue
            column_count += 1
            cell = self.ws.cell(
                row=row_count, column=column_count, value=value,
            )
            cell.style = self.body_style
        self.ws.row_dimensions[row_count].height = self.body.get("height", 40)
        if "row_color" in row:
            last_letter = get_column_letter(column_count)
            cell_range = self.ws[
                         "A{}".format(row_count): "{}{}".format(last_letter, row_count)
                         ]
            fill = PatternFill(fill_type="solid", start_color=row["row_color"])
            for r in cell_range:
                for c in r:
                    c.fill = fill