import json

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.writer.excel import save_virtual_workbook
from rest_framework.renderers import BaseRenderer


class XLSXRenderer(BaseRenderer):
    """
    Renderer for Excel spreadsheet open data format (xlsx).
    """
    media_type = 'application/xlsx'
    format = 'xlsx'

    def render(self, data, accepted_media_type=None, renderer_context=None):
        """
        Render `data` into JSON, returning a bytestring.
        """
        if not self._check_validatation_data(data):
            return self._json_format_response(data)

        row_count = 0

        if data is None:
            return bytes()

        wb = Workbook()

        ws = wb.active
        ws.title = 'Data Browser Export'

        for row in data['results']:
            # Reset the column count
            column_count = 0

            # Increase the row count
            row_count += 1

            # Set the header row
            if row_count == 1:
                for column_name in row.keys():
                    column_count += 1
                    ws.cell(
                        row=row_count,
                        column=column_count,
                        value='{}'.format(
                            column_name,
                        ),
                    )

                column_count = 0
                row_count += 1

            for column in row.items():
                column_count += 1
                ws.cell(
                    row=row_count,
                    column=column_count,
                    value='{}'.format(
                        column[1],
                    ),
                )

        for ws_column in range(1, column_count + 1):
            col_letter = get_column_letter(ws_column)
            ws.column_dimensions[col_letter].width = 15

        return save_virtual_workbook(wb)

    def _check_validatation_data(self, data):
        detail_key = 'detail'
        if detail_key in data:
            return False
        return True

    def _json_format_response(self, response_data):
        return json.dumps(response_data)