from openpyxl import Workbook
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
        row_count = 0

        if data is None:
            return bytes()

        wb = Workbook()

        ws = wb.active
        ws.title = 'Data Browser Export'

        for row in data['results']:
            # Set the header row
            if row_count == 0:
                ws.append(list(row.keys()))

            # Increase the row count
            row_count += 1

            # Reset the column count
            column_count = 0

            # Build the column cells
            for column in row.items():
                column_count += 1
                ws.cell(
                    row=row_count,
                    column=column_count,
                    value='{}'.format(
                        column[1],
                    ),
                )

        for ws_column in range(1, column_count):
            col_letter = get_column_letter(ws_column)
            ws.column_dimensions[col_letter].auto_size = True

        return save_virtual_workbook(wb)
