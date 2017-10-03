from openpyxl import Workbook
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
            # Reset the column count
            column_count = 0

            # Increase the row count
            row_count += 1

            # Set the XLSX header row
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

        return save_virtual_workbook(wb)
