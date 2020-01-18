from rest_framework.response import Response


class XLSXFileMixin(object):
    """
    Mixin which allows the override of the filename being
    passed back to the user when the spreadsheet is downloaded.
    """

    filename = "export.xlsx"

    def get_filename(self):
        """
        Returns a custom filename for the spreadsheet.
        """
        return self.filename

    def finalize_response(self, request, response, *args, **kwargs):
        """
        Return the response with the proper content disposition and the customized
        filename instead of the browser default (or lack thereof).
        """
        response = super(XLSXFileMixin, self).finalize_response(
            request, response, *args, **kwargs
        )
        if (
            isinstance(response, Response)
            and response.accepted_renderer.format == "xlsx"
        ):
            response["content-disposition"] = "attachment; filename={}".format(
                self.get_filename(),
            )
        return response
