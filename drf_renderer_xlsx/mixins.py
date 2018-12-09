class XLSXFileMixin(object):
    """
    Mixin which allows the override of the filename being
    passed back to the user when the spreadsheet is downloaded.
    """

    filename = "export.xlsx"

    def finalize_response(self, request, response, *args, **kwargs):
        response = super(XLSXFileMixin, self).finalize_response(
            request, response, *args, **kwargs
        )
        if response.accepted_renderer.format == "xlsx":
            response["content-disposition"] = "attachment; filename={}".format(
                self.filename
            )
        return response
