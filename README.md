# Django REST Framework Renderer: XLSX

`drf-renderer-xlsx` provides an XLSX renderer for Django REST Framework. It uses OpenPyXL to create the spreadsheet and returns the data.

# Requirements

It may work with earlier versions, but has been tested with the following:

* Django >= 1.11
* Django REST Framework >= 3.6
* OpenPyXL >= 2.4

# Installation

    pip install drf-renderer-xls

Then add the following to your `REST_FRAMEWORK` settings:

    REST_FRAMEWORK = {
        ...

        'DEFAULT_RENDERER_CLASSES': (
            'rest_framework.renderers.JSONRenderer',
            'rest_framework.renderers.BrowsableAPIRenderer',
            'drf_renderer_xlsx.renderers.XLSXRenderer',
        ),
    }

To avoid having a file streamed without a filename (which the browser will often default to the filename "download", with no extension), we need to use a mixin to override the Content-Disposition like so:

    from rest_framework.viewsets import ReadOnlyModelViewSet
    from drf_renderer_xlsx.mixins import XLSXFileMixin

    from .models import MyExampleModel
    from .serializers import MyExampleSerializer

    class MyExampleViewSet(XLSXFileMixin, ReadOnlyModelViewSet):
        queryset = MyExampleModel.objects.all()
        serializer_class = MyExampleSerializer

# Contributors

* [Timothy Allen](https://github.com/FlipperPA)
