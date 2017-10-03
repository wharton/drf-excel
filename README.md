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

# Contributors

* [Timothy Allen](https://github.com/FlipperPA)
