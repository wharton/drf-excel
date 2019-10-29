# Django REST Framework Renderer: XLSX

`drf-renderer-xlsx` provides an XLSX renderer for Django REST Framework. It uses OpenPyXL to create the spreadsheet and returns the data.

# Requirements

It may work with earlier versions, but has been tested with the following:

* Python >= 3.6
* Django >= 1.11
* Django REST Framework >= 3.6
* OpenPyXL >= 2.4

# Installation

```bash
pip install drf-renderer-xlsx
```

Then add the following to your `REST_FRAMEWORK` settings:

```python
    REST_FRAMEWORK = {
        ...

        'DEFAULT_RENDERER_CLASSES': (
            'rest_framework.renderers.JSONRenderer',
            'rest_framework.renderers.BrowsableAPIRenderer',
            'drf_renderer_xlsx.renderers.XLSXRenderer',
        ),
    }
```

To avoid having a file streamed without a filename (which the browser will often default to the filename "download", with no extension), we need to use a mixin to override the `Content-Disposition` header. If no `filename` is provided, it will default to `export.xlsx`. For example:

```python
from rest_framework.viewsets import ReadOnlyModelViewSet
from drf_renderer_xlsx.mixins import XLSXFileMixin
from drf_renderer_xlsx.renderers import XLSXRenderer

from .models import MyExampleModel
from .serializers import MyExampleSerializer

class MyExampleViewSet(XLSXFileMixin, ReadOnlyModelViewSet):
    queryset = MyExampleModel.objects.all()
    serializer_class = MyExampleSerializer
    renderer_classes = (XLSXRenderer,)
    filename = 'my_export.xlsx'
```

The `XLSXFileMixin` also provides a `get_filename()` method which can be overridden, if you prefer to provide a filename programmatically instead of the `filename` attribute.

# Configuring Styles 

Styles can be added to your worksheet header, column header row, and body rows, from view attributes `header`, `column_header`, `body`. Any arguments from [the openpyxl library](https://openpyxl.readthedocs.io/en/stable/styles.html) can be used for font, alignment, fill and border_side (border will always be all side of cell).   

```python
class MyExampleViewSet(XLSXFileMixin, ReadOnlyModelViewSet):
    queryset = MyExampleModel.objects.all()
    serializer_class = MyExampleSerializer
    renderer_classes = (XLSXRenderer,)

    column_header = {
        'titles': [
            "Column_1_name",
            "Column_2_name",
            "Column_3_name",
        ],
        'column_width': [17, 30, 17],
        'height': 25,
        'style': {
            'fill': {
                'fill_type': 'solid',
                'start_color': 'FFCCFFCC',
            },
            'alignment': {
                'horizontal': 'center',
                'vertical': 'center',
                'wrapText': True,
                'shrink_to_fit': True,
            },
            'border_side': {
                'border_style': 'thin',
                'color': 'FF000000',
            },
            'font': {
                'name': 'Arial',
                'size': 14,
                'bold': True,
                'color': 'FF000000',
            },
        },
    }
    body = {
        'style': {
            'fill': {
                'fill_type': 'solid',
                'start_color': 'FFCCFFCC',
            },
            'alignment': {
                'horizontal': 'center',
                'vertical': 'center',
                'wrapText': True,
                'shrink_to_fit': True,
            },
            'border_side': {
                'border_style': 'thin',
                'color': 'FF000000',
            },
            'font': {
                'name': 'Arial',
                'size': 14,
                'bold': False,
                'color': 'FF000000',
            }
        },
        'height': 40,
    }
```

Also you can dynamically generate style attributes in methods `get_body`, `get_header`, `get_column_header`.

```python
def get_header(self):
    starttime, endtime = parse_times(request=self.request)
    datetime_format = "%H:%M:%S %d.%m.%Y"
    return {
        'tab_title': 'MyReport',
        'header_title': 'Report from {} to {}'.format(
            starttime.strftime(datetime_format),
            endtime.strftime(datetime_format),
        ),
        'height': 45,
        'img': 'app/images/MyLogo.png',
        'style': {
            'fill': {
                'fill_type': 'solid',
                'start_color': 'FFFFFFFF',
            },
            'alignment': {
                'horizontal': 'center',
                'vertical': 'center',
                'wrapText': True,
                'shrink_to_fit': True,
            },
            'border_side': {
                'border_style': 'thin',
                'color': 'FF000000',
            },
            'font': {
                'name': 'Arial',
                'size': 16,
                'bold': True,
                'color': 'FF000000',
            }
        }
    }
```

Also you can add `color` field to your serializer and fill body rows.

```python
class ExampleSerializer(serializers.Serializer):
    color = serializers.SerializerMethodField()

    def get_color(self, instance):
        color_map = {'w': 'FFFFFFCC', 'a': 'FFFFCCCC'}
        return color_map.get(instance.alarm_level, 'FFFFFFFF')
```

## Release Notes

### 0.3.4

* Switch to `setuptools_scm`. Add support for `ReturnDict` in addition to `ReturnList`.

### 0.3.3

* Add support for nested arrays, flattening them into a string: `value1, value2, value3`, etc.

### 0.3.2

* Add supported for nested values; flattens sub-values into `sub.value1, sub.value2, sub.value3`, etc.

### 0.3.1

* Fix an error when an empty result set was returned from the endpoint. Now, it will properly just download an empty spreadsheet.
* Remove an errant `format()` function which was removing typing from the spreadsheet.

### 0.3.0

* Add support for custom spreadsheet styles (thanks, Pavel Bryantsev!)
* Add an attribute for setting the download filename instead of `export.xlsx` per view.

## Maintainer

* [Timothy Allen](https://github.com/FlipperPA) at [The Wharton School](https://github.com/wharton)

This package is maintained by the staff of [Wharton Research Data Services](https://wrds.wharton.upenn.edu/). We are thrilled that [The Wharton School](https://www.wharton.upenn.edu/) allows us a certain amount of time to contribute to open-source projects. We add features as they are necessary for our projects, and try to keep up with Issues and Pull Requests as best we can. Due to constraints of time (our full time jobs!), Feature Requests without a Pull Request may not be implemented, but we are always open to new ideas and grateful for contributions and our package users.

## Contributors (Thank You!)

* [Pavel Bryantsev](https://github.com/Tigven)
* [Felipe Schmitt](https://github.com/fsschmitt)
* [ffruit](https://github.com/frruit)
* [Pavel Tolstolytko](https://github.com/eshikvtumane)
* [Thomas Willems](https://github.com/willtho89)
