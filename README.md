# Django REST Framework Renderer: XLSX

`drf-renderer-xlsx` provides an XLSX renderer for Django REST Framework. It uses OpenPyXL to create the spreadsheet and returns the data.

# Requirements

It may work with earlier versions, but has been tested with the following:

* Python >= 3.6
* Django >= 2.2
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

Styles can be added to your worksheet header, column header row, and body rows, from view attributes `header`, `column_header`, `body`. Any arguments from [the OpenPyXL package](https://openpyxl.readthedocs.io/en/stable/styles.html) can be used for font, alignment, fill and border_side (border will always be all side of cell).   

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
        'use_header': True,  # show the header_title 
        'header_title': 'Report from {} to {}'.format(
            starttime.strftime(datetime_format),
            endtime.strftime(datetime_format),
        ),
        'tab_title': 'Report',  # title of tab/workbook
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

# Controlling XLSX headers and values

## Use Serializer Field labels as header names

By default, headers will use the same 'names' as they are returned by the API. This can be changed by setting `xlsx_use_labels = True` inside your API View. 

Instead of using the field names, the export will use the labels as they are defined inside your Serializer. A serializer field defined as `title = serializers.CharField(label=_("Some title"))` would return `Some title` instead of `title`, also supporting translations. If no label is set, it will fall back to using `title`.


## Ignore fields

By default, all fields are exported, but you might want to exclude some fields from your export. To do so, you can set an array with fields you want to exclude: `xlsx_ignore_headers = [<excluded fields>]`.

This also works with nested fields, separated with a dot (i.e. `icon.url`).


## Name boolean values

`True` and `False` as values for boolean fields are not always the best representation and don't support translation. This can be controlled with `xlsx_boolean_labels`. 

`xlsx_boolean_labels = {True: _('Yes'), False: _('No')}` will replace `True` with `Yes` and `False` with `No`.


## Format dates

To format dates differently than what DRF returns (eg. 2013-01-29T12:34:56.000000Z) `xlsx_date_format_mappings` takes a ´dict` with the field name as its key and the date(time) format as its value:

```    
xlsx_date_format_mappings = {
    'created_at': '%d.%m.%Y %H:%M',
    'updated_at': '%d.%m.%Y %H:%M'
}
```


## Custom columns

You might find yourself explicitly returning a dict in your API response and would like to use its data to display additional columns. This can be done by passing `xlsx_custom_cols`.
```
xlsx_custom_cols = {
    'my_custom_col.val1.title': {
        'label': 'Custom column!',
        'formatter': custom_value_formatter
    }
}

# Example function:
def custom_value_formatter(val):
    return val + '!!!'

# Example response:
{ 
    results: [
        {
            title: 'XLSX renderer',
            url: 'https://github.com/wharton/drf-renderer-xlsx'
            returned_dict: {
                val1: {
                    title: 'Sometimes'
                },
                val2: {
                    title: 'There is no way around'
                }
            }
        }
    ]
}
```

When no `label` is passed, `drf-renderer-xlsx` will display the key name in the header.
`formatter` is also optional and accepts a function, which will then receive the value it is mapped to (it would receive "Sometimes" and return "Sometimes!!!" in our example).


## Custom mappings

Assuming you have a field that returns a `dict` instead of a simple `str`, you might not want to return the whole object but only a value of it. Let's say `status` returns `{ value: 1, display: 'Active' }`. To return the `display` value in the `status` column, we can do this:
```
xlsx_custom_mappings = {
    'status': 'display'
}
```
A probably more common case is that you want to change how a value is formatted. `xlsx_custom_mappings` also takes functions as values. Assuming we have a field `description`, and for some strange reason want to reverse the text, we can do this:

```
def reverse_text(val):
    return val[::-1]

xlsx_custom_mappings = {
    'description': reverse_text
}
```


# Release Notes

Release notes are [available on GitHub](https://github.com/wharton/drf-renderer-xlsx/releases).

## Maintainers

* [Timothy Allen](https://github.com/FlipperPA) at [The Wharton School](https://github.com/wharton)
* [Thomas Willems](https://github.com/willtho89)

This package was created by the staff of [Wharton Research Data Services](https://wrds.wharton.upenn.edu/). We are thrilled that [The Wharton School](https://www.wharton.upenn.edu/) allows us a certain amount of time to contribute to open-source projects. We add features as they are necessary for our projects, and try to keep up with Issues and Pull Requests as best we can. Due to constraints of time (our full time jobs!), Feature Requests without a Pull Request may not be implemented, but we are always open to new ideas and grateful for contributions and our package users.

## Contributors (Thank You!)

* [Armaan Tobaccowalla](https://github.com/ArmaanT)
* [Davis Haupt](https://github.com/davish)
* [Eric Wang](https://github.com/ezwang)
* [Felipe Schmitt](https://github.com/fsschmitt)
* [ffruit](https://github.com/frruit)
* [Gonzalo Ayuso](https://github.com/gonzalo123)
* [LeeHanYeong](https://github.com/LeeHanYeong)
* [Mathieu Rampant](https://github.com/rptmat57)
* [Nick Kozhenin](https://github.com/mast22)
* [paveloder](https://github.com/paveloder)
* [Pavel Bryantsev](https://github.com/Tigven)
* [Pavel Tolstolytko](https://github.com/eshikvtumane)
* [Tim](https://github.com/Shin--/)
* [Vincenz E.](https://github.com/vincenz-e)
* [YunpengZhan](https://github.com/runningzyp)
