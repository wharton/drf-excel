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
    from drf_renderer_xlsx.renderers import XLSXRenderer

    from .models import MyExampleModel
    from .serializers import MyExampleSerializer

    class MyExampleViewSet(XLSXFileMixin, ReadOnlyModelViewSet):
        queryset = MyExampleModel.objects.all()
        serializer_class = MyExampleSerializer
        renderer_classes = (XLSXRenderer,)

# Configure styles 
You can add styles to your worksheet header, column header row and body rows from view attributes `header`, `column_header`, `body`.
You can use any arguments from [openpyxl](https://openpyxl.readthedocs.io/en/stable/styles.html) library to font, alignment, fill and border_side (border always at 4 side of cell).   



    class MyExampleViewSet(XLSXFileMixin, ReadOnlyModelViewSet):
        queryset = MyExampleModel.objects.all()
        serializer_class = MyExampleSerializer
        renderer_classes = (XLSXRenderer,)
    
        column_header = {'titles': ["Column_1_name", "Column_2_name", "Column_3_name"],
                         'width': [17, 30, 17],
                         'height': 25,
                         'style': {'fill': {'fill_type': 'solid',
                                            'start_color': 'FFCCFFCC'},
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
                                   },
                         }
        body = {'style': {'fill': {'fill_type': 'solid',
                                   'start_color': 'FFCCFFCC'},
                          'alignment': {'horizontal': 'center',
                                        'vertical': 'center',
                                        'wrapText': True,
                                        'shrink_to_fit': True},
                          'border_side': {'border_style': 'thin',
                                          'color': 'FF000000'},
                          'font': {'name': 'Arial',
                                   'size': 14,
                                   'bold': False,
                                   'color': 'FF000000'}},
                'height': 40}

Also you can dynamically generate style attributes in methods `get_body`, `get_header`, `get_column_header`.

        def get_header(self):
            starttime, endtime = parse_times(request=self.request)
            datetime_format = "%H:%M:%S %d.%m.%Y"
            return {'tab_title': 'MyReport',
                    'header_title': 'Report from {} to {}'.format(starttime.strftime(datetime_format),
                                                                  endtime.strftime(datetime_format)),
                    'height': 45,
                    'img': 'app/images/MyLogo.png',
                    'style': {'fill': {'fill_type': 'solid',
                                       'start_color': 'FFFFFFFF'},
                              'alignment':
                                  {'horizontal': 'center',
                                   'vertical': 'center',
                                   'wrapText': True,
                                   'shrink_to_fit': True},
                              'border_side': {'border_style': 'thin',
                                              'color': 'FF000000'},
                              'font': {'name': 'Arial',
                                       'size': 16,
                                       'bold': True,
                                       'color': 'FF000000'}
                              }
                   }
Also you can add `color` field to your serializer and fill body rows.

    class exampleSerializer(serializers.Serializer):
        color = serializers.SerializerMethodField()
    
        def get_color(self, instance):
            color_map = {'w': 'FFFFFFCC', 'a': 'FFFFCCCC'}
            return color_map.get(instance.alarm_level, 'FFFFFFFF')

# Contributors

* [Timothy Allen](https://github.com/FlipperPA)
