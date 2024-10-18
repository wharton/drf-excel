from rest_framework.viewsets import ReadOnlyModelViewSet

from drf_excel.mixins import XLSXFileMixin
from drf_excel.renderers import XLSXRenderer

from .models import AllFieldsModel, ExampleModel, SecretFieldModel
from .serializers import AllFieldsSerializer, ExampleSerializer, SecretFieldSerializer


class ExampleViewSet(XLSXFileMixin, ReadOnlyModelViewSet):
    queryset = ExampleModel.objects.all()
    serializer_class = ExampleSerializer
    renderer_classes = (XLSXRenderer,)
    filename = "my_export.xlsx"


class AllFieldsViewSet(XLSXFileMixin, ReadOnlyModelViewSet):
    queryset = AllFieldsModel.objects.all()
    serializer_class = AllFieldsSerializer
    renderer_classes = (XLSXRenderer,)
    filename = "al_fileds.xlsx"


class SecretFieldViewSet(XLSXFileMixin, ReadOnlyModelViewSet):
    queryset = SecretFieldModel.objects.all()
    serializer_class = SecretFieldSerializer
    renderer_classes = (XLSXRenderer,)
    filename = "secret.xlsx"
