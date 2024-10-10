from PIL import Image
from rest_framework import serializers
from rest_framework.generics import GenericAPIView
from rest_framework.response import Response

from drf_excel.renderers import XLSXRenderer


class MySerializer(serializers.Serializer):
    title = serializers.CharField()


class MyBaseView(GenericAPIView):
    serializer_class = MySerializer

    def retrieve(self, request, *args, **kwargs):
        return Response({"title": "example"})


class TestXLSXRenderer:
    renderer = XLSXRenderer()

    def test_validation_error(self):
        assert self.renderer.render({"detail": "invalid"}) == '{"detail": "invalid"}'

    def test_none(self):
        assert self.renderer.render(None) == b""

    def test_with_header_attribute(self, tmp_path, workbook_reader):
        image_path = tmp_path / "image.png"
        with Image.new(mode="RGB", size=(100, 100), color="blue") as img:
            img.save(image_path, format="png")

        class MyView(MyBaseView):
            header = {
                "use_header": True,
                "header_title": "My Header",
                "tab_title": "My Tab",
                "img": str(image_path),
                "style": {"font": {"name": "Arial"}},
            }

        result = self.renderer.render({}, renderer_context={"view": MyView})
        wb = workbook_reader(result)
        sheet = wb.worksheets[0]
        rows = list(sheet.rows)
        assert len(rows) == 1
        row0_col0 = rows[0][0]
        assert row0_col0.value == "My Header"
        assert row0_col0.font.name == "Arial"
