from rest_framework import serializers

from .models import ExampleModel, AllFieldsModel


class ExampleSerializer(serializers.ModelSerializer):
    class Meta:
        model = ExampleModel
        fields = ("title", "description")


class AllFieldsSerializer(serializers.ModelSerializer):
    tags = serializers.ListField(source="get_tag_names")

    class Meta:
        model = AllFieldsModel
        fields = (
            "title",
            "created_at",
            "updated_date",
            "updated_time",
            "age",
            "is_active",
            "tags",
        )
