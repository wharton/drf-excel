from django.db import models


class ExampleModel(models.Model):
    title = models.CharField(max_length=100)
    description = models.TextField()

    def __str__(self):
        return self.title


class Tag(models.Model):
    name = models.CharField(max_length=100)

    def __str__(self):
        return self.name


class AllFieldsModel(models.Model):
    title = models.CharField(max_length=100)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_date = models.DateField(auto_now=True)
    updated_time = models.TimeField(auto_now=True)
    age = models.IntegerField()
    is_active = models.BooleanField(default=True)
    tags = models.ManyToManyField(Tag, related_name="all_fields")

    def __str__(self):
        return self.title

    def get_tag_names(self):
        return [tag.name for tag in self.tags.all()]


class SecretFieldModel(models.Model):
    title = models.CharField(max_length=100)
    secret = models.CharField(max_length=100)

    def __str__(self):
        return self.title
