import io

import pytest
from openpyxl.reader.excel import load_workbook
from rest_framework.test import APIClient

from tests.testapp.models import ExampleModel

@pytest.fixture
def api_client():
    return APIClient()


@pytest.mark.django_db
def test_simple_viewset_model(api_client):
    ExampleModel.objects.create(title="test 1", description="This is a test")
    ExampleModel.objects.create(title="test 2", description="Another test")
    ExampleModel.objects.create(title="test 3", description="Testing this out")

    response = api_client.get("/examples/")

    assert response.status_code == 200
    assert response.headers["Content-Type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8"
    assert response.headers["content-disposition"] == "attachment; filename=my_export.xlsx"

    workbook_buffer = io.BytesIO(response.content)
    workbook = load_workbook(workbook_buffer, read_only=True)

    assert len(workbook.worksheets) == 1
    sheet = workbook.worksheets[0]
    rows = list(sheet.rows)
    assert len(rows) == 4
    r0, r1, r2, r3 = rows

    assert len(r0) == 2
    assert r0[0].value == "title"
    assert r0[1].value == "description"

    assert len(r1) == 2
    assert r1[0].value == "test 1"
    assert r1[1].value == "This is a test"

    assert len(r2) == 2
    assert r2[0].value == "test 2"
    assert r2[1].value == "Another test"

    assert len(r3) == 2
    assert r3[0].value == "test 3"
    assert r3[1].value == "Testing this out"
