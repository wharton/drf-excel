import pytest
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


@pytest.fixture
def workbook() -> Workbook:
    return Workbook()


@pytest.fixture
def worksheet(workbook: Workbook) -> Worksheet:
    return Worksheet(workbook)
