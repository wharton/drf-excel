from drf_excel.utilities import get_setting


def test_get_setting_not_found():
    assert get_setting("INTEGER_FORMAT") is None