"""不依赖 Office 的冒烟测试：包导入、版本、常量。"""

import office_kit


def test_version():
    assert office_kit.__version__ == "0.1.0"


def test_progids():
    assert set(office_kit.PROGIDS) == {"word", "excel", "ppt"}
