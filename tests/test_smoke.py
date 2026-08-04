"""不依赖 Office 的冒烟测试：包导入、版本、常量。"""

import offipy


def test_version():
    assert offipy.__version__ == "0.1.0"


def test_progids():
    assert set(offipy.PROGIDS) == {"word", "excel", "ppt"}
