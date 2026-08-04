"""不依赖 Office 的冒烟测试：包导入、版本、常量。"""

import re

import offipy


def test_version():
    # 版本单一来源在 src/offipy/__init__.py；断言 semver 格式，避免每次 bump 改测试
    assert re.fullmatch(r"\d+\.\d+\.\d+", offipy.__version__)


def test_progids():
    assert set(offipy.PROGIDS) == {"word", "excel", "ppt"}
