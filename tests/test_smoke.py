"""不依赖 Office 的冒烟测试：包导入、版本、常量。"""

from packaging.version import Version

import offipy


def test_version():
    # 版本单一来源在 src/offipy/__init__.py；断言合法 PEP 440（含 a/b/rc 预发布）
    # 用 Version() 而非正则：0.9.0a1 这类预发布不炸，稳定版 1.0.0 也覆盖
    Version(offipy.__version__)


def test_progids():
    assert set(offipy.PROGIDS) == {"word", "excel", "ppt"}
