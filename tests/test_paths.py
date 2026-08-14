"""用户数据路径测试：env 覆盖 + 各平台回退 + 默认落盘路径。"""

import sys
from datetime import datetime as _real_datetime
from datetime import timezone
from pathlib import Path

import pytest

from offipy.paths import converter_data_dir, default_save_path, user_data_dir


def test_user_data_dir_windows_uses_localappdata(monkeypatch):
    monkeypatch.setattr(sys, "platform", "win32")
    # 正斜杠：Path 在 Windows 归一化为反斜杠、POSIX 保持正斜杠，跨平台断言一致
    monkeypatch.setenv("LOCALAPPDATA", "C:/Users/test/AppData/Local")
    assert user_data_dir() == Path("C:/Users/test/AppData/Local") / "offipy"


def test_user_data_dir_windows_fallback_home(monkeypatch):
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    monkeypatch.setattr("offipy.paths.Path.home", lambda: Path("C:/Users/test"))
    assert user_data_dir() == Path("C:/Users/test") / ".offipy"


def test_user_data_dir_linux_xdg(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")
    monkeypatch.setenv("XDG_DATA_HOME", "/home/u/.local/share")
    assert user_data_dir() == Path("/home/u/.local/share/offipy")


def test_user_data_dir_macos(monkeypatch):
    monkeypatch.setattr(sys, "platform", "darwin")
    assert user_data_dir() == Path.home() / "Library" / "Application Support" / "offipy"


def test_converter_data_dir_env_override(monkeypatch, tmp_path):
    monkeypatch.setenv("OFFIPY_CONVERTER_DATA_DIR", str(tmp_path))
    assert converter_data_dir() == tmp_path


def test_converter_data_dir_default_under_user_data(monkeypatch, tmp_path):
    monkeypatch.delenv("OFFIPY_CONVERTER_DATA_DIR", raising=False)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    assert converter_data_dir() == tmp_path / "converter"


class _FixedDatetime:
    """钉住 default_save_path 的时间戳，让断言确定。"""

    @classmethod
    def now(cls, tz=None):
        return _real_datetime(2026, 8, 5, 9, 15, 30, tzinfo=timezone.utc)


def test_default_save_path_sanitizes_name_and_stamps(monkeypatch, tmp_path):
    # P1-3：默认落盘不依赖调用方 CWD，统一写用户数据目录/documents
    monkeypatch.setattr("offipy.paths.datetime", _FixedDatetime)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    docs = tmp_path / "documents"
    expected = str(docs / "工作簿1测试文档_20260805_091530_000000.xlsx")
    assert default_save_path('工作簿1"测试/文档', ".xlsx") == expected


def test_default_save_path_fully_sanitized_falls_back_document(monkeypatch, tmp_path):
    monkeypatch.setattr("offipy.paths.datetime", _FixedDatetime)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    docs = tmp_path / "documents"
    expected = str(docs / "document_20260805_091530_000000.xlsx")
    assert default_save_path("//\\\\", ".xlsx") == expected


def test_default_save_path_creates_documents_dir(monkeypatch, tmp_path):
    # 目录自动创建：用户数据目录/documents 不存在时 mkdir(parents=True)
    monkeypatch.setattr("offipy.paths.datetime", _FixedDatetime)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    out = default_save_path("报告", ".pptx")
    assert Path(out).parent == tmp_path / "documents"
    assert Path(out).exists() is False  # 只建目录，不建文件（SaveAs 时才写）


class _AdvancingDatetime:
    """每次 now() 微秒前进，模拟真实时钟推进。"""

    _counter = 0

    @classmethod
    def now(cls, tz=None):
        cls._counter += 1
        return _real_datetime(2026, 8, 5, 9, 15, 30, cls._counter, tzinfo=timezone.utc)


def test_default_save_path_no_collision_same_second(monkeypatch, tmp_path):
    # P0-4：同一秒内连续保存不碰撞——微秒时间戳 + 存在检查循环保证唯一
    monkeypatch.setattr("offipy.paths.datetime", _AdvancingDatetime)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    p1 = default_save_path("报告", ".pptx")
    Path(p1).write_text("x")  # 占位第一个时间戳的文件（模拟刚保存过）
    p2 = default_save_path("报告", ".pptx")
    assert p2 != p1
    assert Path(p2).exists() is False
    assert p1.startswith(str(tmp_path / "documents" / "报告_"))
    assert p2.startswith(str(tmp_path / "documents" / "报告_"))


def test_default_save_path_same_timestamp_collision_appends_suffix(monkeypatch, tmp_path):
    # P0-4：Windows 时钟 ~1ms，datetime.now() 连续两次可能同值。文件已存在时
    # 不再只重取时间戳（永远同值），而是追加 _<n> 序号保证唯一
    monkeypatch.setattr("offipy.paths.datetime", _FixedDatetime)  # 恒定时钟，模拟分辨率低
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    p1 = default_save_path("报告", ".pptx")
    Path(p1).write_text("x")  # 模拟第一份已保存
    p2 = default_save_path("报告", ".pptx")
    assert p2 != p1
    assert p2.endswith("报告_20260805_091530_000000_1.pptx")
    assert Path(p2).exists() is False


def test_default_save_path_mkdir_failure_raises_offipy_error(monkeypatch):
    # P0-4：默认保存目录创建失败要报明确错误，不再静默吞掉
    from offipy.exceptions import OffipyError

    def _raise_mkdir(self, *a, **k):
        raise OSError("permission denied")

    monkeypatch.setattr("offipy.paths.Path.mkdir", _raise_mkdir)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: Path("C:/no/such/dir"))
    with pytest.raises(OffipyError, match="无法创建默认保存目录"):
        default_save_path("报告", ".pptx")
