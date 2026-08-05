"""用户数据路径测试：env 覆盖 + 各平台回退 + 默认落盘路径。"""

import sys
from datetime import datetime as _real_datetime
from pathlib import Path

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
    def now(cls):
        return _real_datetime(2026, 8, 5, 9, 15, 30)


def test_default_save_path_sanitizes_name_and_stamps(monkeypatch):
    # P1-3：默认落盘不依赖调用方 CWD，统一写用户数据目录/documents
    monkeypatch.setattr("offipy.paths.datetime", _FixedDatetime)
    monkeypatch.setattr(
        "offipy.paths.user_data_dir", lambda: Path("C:/Users/test/AppData/Local/offipy")
    )
    docs = Path("C:/Users/test/AppData/Local/offipy") / "documents"
    expected = str(docs / "工作簿1测试文档_20260805_091530.xlsx")
    assert default_save_path('工作簿1"测试/文档', ".xlsx") == expected


def test_default_save_path_fully_sanitized_falls_back_document(monkeypatch):
    monkeypatch.setattr("offipy.paths.datetime", _FixedDatetime)
    monkeypatch.setattr(
        "offipy.paths.user_data_dir", lambda: Path("C:/Users/test/AppData/Local/offipy")
    )
    docs = Path("C:/Users/test/AppData/Local/offipy") / "documents"
    expected = str(docs / "document_20260805_091530.xlsx")
    assert default_save_path("//\\\\", ".xlsx") == expected


def test_default_save_path_creates_documents_dir(monkeypatch, tmp_path):
    # 目录自动创建：用户数据目录/documents 不存在时 mkdir(parents=True)
    monkeypatch.setattr("offipy.paths.datetime", _FixedDatetime)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    out = default_save_path("报告", ".pptx")
    assert Path(out).parent == tmp_path / "documents"
    assert Path(out).exists() is False  # 只建目录，不建文件（SaveAs 时才写）
