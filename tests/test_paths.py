"""用户数据路径测试：env 覆盖 + 各平台回退。"""

import sys
from pathlib import Path

from offipy.paths import converter_data_dir, user_data_dir


def test_user_data_dir_windows_uses_localappdata(monkeypatch):
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setenv("LOCALAPPDATA", r"C:\Users\test\AppData\Local")
    assert user_data_dir() == Path(r"C:\Users\test\AppData\Local\offipy")


def test_user_data_dir_windows_fallback_home(monkeypatch):
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    monkeypatch.setattr("offipy.paths.Path.home", lambda: Path(r"C:\Users\test"))
    assert user_data_dir() == Path(r"C:\Users\test\.offipy")


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
