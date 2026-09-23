import sys
from pathlib import Path

from offipy import runtime


def test_dev_server_command_uses_python_module(monkeypatch):
    monkeypatch.setattr(runtime, "is_packaged", lambda: False)

    assert runtime.server_command(8891) == [
        sys.executable,
        "-m",
        "offipy.server",
        "--port",
        "8891",
    ]


def test_dev_converter_command_uses_vendored_script(monkeypatch, tmp_path):
    monkeypatch.setattr(runtime, "is_packaged", lambda: False)
    html = tmp_path / "deck.html"

    command = runtime.converter_command(html)

    assert command[0] == sys.executable
    assert Path(command[1]) == runtime.converter_script()
    assert command[2] == str(html)


def test_packaged_commands_use_sibling_helpers(monkeypatch, tmp_path):
    monkeypatch.setattr(runtime, "is_packaged", lambda: True)
    monkeypatch.setattr(runtime, "resource_root", lambda: tmp_path)

    assert runtime.server_command(8891) == [str(tmp_path / "Offipy.Server.exe"), "--port", "8891"]
    assert runtime.converter_command("deck.html") == [
        str(tmp_path / "Offipy.Converter.exe"),
        "deck.html",
    ]


def test_packaged_chromium_path_is_reported_only_when_bundled(monkeypatch, tmp_path):
    monkeypatch.setattr(runtime, "is_packaged", lambda: True)
    monkeypatch.setattr(runtime, "resource_root", lambda: tmp_path)

    assert runtime.chromium_path() is None
    chromium = tmp_path / "chromium"
    chromium.mkdir()
    assert runtime.chromium_path() == chromium
