from pathlib import Path

from offipy import frozen_cli, frozen_converter, frozen_mcp, frozen_server


def test_frozen_cli_dispatches_to_cli(monkeypatch):
    calls = []
    monkeypatch.setattr(frozen_cli.cli, "main", lambda argv=None: calls.append(argv) or 7)

    assert frozen_cli.main(["--help"]) == 7
    assert calls == [["--help"]]


def test_frozen_mcp_dispatches_to_mcp_server(monkeypatch):
    calls = []
    monkeypatch.setattr(frozen_mcp.mcp_server, "main", lambda: calls.append(True))

    assert frozen_mcp.main() is None
    assert calls == [True]


def test_frozen_server_dispatches_arguments(monkeypatch):
    calls = []
    monkeypatch.setattr(frozen_server.server, "main", lambda argv=None: calls.append(argv))

    assert frozen_server.main(["--port", "8891"]) is None
    assert calls == [["--port", "8891"]]


def test_frozen_converter_runs_vendored_script(monkeypatch, tmp_path):
    converter_dir = tmp_path / "converter"
    converter_dir.mkdir()
    script = converter_dir / "convert.py"
    script.write_text("", encoding="utf-8")
    calls = []
    monkeypatch.setattr(frozen_converter, "_CONVERTER_DIR", converter_dir)
    monkeypatch.setattr(
        frozen_converter.runpy,
        "run_path",
        lambda path, run_name: calls.append((path, run_name)),
    )

    assert frozen_converter.main() is None
    assert calls == [(str(script), "__main__")]
    assert str(converter_dir) in frozen_converter.sys.path

