from pathlib import Path

import pytest
from scripts import build_frozen


def test_artifacts_have_stable_names_and_entrypoints():
    assert build_frozen.ARTIFACTS == {
        "Offipy.exe": "offipy.frozen_cli",
        "Offipy.MCP.exe": "offipy.frozen_mcp",
        "Offipy.Server.exe": "offipy.frozen_server",
        "Offipy.Converter.exe": "offipy.frozen_converter",
    }


def test_build_rejects_non_windows(monkeypatch, tmp_path):
    monkeypatch.setattr(build_frozen.sys, "platform", "linux")

    with pytest.raises(build_frozen.FrozenBuildError, match="Windows"):
        build_frozen.build(tmp_path)


def test_build_invokes_onefile_for_each_artifact(monkeypatch, tmp_path):
    monkeypatch.setattr(build_frozen.sys, "platform", "win32")
    commands = []

    def fake_run(command, **kwargs):
        commands.append((command, kwargs))
        output_dir = Path(command[command.index("--distpath") + 1])
        name = command[command.index("--name") + 1]
        output_dir.mkdir(parents=True, exist_ok=True)
        (output_dir / f"{name}.exe").write_bytes(b"fake")

    monkeypatch.setattr(build_frozen.subprocess, "run", fake_run)

    outputs = build_frozen.build(tmp_path)

    assert outputs == [tmp_path / name for name in build_frozen.ARTIFACTS]
    assert len(commands) == 4
    for command, kwargs in commands:
        assert "--onefile" in command
        assert "--noconfirm" in command
        assert "--paths" in command
        assert kwargs["check"] is True
        assert kwargs["env"]["PYTHONIOENCODING"] == "utf-8"
