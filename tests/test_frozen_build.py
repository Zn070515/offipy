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


def test_converter_command_includes_dynamic_fonttools_modules(tmp_path):
    command = build_frozen._command(
        "Offipy.Converter.exe",
        "offipy.frozen_converter",
        tmp_path / "dist",
        tmp_path / "work",
    )

    hidden_imports = {
        command[index + 1] for index, value in enumerate(command[:-1]) if value == "--hidden-import"
    }
    assert {
        "fontTools.subset",
        "fontTools.ttLib",
        "fontTools.varLib.instancer",
    } <= hidden_imports


def _fake_browser_cache(root: Path) -> Path:
    package = root / "chromium_headless_shell-1234" / "chrome-headless-shell-win64"
    package.mkdir(parents=True)
    (package / "chrome-headless-shell.exe").write_bytes(b"browser")
    locales = package / "locales"
    locales.mkdir()
    for name in ("en-US.pak", "zh-CN.pak", "fr.pak"):
        (locales / name).write_bytes(name.encode())
    (package / "icudtl.dat").write_bytes(b"icu")
    return root


def test_stage_chromium_prefers_headless_shell_and_prunes_locales(tmp_path):
    source = _fake_browser_cache(tmp_path / "cache")

    staged = build_frozen.stage_chromium(tmp_path / "dist", source)

    assert staged == tmp_path / "dist" / "chromium"
    assert (
        staged
        / "chromium_headless_shell-1234"
        / "chrome-headless-shell-win64"
        / "chrome-headless-shell.exe"
    ).is_file()
    locale_names = {
        path.name
        for path in (
            staged / "chromium_headless_shell-1234" / "chrome-headless-shell-win64" / "locales"
        ).iterdir()
    }
    assert locale_names == {"en-US.pak", "zh-CN.pak"}


def test_stage_chromium_full_keeps_all_locales(tmp_path):
    source = _fake_browser_cache(tmp_path / "cache")

    staged = build_frozen.stage_chromium(tmp_path / "dist", source, prune_locales=False)

    locales = staged / "chromium_headless_shell-1234" / "chrome-headless-shell-win64" / "locales"
    assert {path.name for path in locales.iterdir()} == {"en-US.pak", "zh-CN.pak", "fr.pak"}


def test_stage_chromium_rejects_missing_source(tmp_path):
    with pytest.raises(build_frozen.FrozenBuildError, match="Chromium"):
        build_frozen.stage_chromium(tmp_path / "dist", tmp_path / "missing")
