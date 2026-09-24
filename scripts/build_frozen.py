"""Build the Windows no-Python offipy executables with PyInstaller.

The build environment may use Python, but the resulting artifacts do not. The
four executables deliberately share one output directory because packaged
``offipy.runtime`` resolves helper processes next to the main executable.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE_ROOT = ROOT / "src"
CONVERTER_DIR = SOURCE_ROOT / "offipy" / "_vendor" / "html_to_editable_pptx"
_CHROMIUM_RUNTIME_DIR = "chromium"
_HEADLESS_PREFIX = "chromium_headless_shell-"
_CHROMIUM_PREFIX = "chromium-"
_KEEP_LOCALES = frozenset({"en-US.pak", "zh-CN.pak"})

ARTIFACTS = {
    "Offipy.exe": "offipy.frozen_cli",
    "Offipy.MCP.exe": "offipy.frozen_mcp",
    "Offipy.Server.exe": "offipy.frozen_server",
    "Offipy.Converter.exe": "offipy.frozen_converter",
}


class FrozenBuildError(RuntimeError):
    """Raised when a frozen build cannot be produced or verified."""


def _browser_executable(package: Path) -> Path | None:
    for relative in (
        Path("chrome-headless-shell-win64") / "chrome-headless-shell.exe",
        Path("chrome-win64") / "chrome.exe",
    ):
        executable = package / relative
        if executable.is_file():
            return executable
    return None


def _find_browser_package(source: Path, prefer_headless: bool = True) -> Path:
    source = source.resolve()
    if not source.is_dir():
        raise FrozenBuildError(f"Chromium source 不存在: {source}")
    if _browser_executable(source) is not None:
        return source

    prefixes = (_HEADLESS_PREFIX, _CHROMIUM_PREFIX) if prefer_headless else (_CHROMIUM_PREFIX,)
    candidates = [
        child
        for child in source.iterdir()
        if child.is_dir()
        and any(child.name.startswith(prefix) for prefix in prefixes)
        and _browser_executable(child) is not None
    ]
    if not candidates:
        raise FrozenBuildError(f"Chromium source 未找到可用的 Windows Chromium package: {source}")

    def version_key(path: Path) -> tuple[int, str]:
        try:
            return (int(path.name.rsplit("-", 1)[-1]), path.name)
        except ValueError:
            return (-1, path.name)

    for prefix in prefixes:
        preferred = [path for path in candidates if path.name.startswith(prefix)]
        if preferred:
            return max(preferred, key=version_key)
    raise FrozenBuildError(f"Chromium source 未找到可用 package: {source}")


def _default_chromium_source() -> Path:
    try:
        dry_run = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "--dry-run", "chromium"],
            check=True,
            capture_output=True,
            cwd=str(ROOT),
            encoding="utf-8",
            errors="replace",
        )
    except (OSError, subprocess.CalledProcessError):
        dry_run = None
    if dry_run is not None:
        lines = str(dry_run.stdout or "").splitlines()
        for index, line in enumerate(lines):
            if "playwright chromium-headless-shell" not in line:
                continue
            for candidate in lines[index + 1 : index + 4]:
                marker = "Install location:"
                if marker in candidate:
                    return Path(candidate.split(marker, 1)[1].strip())

    configured = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    if configured and configured != "0":
        return Path(configured)
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "ms-playwright"
    return Path.home() / "AppData" / "Local" / "ms-playwright"


def _directory_size(path: Path) -> int:
    return sum(file.stat().st_size for file in path.rglob("*") if file.is_file())


def stage_chromium(
    output_dir: Path,
    source: Path | None = None,
    *,
    prefer_headless: bool = True,
    prune_locales: bool = True,
) -> Path:
    """Copy one Playwright Chromium package into the shared frozen layout."""

    package = _find_browser_package(source or _default_chromium_source(), prefer_headless)
    destination_root = Path(output_dir).resolve() / _CHROMIUM_RUNTIME_DIR
    if destination_root.exists():
        shutil.rmtree(destination_root)
    destination_root.mkdir(parents=True, exist_ok=True)
    destination_package = destination_root / package.name
    shutil.copytree(package, destination_package)

    if prune_locales:
        locales = destination_package / "chrome-headless-shell-win64" / "locales"
        if not locales.is_dir():
            locales = destination_package / "chrome-win64" / "locales"
        if locales.is_dir():
            for locale in locales.glob("*.pak"):
                if locale.name not in _KEEP_LOCALES:
                    locale.unlink()

    executable = _browser_executable(destination_package)
    if executable is None:
        raise FrozenBuildError(f"复制后的 Chromium runtime 不完整: {destination_package}")
    size_mib = _directory_size(destination_root) / (1024 * 1024)
    print(f"staged Chromium {executable} ({size_mib:.1f} MiB)")
    return destination_root


def _entrypoint_path(module: str) -> Path:
    return SOURCE_ROOT / Path(*module.split(".")).with_suffix(".py")


def _converter_hidden_imports() -> list[str]:
    return sorted(path.stem for path in CONVERTER_DIR.glob("*.py") if path.stem != "convert")


def _command(executable: str, module: str, output_dir: Path, work_dir: Path) -> list[str]:
    entrypoint = _entrypoint_path(module)
    if not entrypoint.is_file():
        raise FrozenBuildError(f"冻结入口不存在: {entrypoint}")

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--noconfirm",
        "--name",
        executable.removesuffix(".exe"),
        "--distpath",
        str(output_dir),
        "--workpath",
        str(work_dir / executable.removesuffix(".exe")),
        "--specpath",
        str(work_dir / "specs"),
        "--paths",
        str(SOURCE_ROOT),
        "--collect-submodules",
        "offipy",
        "--collect-data",
        "offipy",
    ]

    if module == "offipy.frozen_mcp":
        command.extend(
            [
                "--collect-submodules",
                "mcp.server",
                "--collect-data",
                "mcp",
                "--collect-submodules",
                "mcp_types",
                "--collect-data",
                "mcp_types",
            ]
        )
    elif module == "offipy.frozen_converter":
        command.extend(
            [
                "--paths",
                str(CONVERTER_DIR),
                "--add-data",
                f"{CONVERTER_DIR}{os.pathsep}offipy/_vendor/html_to_editable_pptx",
            ]
        )
        for hidden_import in _converter_hidden_imports():
            command.extend(["--hidden-import", hidden_import])
        for hidden_import in (
            "fontTools.subset",
            "fontTools.ttLib",
            "fontTools.varLib.instancer",
        ):
            command.extend(["--hidden-import", hidden_import])

    command.append(str(entrypoint))
    return command


def build(
    output_dir: Path,
    clean: bool = True,
    *,
    with_chromium: bool = False,
    chromium_source: Path | None = None,
    full_chromium: bool = False,
) -> list[Path]:
    """Build and verify all frozen executables in ``output_dir``."""

    if sys.platform != "win32":
        raise FrozenBuildError("CF-3 frozen build 仅支持 Windows x64")

    output = Path(output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    work_dir = output.parent / ".offipy-pyinstaller"
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    outputs: list[Path] = []
    for executable, module in ARTIFACTS.items():
        command = _command(executable, module, output, work_dir)
        if clean:
            command.insert(command.index("--noconfirm"), "--clean")
        subprocess.run(command, check=True, cwd=str(ROOT), env=env)
        artifact = output / executable
        if not artifact.is_file():
            raise FrozenBuildError(f"PyInstaller 未生成预期文件: {artifact}")
        outputs.append(artifact)
    if with_chromium or chromium_source is not None or full_chromium:
        stage_chromium(output, chromium_source, prune_locales=not full_chromium)
    return outputs


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT / "build" / "frozen",
        help="冻结 exe 输出目录（默认 build/frozen）",
    )
    parser.add_argument(
        "--no-clean",
        action="store_true",
        help="不传 PyInstaller 的 --noconfirm（默认允许覆盖同名构建产物）",
    )
    parser.add_argument(
        "--with-chromium",
        action="store_true",
        help="把 Playwright Chromium headless shell 复制到共享 chromium/ 目录",
    )
    parser.add_argument(
        "--chromium-source",
        type=Path,
        help="Chromium package 或 Playwright browser cache 路径（默认自动发现）",
    )
    parser.add_argument(
        "--full-chromium",
        action="store_true",
        help="保留 Chromium 全部 locale（默认只保留 en-US/zh-CN 以控制体积）",
    )
    args = parser.parse_args(argv)
    for artifact in build(
        args.output_dir,
        clean=not args.no_clean,
        with_chromium=args.with_chromium,
        chromium_source=args.chromium_source,
        full_chromium=args.full_chromium,
    ):
        print(f"built {artifact}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
