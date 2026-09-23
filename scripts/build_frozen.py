"""Build the Windows no-Python offipy executables with PyInstaller.

The build environment may use Python, but the resulting artifacts do not. The
four executables deliberately share one output directory because packaged
``offipy.runtime`` resolves helper processes next to the main executable.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE_ROOT = ROOT / "src"
CONVERTER_DIR = SOURCE_ROOT / "offipy" / "_vendor" / "html_to_editable_pptx"

ARTIFACTS = {
    "Offipy.exe": "offipy.frozen_cli",
    "Offipy.MCP.exe": "offipy.frozen_mcp",
    "Offipy.Server.exe": "offipy.frozen_server",
    "Offipy.Converter.exe": "offipy.frozen_converter",
}


class FrozenBuildError(RuntimeError):
    """Raised when a frozen build cannot be produced or verified."""


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

    command.append(str(entrypoint))
    return command


def build(output_dir: Path, clean: bool = True) -> list[Path]:
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
    args = parser.parse_args(argv)
    for artifact in build(args.output_dir, clean=not args.no_clean):
        print(f"built {artifact}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
