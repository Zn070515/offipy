"""TestPyPI 精确安装门禁（P0-5）：从 index 精确下载 wheel → sha256 比对 → 干净 venv 安装 → 冒烟。

任一断言失败 → 非 0 退出。验证点：
- `uv pip download` 从 index 精确拿到 offipy==<version> 的 wheel（无 --no-deps 重解析，
  下载产物与构建 artifact 由 --expected-sha256 比对兜底，防「旧同版本被 skip」）
- 新建干净 venv：`--no-deps` 装下载的 wheel，再从正式 PyPI 解析 [deck,mcp] extras
  运行时依赖（Linux 上 pywin32 因 platform_system marker 自动跳过）
- `import offipy` / __version__ == --version / `offipy --help` /
  `offipy check --profile all --json`（合法 JSON 且 version 匹配）/
  `offipy mcp --help`

CLI:
    --version 必填；--index 默认 https://test.pypi.org/simple；
    --expected-sha256 给了才比对（构建 artifact 的 sha，门禁用）。

注：不断言 check 的退出码——Chromium/Office 是否就绪是运行环境问题，
不是打包问题；这里只证明「TestPyPI 上的 wheel == 构建产物，且可装可跑」。
"""

from __future__ import annotations

import argparse
import glob
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

_DEFAULT_INDEX = "https://test.pypi.org/simple"


def _venv_python(venv: Path) -> str:
    return str(venv / ("Scripts/python.exe" if os.name == "nt" else "bin/python"))


def _venv_script(venv: Path, name: str) -> str:
    return str(venv / ("Scripts" if os.name == "nt" else "bin") / name)


def _run(cmd: list[str]) -> subprocess.CompletedProcess:
    # 子进程统一 UTF-8：Windows 默认 GBK 会把中文输出打乱/解码失败
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    return subprocess.run(
        cmd,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _find_wheel(download_dir: Path) -> Path:
    wheels = sorted(glob.glob(str(download_dir / "offipy-*.whl")))
    if not wheels:
        raise SystemExit(f"[pypi-smoke] index 下载目录里没有 wheel: {download_dir}")
    return Path(wheels[-1])


def main() -> int:
    parser = argparse.ArgumentParser(description="TestPyPI 精确安装门禁。")
    parser.add_argument("--version", required=True, help="要下载安装验证的版本，如 0.9.0a1")
    parser.add_argument(
        "--index",
        default=_DEFAULT_INDEX,
        help=f"index 根 URL（默认 {_DEFAULT_INDEX}）",
    )
    parser.add_argument(
        "--expected-sha256",
        help="构建 artifact 的 wheel sha256；给了则断言下载产物同一内容",
    )
    args = parser.parse_args()
    print(f"[pypi-smoke] index = {args.index}, version = {args.version}")

    tmp = Path(tempfile.mkdtemp(prefix="offipy-pypi-smoke-"))
    try:
        venv = tmp / "venv"
        download = tmp / "download"
        download.mkdir()
        print("[pypi-smoke] 创建干净 venv（用于解析 wheel 平台 + 安装）...")
        _run(["uv", "venv", str(venv)])
        py = _venv_python(venv)

        print(f"[pypi-smoke] uv pip download offipy=={args.version}（精确下载）...")
        _run(
            [
                "uv",
                "pip",
                "download",
                "--python",
                py,
                "--index-url",
                args.index,
                "--no-deps",
                "--no-binary",
                ":none:",
                "-d",
                str(download),
                f"offipy=={args.version}",
            ]
        )
        wheel = _find_wheel(download)
        print(f"[pypi-smoke] 下载到 {wheel.name}")

        sha = _sha256(wheel)
        if args.expected_sha256:
            if sha != args.expected_sha256.lower():
                raise SystemExit(
                    f"[pypi-smoke] FAIL: wheel sha256 {sha} != 构建产物 "
                    f"{args.expected_sha256}（TestPyPI 上不是本轮构建内容）"
                )
            print("[pypi-smoke] sha256 == 构建产物 ✓")
        else:
            print(f"[pypi-smoke] sha256 = {sha}（未给 --expected-sha256，跳过比对）")

        print("[pypi-smoke] 干净 venv：--no-deps 装 wheel + PyPI 解析 [deck,mcp] extras ...")
        # offipy[deck,mcp] 的运行时依赖从正式 PyPI 装；pywin32 由
        # platform_system=='Windows' marker 在 Linux 上自动跳过
        _run(
            [
                "uv",
                "pip",
                "install",
                "--python",
                py,
                "--index-url",
                "https://pypi.org/simple",
                f"{wheel}[deck,mcp]",
            ]
        )
        script = _venv_script(venv, "offipy")

        print("[pypi-smoke] import offipy + __version__ 匹配 ...")
        r = _run([py, "-c", "import offipy; print(offipy.__version__)"])
        if r.stdout.strip() != args.version:
            raise SystemExit(
                f"[pypi-smoke] FAIL: import offipy.__version__ = {r.stdout.strip()!r} "
                f"!= --version {args.version!r}"
            )
        print(f"[pypi-smoke] __version__ == {args.version} ✓")

        print("[pypi-smoke] offipy --help ...")
        r = _run([script, "--help"])
        if "usage" not in r.stdout.lower():
            raise SystemExit("[pypi-smoke] FAIL: offipy --help 无 usage 输出")

        print("[pypi-smoke] offipy check --profile all --json ...")
        r = _run([script, "check", "--profile", "all", "--json"])
        report = json.loads(r.stdout.strip().splitlines()[-1])
        if report.get("version") != args.version:
            raise SystemExit(
                f"[pypi-smoke] FAIL: check JSON version = {report.get('version')!r} "
                f"!= {args.version!r}"
            )
        if not isinstance(report.get("checks"), list):
            raise SystemExit("[pypi-smoke] FAIL: check JSON 缺 checks 数组")
        print(
            f"[pypi-smoke] check JSON version 匹配，checks = {len(report['checks'])} 项 "
            f"（rc 不断言，环境就绪由 office-real 承担）"
        )

        print("[pypi-smoke] offipy mcp --help ...")
        r = _run([script, "mcp", "--help"])
        if "usage" not in r.stdout.lower():
            raise SystemExit("[pypi-smoke] FAIL: offipy mcp --help 无 usage 输出")

        print(f"[pypi-smoke] OK — {args.version} 从 index 精确下载 + 干净安装冒烟通过")
        return 0
    finally:
        print(f"[pypi-smoke] 清理 {tmp}")
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
