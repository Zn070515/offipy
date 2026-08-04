"""PyPI 冒烟：从指定 index（默认 TestPyPI）装指定版本，验证可导入/版本/check。

用法:
    python scripts/pypi_smoke.py [--index test.pypi.org] [--version 0.9.0a1]

验证点（任一失败 → 非 0 退出，供 TestPyPI 预发布门禁用）：
  1. uv pip install 从 index 解析并装进干净 venv
  2. `import offipy` 成功，__version__ 与 --version 一致
  3. `offipy check --json` 能跑，输出合法 JSON

注：不断言 check 退出码——Chromium/Office 就绪是运行环境问题；只证明包本身可装可跑。
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

CHECK_JSON = (
    "import json, offipy\n"
    "assert offipy.__version__ == {version!r}, (offipy.__version__, {version!r})\n"
    "import offipy.cli\n"
    "rc = offipy.cli.main(['check', '--json'])\n"
    "print(json.dumps({'rc': rc, 'version': offipy.__version__}))\n"
)


def _venv_python(venv: Path) -> str:
    return str(venv / ("Scripts/python.exe" if os.name == "nt" else "bin/python"))


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


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--index", default="https://test.pypi.org/simple", help="pip index 根 URL")
    parser.add_argument("--version", required=True, help="要安装验证的版本，如 0.9.0a1")
    args = parser.parse_args()
    print(f"[pypi-smoke] index = {args.index}, version = {args.version}")

    tmp = Path(tempfile.mkdtemp(prefix="offipy-pypi-smoke-"))
    try:
        venv = tmp / "venv"
        print("[pypi-smoke] 创建干净 venv ...")
        _run(["uv", "venv", str(venv)])
        py = _venv_python(venv)

        print("[pypi-smoke] uv pip install offipy=={args.version} ...")
        _run(
            [
                "uv",
                "pip",
                "install",
                "--python",
                py,
                "--index-url",
                args.index,
                "--extra-index-url",
                "https://pypi.org/simple",
                "--index-strategy",
                "unsafe-best-match",
                f"offipy=={args.version}",
            ]
        )

        print("[pypi-smoke] import + 版本 + check 可跑 ...")
        r = _run([py, "-c", CHECK_JSON.format(version=args.version)])
        report = json.loads(r.stdout.strip().splitlines()[-1])
        print(f"[pypi-smoke] version = {report['version']}, check rc = {report['rc']}")

        print("[pypi-smoke] OK — 从 index 安装冒烟通过")
        return 0
    finally:
        print(f"[pypi-smoke] 清理 {tmp}")
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
