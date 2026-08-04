"""安装冒烟：在干净 venv 里安装构建好的 wheel，验证可导入/版本/converter 存在/check 可跑。

用法:
    python scripts/install_smoke.py [path/to/offipy-*.whl]
    # 缺省时取 dist/ 里最新构建的 wheel

验证点（任一失败 → 非 0 退出，供 CI release 门禁用）：
  1. uv pip install 能解析依赖并装进干净 venv（依赖走 uv 全局缓存，无需重下载）
  2. `import offipy` 成功，__version__ 非空
  3. offipy.deck.CONVERT_PY 存在（P0-1：converter vendored 进 wheel）
  4. `offipy check --json` 能跑，输出合法 JSON 且 version 匹配

注：不断言 check 的退出码——Chromium/Office 是否就绪是运行环境问题，
不是打包问题；这里只证明打包本身正确。
"""

from __future__ import annotations

import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CHECK_JSON = (
    "import json, offipy\n"
    "from offipy.deck import CONVERT_PY\n"
    "assert CONVERT_PY.exists(), CONVERT_PY\n"
    "assert CONVERT_PY.is_file(), CONVERT_PY\n"
    "import offipy.cli\n"
    "rc = offipy.cli.main(['check', '--json'])\n"
    "print(json.dumps({'rc': rc, 'version': offipy.__version__}))\n"
)


def _pick_wheel(explicit: str | None) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            raise SystemExit(f"找不到 wheel: {p}")
        return p
    wheels = sorted(glob.glob(str(ROOT / "dist" / "offipy-*.whl")))
    if not wheels:
        raise SystemExit("dist/ 里没有 wheel，先跑 uv build")
    return Path(wheels[-1])


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
    wheel = _pick_wheel(sys.argv[1] if len(sys.argv) > 1 else None)
    print(f"[install-smoke] wheel = {wheel}")

    tmp = Path(tempfile.mkdtemp(prefix="offipy-smoke-"))
    try:
        venv = tmp / "venv"
        print("[install-smoke] 创建干净 venv ...")
        _run(["uv", "venv", str(venv)])
        py = _venv_python(venv)

        print("[install-smoke] uv pip install wheel（依赖走 uv 缓存）...")
        _run(["uv", "pip", "install", "--python", py, str(wheel)])

        print("[install-smoke] import + 版本 + converter 存在 + check 可跑 ...")
        r = _run([py, "-c", CHECK_JSON])
        report = json.loads(r.stdout.strip().splitlines()[-1])
        print(f"[install-smoke] version = {report['version']}, check rc = {report['rc']}")
        if not report["version"]:
            raise SystemExit("[install-smoke] 版本为空")

        print("[install-smoke] OK — 打包冒烟通过")
        return 0
    finally:
        print(f"[install-smoke] 清理 {tmp}")
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
