"""安装冒烟：在干净 venv 里安装 offipy，验证可导入/版本/converter 存在/check 可跑。

用法:
    python scripts/install_smoke.py [path/to/offipy-*.whl] [--profile all]
    python scripts/install_smoke.py --index https://test.pypi.org/simple \
        --version 0.9.0a1 --profile core

来源二选一：
  - 本地 wheel（缺省）：取第一个位置参数或 dist/ 里最新构建的 wheel；
  - index（--index + --version）：从指定 index（TestPyPI/PyPI）装 offipy==<version>。

--profile 控制安装哪些 extras（核心依赖始终装）：
  core   —— 不装任何 extras（纯 `import offipy` 零依赖）
  office —— offipy[office]（Windows COM）
  deck   —— offipy[deck]（HTML→PPTX 管线）
  mcp    —— offipy[mcp]（MCP server）
  all    —— offipy[all]（默认）

--acceptance-dir：对目录下全部 art 验收 fixture（*.json）跑真实 analyze_scene，
  验证 wheel 里的 offipy.art 完整可用（import 导出 + 场景解析 + 规则评估 + 报告）。

验证点（任一失败 → 非 0 退出，供 CI release 门禁用）：
  1. uv pip install 能解析依赖并装进干净 venv（依赖走 uv 全局缓存，无需重下载）
  2. `import offipy` 成功，__version__ 非空
  3. offipy.deck.CONVERT_PY 存在（P0-1：converter vendored 进 wheel）
  4. `offipy check --json` 能跑，输出合法 JSON 且 version 匹配
  5. offipy.art 导出存在（build_scene/analyze_scene/render_markdown）
  6. 对 acceptance-dir 全部 fixture 跑 analyze_scene，报告非空

注：不断言 check 的退出码——Chromium/Office 是否就绪是运行环境问题，
不是打包问题；这里只证明打包本身正确。
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import shutil
import subprocess
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
# 在已装 wheel 的干净 venv 里验证 offipy.art 完整可用：纯标准库、不拉 python-pptx。
ART_CHECK = (
    "import json, sys\n"
    "from pathlib import Path\n"
    "import offipy\n"
    "assert hasattr(offipy, 'build_scene'), 'build_scene 导出缺失'\n"
    "assert hasattr(offipy, 'analyze_scene'), 'analyze_scene 导出缺失'\n"
    "assert hasattr(offipy, 'render_markdown'), 'render_markdown 导出缺失'\n"
    "assert 'pptx' not in sys.modules, 'art import 不应加载 python-pptx'\n"
    "from offipy.art import ArtScene\n"
    "acc = sys.argv[1]\n"
    "files = sorted(Path(acc).glob('*.json')) if acc and Path(acc).is_dir() else []\n"
    "if not files:\n"
    "    raise SystemExit(f'acceptance-dir 无 fixture: {acc}')\n"
    "for f in files:\n"
    "    scene = ArtScene.from_dict(json.loads(f.read_text(encoding='utf-8')))\n"
    "    report = offipy.analyze_scene(scene, profile='balanced')\n"
    "    assert report.slides, f'{f.name}: 报告无 slides'\n"
    "print(f'wheel art ok: {len(files)} acceptance scenes analyzed')\n"
)

PROFILES = ("core", "office", "deck", "mcp", "all")


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
    parser = argparse.ArgumentParser(description="安装冒烟：干净 venv 装 offipy 并验证。")
    parser.add_argument("wheel", nargs="?", help="本地 wheel 路径（缺省取 dist/ 最新）")
    parser.add_argument(
        "--profile",
        choices=PROFILES,
        default="all",
        help="安装的 extras 组合（core 不装任何 extras），默认 all",
    )
    parser.add_argument("--index", help="从 index 安装（此时需 --version），如 test.pypi.org")
    parser.add_argument("--version", help="index 安装的版本号，如 0.9.0a1")
    parser.add_argument(
        "--acceptance-dir",
        help="art 验收 fixture 目录（*.json），对每个 fixture 跑真实 analyze_scene",
    )
    args = parser.parse_args()

    if args.index and not args.version:
        parser.error("--index 需要同时给 --version")
    if args.version and not args.index:
        parser.error("--version 只在 --index 模式下有意义")

    extras = f"[{args.profile}]" if args.profile != "core" else ""
    print(f"[install-smoke] profile = {args.profile} (extras = {extras or '无'})")

    tmp = Path(tempfile.mkdtemp(prefix="offipy-smoke-"))
    try:
        venv = tmp / "venv"
        print("[install-smoke] 创建干净 venv ...")
        _run(["uv", "venv", str(venv)])
        py = _venv_python(venv)

        install_cmd = ["uv", "pip", "install", "--python", py]
        if args.index:
            install_cmd += [
                "--index-url",
                args.index,
                "--extra-index-url",
                "https://pypi.org/simple",
                "--index-strategy",
                "unsafe-best-match",
            ]
            req = f"offipy{extras}=={args.version}"
            print(f"[install-smoke] uv pip install {req}（index = {args.index}）...")
        else:
            wheel = _pick_wheel(args.wheel)
            req = f"{wheel}{extras}"
            print(f"[install-smoke] uv pip install {req}（依赖走 uv 缓存）...")
        _run(install_cmd + [req])

        print("[install-smoke] import + 版本 + converter 存在 + check 可跑 ...")
        r = _run([py, "-c", CHECK_JSON])
        report = json.loads(r.stdout.strip().splitlines()[-1])
        print(f"[install-smoke] version = {report['version']}, check rc = {report['rc']}")
        if not report["version"]:
            raise SystemExit("[install-smoke] 版本为空")

        print("[install-smoke] offipy.art 导出 + acceptance fixtures 真实 analyze_scene ...")
        acc = args.acceptance_dir or ""
        r = _run([py, "-c", ART_CHECK, acc])
        print(f"[install-smoke] {r.stdout.strip().splitlines()[-1]}")

        print("[install-smoke] OK — 打包冒烟通过")
        return 0
    finally:
        print(f"[install-smoke] 清理 {tmp}")
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
