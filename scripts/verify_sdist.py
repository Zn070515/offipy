"""sdist 重建验证（P0-6）：解包 dist/ 最新 sdist，断言关键发布文件存在，
在全新目录里只凭 sdist 重建 wheel，再对该 wheel 跑 verify_wheel.py 实证。

任一断言/重建失败 → 非 0 退出。证明「sdist 自足：只凭它 + 网络上的 build
依赖就能重建出同内容的 wheel」，防止发布时 sdist 与 wheel 来源漂移。

用法: python scripts/verify_sdist.py [sdist 路径]
"""

from __future__ import annotations

import argparse
import glob
import shutil
import subprocess
import sys
import tarfile
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# sdist 必须携带的发布关键文件（重建 wheel 的输入 + 许可/类型标记）
_NEEDED_IN_SDIST = (
    "pyproject.toml",
    "README.md",
    "LICENSE",
    "THIRD_PARTY_NOTICES.md",
    "src/offipy/py.typed",
    "src/offipy/_vendor/html_to_editable_pptx/LICENSE",
    "src/offipy/_vendor/diagram-design/LICENSE",
    "src/offipy/_vendor/diagram-design/THIRD_PARTY_LICENSES.md",
    "src/offipy/assets/icons/manifest.json",
)


def _find_sdist(explicit: str | None) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            raise SystemExit(f"找不到 sdist: {p}")
        return p
    sdists = sorted(glob.glob(str(ROOT / "dist" / "offipy-*.tar.gz")))
    if not sdists:
        raise SystemExit("dist/ 里没有 sdist，先跑 uv build")
    return Path(sdists[-1])


def main() -> int:
    parser = argparse.ArgumentParser(description="从 sdist 独立重建 wheel 并实证。")
    parser.add_argument("sdist", nargs="?", help="sdist 路径（缺省取 dist/ 最新）")
    args = parser.parse_args()

    sdist = _find_sdist(args.sdist)
    print(f"[verify-sdist] 验证 {sdist.name}（{sdist.stat().st_size:,} bytes）")

    work = Path(tempfile.mkdtemp(prefix="offipy-sdist-"))
    try:
        extract = work / "extract"
        extract.mkdir()
        with tarfile.open(sdist, "r:gz") as t:
            try:
                t.extractall(extract, filter="data")  # py3.11.4+/3.12 防路径穿越
            except TypeError:
                t.extractall(extract)  # py3.10 无 filter 参数
        subdirs = [p for p in extract.iterdir() if p.is_dir()]
        if len(subdirs) != 1:
            raise SystemExit(f"[verify-sdist] FAIL: sdist 顶层目录数 != 1（{subdirs}）")
        root = subdirs[0]

        missing = [n for n in _NEEDED_IN_SDIST if not (root / n).exists()]
        if missing:
            raise SystemExit(f"[verify-sdist] FAIL: sdist 缺 {missing}")
        for n in _NEEDED_IN_SDIST:
            print(f"[verify-sdist] sdist 含 {n} ✓")

        rebuild = work / "rebuild"
        shutil.copytree(root, rebuild)
        print("[verify-sdist] 在全新目录从 sdist 重建 wheel ...")
        proc = subprocess.run(
            ["uv", "build"],
            cwd=str(rebuild),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if proc.returncode != 0:
            print(proc.stdout, file=sys.stderr)
            print(proc.stderr, file=sys.stderr)
            raise SystemExit("[verify-sdist] FAIL: 从 sdist 重建 wheel 失败")

        rebuilt = sorted(glob.glob(str(rebuild / "dist" / "offipy-*.whl")))
        if not rebuilt:
            raise SystemExit("[verify-sdist] FAIL: 重建未产出 wheel")

        verifier = ROOT / "scripts" / "verify_wheel.py"
        vp = subprocess.run(
            [sys.executable, str(verifier), rebuilt[-1]],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        print(vp.stdout, end="")
        if vp.returncode != 0:
            print(vp.stderr, file=sys.stderr)
            raise SystemExit("[verify-sdist] FAIL: 重建 wheel 实证未通过")
    finally:
        shutil.rmtree(work, ignore_errors=True)

    print(f"[verify-sdist] OK — {sdist.name} 重建并实证通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
