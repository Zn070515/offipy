"""wheel 实证（P0-5/P0-6）：解包 dist/ 最新 wheel，逐项断言发布内容。

任一断言失败 → 非 0 退出。验证点：
- vendored converter（convert.py）与 py.typed 进包
- vendored 自检资产（assets/、tests/）被剔除
- dist-info/licenses/ 齐备（根 LICENSE、THIRD_PARTY_NOTICES、两图标许可证、vendor LICENSE）
- icons manifest 进包且 SVG 数 == 3268
- wheel 体积 < 上限（默认 5MB；资产漏剔除会撑到 ~10MB）
- METADATA：version 与 __version__ 一致 / requires-python(>=3.10) /
  license("MIT AND ISC") / urls 含 Issues+Documentation
- entry_points：offipy = offipy.cli:main
- Provides-Extra 含 office/deck/mcp

用法: python scripts/verify_wheel.py [wheel 路径]
"""

from __future__ import annotations

import argparse
import glob
import json
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_ICON_COUNT = 3268
_MAX_WHEEL = 5 * 1024 * 1024


def _find_wheel(explicit: str | None) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            raise SystemExit(f"找不到 wheel: {p}")
        return p
    wheels = sorted(glob.glob(str(ROOT / "dist" / "offipy-*.whl")))
    if not wheels:
        raise SystemExit("dist/ 里没有 wheel，先跑 uv build")
    return Path(wheels[-1])


def _src_version() -> str:
    src = (ROOT / "src" / "offipy" / "__init__.py").read_text(encoding="utf-8")
    m = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', src)
    if not m:
        raise SystemExit("[verify-wheel] 无法从 __init__.py 解析 __version__")
    return m.group(1)


def _meta_field(meta: str, field: str) -> str:
    for line in meta.splitlines():
        if line.startswith(field + ":"):
            return line.split(":", 1)[1].strip()
    return ""


def main() -> int:
    parser = argparse.ArgumentParser(description="解包实证 wheel 内容与元数据。")
    parser.add_argument("wheel", nargs="?", help="wheel 路径（缺省取 dist/ 最新）")
    parser.add_argument(
        "--max-bytes", type=int, default=_MAX_WHEEL, help="wheel 体积上限（默认 5MB）"
    )
    args = parser.parse_args()

    whl = _find_wheel(args.wheel)
    size = whl.stat().st_size
    print(f"[verify-wheel] {whl.name}（{size:,} bytes）")
    if size >= args.max_bytes:
        raise SystemExit(f"[verify-wheel] FAIL: 体积 {size:,} >= 上限 {args.max_bytes:,}")
    print(f"[verify-wheel] 体积 {size:,} < 上限 {args.max_bytes:,} ✓")

    with zipfile.ZipFile(whl) as z:
        names = set(z.namelist())

        def distinfo(ext: str) -> str:
            hits = [n for n in names if n.endswith(f".dist-info/{ext}")]
            if not hits:
                raise SystemExit(f"[verify-wheel] FAIL: 缺 .dist-info/{ext}")
            return hits[0]

        # 1) vendored converter + py.typed 进包
        if "offipy/_vendor/html_to_editable_pptx/convert.py" not in names:
            raise SystemExit("[verify-wheel] FAIL: vendored convert.py 未进包")
        print("[verify-wheel] vendored convert.py 进包 ✓")
        if "offipy/py.typed" not in names:
            raise SystemExit("[verify-wheel] FAIL: py.typed 未进包")
        print("[verify-wheel] py.typed 进包 ✓")

        # 2) vendored 自检资产/测试剔除
        vpre = "offipy/_vendor/html_to_editable_pptx/"
        if any(n.startswith(vpre + "assets/") for n in names):
            raise SystemExit("[verify-wheel] FAIL: vendored assets/ 泄漏进 wheel")
        print("[verify-wheel] vendored assets/ 已剔除 ✓")
        if any(n.startswith(vpre + "tests/") for n in names):
            raise SystemExit("[verify-wheel] FAIL: vendored tests/ 泄漏进 wheel")
        print("[verify-wheel] vendored tests/ 已剔除 ✓")

        # 3) licenses 齐备（源相对路径，见 hatchling PEP 639 行为）
        lic_suffixes = {
            n.split(".dist-info/licenses/", 1)[1] for n in names if ".dist-info/licenses/" in n
        }
        for need in (
            "LICENSE",
            "THIRD_PARTY_NOTICES.md",
            "src/offipy/assets/icons/LICENSE-lucide.txt",
            "src/offipy/assets/icons/LICENSE-phosphor.txt",
            "src/offipy/_vendor/html_to_editable_pptx/LICENSE",
        ):
            if need not in lic_suffixes:
                raise SystemExit(f"[verify-wheel] FAIL: licenses 缺 {need}")
            print(f"[verify-wheel] licenses 含 {need} ✓")

        # 4) icons manifest + SVG 数
        if "offipy/assets/icons/manifest.json" not in names:
            raise SystemExit("[verify-wheel] FAIL: icons manifest.json 未进包")
        print("[verify-wheel] icons manifest.json 进包 ✓")
        # 解析一次：manifest 本身必须是合法 JSON（否则进包的是坏文件）
        json.loads(z.read("offipy/assets/icons/manifest.json").decode("utf-8"))
        svg_count = sum(
            1 for n in names if n.startswith("offipy/assets/icons/") and n.endswith(".svg")
        )
        if svg_count != _ICON_COUNT:
            raise SystemExit(f"[verify-wheel] FAIL: icons SVG 数 {svg_count} != 预期 {_ICON_COUNT}")
        print(f"[verify-wheel] icons SVG 数 {svg_count} == {_ICON_COUNT} ✓")

        # 5) METADATA
        meta = z.read(distinfo("METADATA")).decode("utf-8", errors="replace")
        if _meta_field(meta, "Version") != _src_version():
            raise SystemExit("[verify-wheel] FAIL: METADATA Version 与 __version__ 不一致")
        print(f"[verify-wheel] METADATA Version == {_src_version()} ✓")
        if _meta_field(meta, "Requires-Python") != ">=3.10":
            raise SystemExit(
                f"[verify-wheel] FAIL: Requires-Python = {_meta_field(meta, 'Requires-Python')!r}"
            )
        print("[verify-wheel] Requires-Python == >=3.10 ✓")
        if _meta_field(meta, "License-Expression") != "MIT AND ISC":
            raise SystemExit("[verify-wheel] FAIL: License-Expression != MIT AND ISC")
        print("[verify-wheel] License-Expression == MIT AND ISC ✓")
        urls = {ln for ln in meta.splitlines() if ln.startswith("Project-URL:")}
        for label in ("Issues", "Documentation"):
            if not any(f"{label}," in u for u in urls):
                raise SystemExit(f"[verify-wheel] FAIL: Project-URL 缺 {label}")
            print(f"[verify-wheel] Project-URL 含 {label} ✓")

        # 6) entry_points
        ep = z.read(distinfo("entry_points.txt")).decode("utf-8", errors="replace")
        if "offipy = offipy.cli:main" not in ep:
            raise SystemExit("[verify-wheel] FAIL: entry_points 缺 offipy = offipy.cli:main")
        print("[verify-wheel] entry_points offipy = offipy.cli:main ✓")

        # 7) Provides-Extra
        extras = {
            ln.split("Provides-Extra:", 1)[1].strip()
            for ln in meta.splitlines()
            if ln.startswith("Provides-Extra:")
        }
        for ex in ("office", "deck", "mcp"):
            if ex not in extras:
                raise SystemExit(f"[verify-wheel] FAIL: Provides-Extra 缺 {ex}")
            print(f"[verify-wheel] Provides-Extra 含 {ex} ✓")

    print(f"[verify-wheel] OK — {whl.name} 实证通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
