"""HTML → 可编辑 PPTX 落地管线。

Claude 写 16:9 HTML 幻灯片 → third_party 的 vector-first 转换器（Playwright
实测 DOM 坐标）转成原生可编辑 .pptx → 通过常驻 server 在真实 PowerPoint 里
打开实况展示（页面在动），并逐页导出 PNG 供 Claude 视觉迭代。
"""

import os
import subprocess
import sys
from pathlib import Path

from .client import call, ensure_server

# deck.py 位于 src/office_kit/，项目根需上溯三级
ROOT = Path(__file__).resolve().parent.parent.parent
CONVERT_PY = ROOT / "third_party" / "html-to-editable-pptx" / "convert.py"


def _convert_cmd(html: str, out: str | None, only_slides, no_visual_audit: bool) -> list[str]:
    cmd = [sys.executable, str(CONVERT_PY), str(html)]
    if out:
        cmd += ["--out", str(out)]
    if only_slides:
        cmd += ["--only-slides", ",".join(str(i) for i in only_slides)]
    if no_visual_audit:
        cmd += ["--no-visual-audit"]
    return cmd


def render(
    html: str,
    out: str | None = None,
    only_slides=None,
    no_visual_audit: bool = False,
    timeout: int = 600,
) -> str:
    """跑完整转换管线，返回产出 .pptx 的绝对路径。"""
    html = os.path.abspath(html)
    if not os.path.exists(html):
        raise FileNotFoundError(html)
    cmd = _convert_cmd(html, out, only_slides, no_visual_audit)
    r = subprocess.run(
        cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout
    )
    if r.returncode != 0:
        raise RuntimeError(f"convert.py 失败 (exit {r.returncode})\n{r.stdout}\n{r.stderr}")
    pptx = os.path.abspath(out) if out else str(Path(html).with_suffix(".pptx"))
    if not os.path.exists(pptx):
        raise FileNotFoundError(f"转换未产出 .pptx: {pptx}\n{r.stdout}\n{r.stderr}")
    return pptx


def open_live(pptx: str) -> None:
    """在真实 PowerPoint 里打开生成的 .pptx（实况展示）。"""
    ensure_server()
    call("ppt", "open_pres", path=os.path.abspath(pptx))


def export_slides(out_dir: str, width: int = 1920, height: int = 1080) -> list[str]:
    """把当前实况演示文稿逐页导出 PNG，供 Claude 视觉迭代。"""
    ensure_server()
    return call(
        "ppt", "export_slides", out_dir=os.path.abspath(out_dir), width=width, height=height
    )


def make(
    html: str, out: str | None = None, open_live_flag: bool = True, feedback_dir: str | None = None
) -> str:
    """render → （可选）打开实况 → （可选）导出 PNG 反馈。返回 .pptx 绝对路径。"""
    pptx = render(html, out)
    if open_live_flag:
        open_live(pptx)
    if feedback_dir:
        export_slides(feedback_dir)
    return pptx
