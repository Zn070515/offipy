"""HTML → 可编辑 PPTX 落地管线。

Claude 写 16:9 HTML 幻灯片 → third_party 的 vector-first 转换器（Playwright
实测 DOM 坐标）转成原生可编辑 .pptx → 通过常驻 server 在真实 PowerPoint 里
打开实况展示（页面在动），并逐页导出 PNG 供 Claude 视觉迭代。

设计系统：render 支持 `theme=` 注入内置主题（见 design.py）。Claude 在 HTML
的 <head> 写 `<style data-theme="<name>"></style>` 占位，render 时替换成主题
CSS；同一份 HTML 换主题即换皮，内容与视觉解耦。
"""

import os
import subprocess
import sys
from pathlib import Path

from .client import call, ensure_server
from .design import inject_theme

# deck.py 位于 src/offipy/，项目根需上溯三级
ROOT = Path(__file__).resolve().parent.parent.parent
CONVERT_PY = ROOT / "third_party" / "html-to-editable-pptx" / "convert.py"


def _convert_cmd(
    html: str, out: str | None, only_slides: list[int] | None, no_visual_audit: bool
) -> list[str]:
    cmd = [sys.executable, str(CONVERT_PY), str(html)]
    if out:
        cmd += ["--out", str(out)]
    if only_slides:
        cmd += ["--only-slides", ",".join(str(i) for i in only_slides)]
    if no_visual_audit:
        cmd += ["--no-visual-audit"]
    return cmd


def _default_out(html: str) -> str:
    """对齐 convert.py 的默认命名：foo.audited.html → foo.pptx，否则 foo.pptx。"""
    p = Path(html)
    if p.name.endswith(".audited.html"):
        return str(p.with_name(p.name[: -len(".audited.html")] + ".pptx"))
    return str(p.with_suffix(".pptx"))


def render(
    html: str,
    out: str | None = None,
    only_slides: list[int] | None = None,
    no_visual_audit: bool = False,
    timeout: int = 600,
    theme: str | None = None,
) -> str:
    """跑完整转换管线，返回产出 .pptx 的绝对路径。

    theme 给定时把内置主题 CSS 注入 HTML 再转换（见 design.inject_theme）；
    输出路径仍基于原 html 名。注入副本是临时文件，转换后删除。
    """
    html = os.path.abspath(html)
    if not os.path.exists(html):
        raise FileNotFoundError(html)
    target = html
    tmp_html = None
    if theme:
        with open(html, encoding="utf-8") as f:
            content = f.read()
        # 以 .audited.html 结尾 → convert 跳过 work-copy 分支，不留 .audited 残留。
        # 输出名必须显式锁到原 html 的默认名，否则 convert 会用临时文件命名的 .pptx。
        out = out or _default_out(html)
        tmp_html = os.path.join(
            os.path.dirname(html), f".{os.path.basename(html)}.{theme}.tmp.audited.html"
        )
        with open(tmp_html, "w", encoding="utf-8") as f:
            f.write(inject_theme(content, theme))
        target = tmp_html
    try:
        cmd = _convert_cmd(target, out, only_slides, no_visual_audit)
        try:
            r = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as e:
            so, se = e.stdout, e.stderr
            out = so.decode("utf-8", errors="replace") if isinstance(so, bytes) else (so or "")
            err = se.decode("utf-8", errors="replace") if isinstance(se, bytes) else (se or "")
            raise RuntimeError(f"convert.py 超时 ({timeout}s)\n{out}\n{err}") from e
        if r.returncode != 0:
            raise RuntimeError(f"convert.py 失败 (exit {r.returncode})\n{r.stdout}\n{r.stderr}")
        pptx = os.path.abspath(out) if out else _default_out(html)
        if not os.path.exists(pptx):
            raise FileNotFoundError(f"转换未产出 .pptx: {pptx}\n{r.stdout}\n{r.stderr}")
        return pptx
    finally:
        if tmp_html and os.path.exists(tmp_html):
            os.unlink(tmp_html)


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
    html: str,
    out: str | None = None,
    open_live_flag: bool = True,
    feedback_dir: str | None = None,
    theme: str | None = None,
) -> str:
    """render → （可选）打开实况 → （可选）导出 PNG 反馈。返回 .pptx 绝对路径。

    theme 给定时注入内置主题 CSS（见 design.py）后再转换。
    """
    pptx = render(html, out, theme=theme)
    if feedback_dir:
        # 导出必须基于本次渲染的 deck：先确保打开它，再逐页导出
        open_live(pptx)
        export_slides(feedback_dir)
    elif open_live_flag:
        open_live(pptx)
    return pptx
