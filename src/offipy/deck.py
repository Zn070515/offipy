"""HTML → 可编辑 PPTX 落地管线。

Claude 写 16:9 HTML 幻灯片 → vendored 的 vector-first 转换器（Playwright
实测 DOM 坐标）转成原生可编辑 .pptx → 通过常驻 server 在真实 PowerPoint 里
打开实况展示（页面在动），并逐页导出 PNG 供 Claude 视觉迭代。

设计系统：render 支持 `theme=` 注入内置主题（见 design.py）。Claude 在 HTML
的 <head> 写 `<style data-theme="<name>"></style>` 占位，render 时替换成主题
CSS；同一份 HTML 换主题即换皮，内容与视觉解耦。
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path

from .client import call, ensure_server
from .design import inject_theme
from .exceptions import ConversionError
from .layouts import inject_layouts
from .paths import converter_data_dir

# vendored 转换器位于包内 _vendor/，site-packages 下经 __file__ 自定位
_CONVERT_DIR = Path(__file__).resolve().parent / "_vendor" / "html_to_editable_pptx"
CONVERT_PY = _CONVERT_DIR / "convert.py"


def _preflight_browser() -> None:
    """渲染前置：确保 chromium 可用，否则给安装提示（P0-2）。"""
    from .envcheck import _check_browser

    check = _check_browser()
    if not check.ok:
        raise ConversionError(
            "HTML→PPTX 渲染需要 Chromium："
            f"{check.detail}。请运行: python -m playwright install chromium"
        )


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
    apply_layouts: bool = False,
    overwrite: bool = False,
) -> str:
    """跑完整转换管线，返回产出 .pptx 的绝对路径。

    theme 给定时把内置主题 CSS 注入 HTML 再转换（见 design.inject_theme）；
    apply_layouts 给定时把 HTML 里 data-layout 引用的布局 CSS 注入
    （见 layouts.inject_layouts）。两者可叠加，输出路径仍基于原 html 名。
    注入副本是临时文件，转换后删除。overwrite=False 时若输出 .pptx 已存在
    抛 FileExistsError（fail-fast，不浪费一次渲染）。

    原子替换（P0-6）：转换先写同目录临时 .pptx，图表/图标后处理作用于
    临时文件，全部成功后才 os.replace 一步替换目标。任何失败/异常都会
    清理临时文件——已存在的 .pptx 绝不因一次失败的渲染被破坏。
    """
    html = os.path.abspath(html)
    if not os.path.exists(html):
        raise FileNotFoundError(html)
    _preflight_browser()
    final_out = os.path.abspath(out) if out else _default_out(html)
    if not overwrite and os.path.exists(final_out):
        raise FileExistsError(
            f"输出 .pptx 已存在（overwrite=False，可传 overwrite=True 覆盖）: {final_out}"
        )
    target = html
    tmp_html = None
    if theme or apply_layouts:
        with open(html, encoding="utf-8") as f:
            content = f.read()
        if apply_layouts:
            content = inject_layouts(content)
        if theme:
            content = inject_theme(content, theme)
        # 以 .audited.html 结尾 → convert 跳过 work-copy 分支，不留 .audited 残留。
        tmp_html = os.path.join(
            os.path.dirname(html),
            f".{os.path.basename(html)}.{theme or 'layouts'}.tmp.audited.html",
        )
        with open(tmp_html, "w", encoding="utf-8") as f:
            f.write(content)
        target = tmp_html
    # 原子替换：与最终输出同目录的临时 .pptx（同卷，保证 os.replace 原子）
    tmp_pptx = os.path.join(
        os.path.dirname(final_out), f".{Path(final_out).stem}.{theme or 'deck'}.tmp.pptx"
    )
    try:
        cmd = _convert_cmd(target, tmp_pptx, only_slides, no_visual_audit)
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"  # 中文 Windows 下 convert.py 输出才不会乱码
        # 转换器可变数据（配置/lessons-learned）落用户数据目录，不写包内
        env["OFFIPY_CONVERTER_DATA_DIR"] = str(converter_data_dir())
        try:
            r = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                env=env,
            )
        except subprocess.TimeoutExpired as e:
            so, se = e.stdout, e.stderr
            out = so.decode("utf-8", errors="replace") if isinstance(so, bytes) else (so or "")
            err = se.decode("utf-8", errors="replace") if isinstance(se, bytes) else (se or "")
            raise ConversionError(f"convert.py 超时 ({timeout}s)\n{out}\n{err}") from e
        if r.returncode != 0:
            raise ConversionError(f"convert.py 失败 (exit {r.returncode})\n{r.stdout}\n{r.stderr}")
        if not os.path.exists(tmp_pptx):
            raise ConversionError(f"转换未产出 .pptx: {tmp_pptx}\n{r.stdout}\n{r.stderr}")
        # 图表后处理：HTML 声明了 data-chart → 读 measurements 替换成原生图表。
        # 惰性 import：charts.py 内部 import python-pptx，不拖慢无图表的路径。
        from .charts import postprocess_charts

        postprocess_charts(html, tmp_pptx)
        # 图标后处理：HTML 声明了 data-icon → 替换成 freeform 矢量图标（同 charts 架构）。
        # 惰性 import：icons.py 内部 import python-pptx，不拖慢无图标的路径。
        from .icons import postprocess_icons

        postprocess_icons(html, tmp_pptx)
        os.replace(tmp_pptx, final_out)
        return final_out
    finally:
        # 任何路径（成功或失败）都清理临时文件；已存在的 final_out 不受影响。
        # convert 的 <out>_audit 审计目录跟着临时 .pptx 名字走，tmp 被替换/删除后
        # 就成了孤儿（charts 后处理在 os.replace 前已读完 measurements），一并清掉。
        for p in (tmp_pptx, tmp_html):
            if p and os.path.exists(p):
                os.unlink(p)
        tmp_audit = os.path.join(os.path.dirname(tmp_pptx), f"{Path(tmp_pptx).stem}_audit")
        if os.path.isdir(tmp_audit):
            shutil.rmtree(tmp_audit, ignore_errors=True)


def open_live(pptx: str) -> None:
    """在真实 PowerPoint 里打开生成的 .pptx（实况展示）。"""
    ensure_server()
    call("ppt", "open_pres", path=os.path.abspath(pptx))


def export_slides(out_dir: str, width: int = 1920, height: int = 1080) -> list[str]:
    """把当前实况演示文稿逐页导出 PNG，供 Claude 视觉迭代（每次重新导出全部页，允许覆盖）。"""
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
    apply_layouts: bool = False,
    overwrite: bool = False,
) -> str:
    """render → （可选）打开实况 → （可选）导出 PNG 反馈。返回 .pptx 绝对路径。

    theme 给定时注入内置主题 CSS（见 design.py）；apply_layouts 给定时注入
    data-layout 布局 CSS（见 layouts.py），两者可叠加。overwrite 透传给 render。
    """
    pptx = render(html, out, theme=theme, apply_layouts=apply_layouts, overwrite=overwrite)
    if feedback_dir:
        # 导出必须基于本次渲染的 deck：先确保打开它，再逐页导出
        open_live(pptx)
        export_slides(feedback_dir)
    elif open_live_flag:
        open_live(pptx)
    return pptx
