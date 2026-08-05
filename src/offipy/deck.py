"""HTML → 可编辑 PPTX 落地管线。

Claude 写 16:9 HTML 幻灯片 → vendored 的 vector-first 转换器（Playwright
实测 DOM 坐标）转成原生可编辑 .pptx → 通过常驻 server 在真实 PowerPoint 里
打开实况展示（页面在动），并逐页导出 PNG 供 Claude 视觉迭代。

设计系统：render 支持 `theme=` 注入内置主题（见 design.py）。Claude 在 HTML
的 <head> 写 `<style data-theme="<name>"></style>` 占位，render 时替换成主题
CSS；同一份 HTML 换主题即换皮，内容与视觉解耦。
"""

import contextlib
import os
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from .audit import AuditConfig, PptxAuditReport, Severity
from .client import call, ensure_server
from .design import inject_theme
from .exceptions import ConversionError, FileConflictError, InvalidArgumentError
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


def _postprocess(label: str, fn, html: str, pptx: str) -> None:
    """统一包装图表/图标后处理：解析/数据错 → InvalidArgumentError，其余 → ConversionError。

    ValueError（HTML 图表/图标声明非法、数据缺失/损坏）属用户输入问题；
    其余异常（measurements 缺失、python-pptx/XML/zip 损坏）属转换产物问题。
    均 from e 保留 __cause__。
    """
    try:
        fn(html, pptx)
    except InvalidArgumentError:
        raise
    except ConversionError:
        raise
    except ValueError as e:
        raise InvalidArgumentError(f"{label}后处理失败（HTML 声明数据非法）: {e}") from e
    except Exception as e:
        raise ConversionError(f"{label}后处理失败: {e}") from e


@contextlib.contextmanager
def _render_tmp(
    html: str,
    out: str | None,
    only_slides: list[int] | None,
    no_visual_audit: bool,
    timeout: int,
    theme: str | None,
    apply_layouts: bool,
    overwrite: bool,
) -> Iterator[tuple[str, str]]:
    """生成到临时 .pptx 的上下文：前置检查 → mkstemp → convert → 图表/图标后处理。

    yield (tmp_pptx, final_out)。调用方在 with 内负责最终 os.replace 到 final_out
    （render），或先对 tmp_pptx 审计再决定替换/拒绝（render_with_report）。
    __exit__ 无论成功失败都清理临时 .pptx、注入副本与孤儿审计目录——已存在的
    final_out 绝不被一次失败的渲染破坏。

    临时文件用 mkstemp（§7）：与最终输出同目录（同卷保证 os.replace 原子），
    随机后缀天然避开旧确定性名（`.x.tmp.pptx`）在并发渲染时的互踩竞态。
    """
    html = os.path.abspath(html)
    if not os.path.exists(html):
        raise InvalidArgumentError(f"源 HTML 文件不存在: {html}")
    _preflight_browser()
    final_out = os.path.abspath(out) if out else _default_out(html)
    if not overwrite and os.path.exists(final_out):
        raise FileConflictError(
            f"输出 .pptx 已存在（overwrite=False，可传 overwrite=True 覆盖）: {final_out}"
        )
    target = html
    tmp_html_dir = None
    if no_visual_audit:
        # 图表/图标注入依赖 visual audit 的 measurements.json；no_visual_audit 不产出
        # → 转换开始前 fail-fast，省一次白跑的 chromium 渲染。
        with open(html, encoding="utf-8") as f:
            content = f.read()
        if "data-chart" in content or "data-icon" in content:
            raise InvalidArgumentError(
                "no_visual_audit=True 但 HTML 声明了图表/图标（data-chart/data-icon）："
                "注入需要 measurements.json，请去掉 no_visual_audit 或用 visual audit 渲染"
            )
    if theme or apply_layouts:
        with open(html, encoding="utf-8") as f:
            content = f.read()
        if apply_layouts:
            content = inject_layouts(content)
        if theme:
            content = inject_theme(content, theme)
        # 以 .audited.html 结尾 → convert 跳过 work-copy 分支，不留 .audited 残留。
        # TemporaryDirectory：注入副本不再落在源目录（避免污染用户目录/并发互踩），
        # finally 随整目录清理。
        tmp_html_dir = tempfile.TemporaryDirectory(prefix="offipy-deck-")
        tmp_html = os.path.join(tmp_html_dir.name, f"{Path(html).stem}.audited.html")
        with open(tmp_html, "w", encoding="utf-8") as f:
            f.write(content)
        target = tmp_html
    tmp_pptx = None
    try:
        # 原子替换：mkstemp 生成与最终输出同目录的临时 .pptx（同卷、随机名）。
        # 占位文件删掉，让 convert.py 以正常权限全新创建（mkstemp 默认 0600
        # 不应泄漏给最终产物）；若转换失败没产出，finally 的清理是 no-op。
        fd, tmp_pptx = tempfile.mkstemp(
            prefix=f".{Path(final_out).stem}.",
            suffix=".pptx",
            dir=os.path.dirname(final_out) or ".",
        )
        os.close(fd)
        os.unlink(tmp_pptx)
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
        # 异常统一包装（_postprocess）：解析/数据错 → InvalidArgumentError，
        # 其余（measurements 缺失/XML/zip 损坏）→ ConversionError，均保留 __cause__。
        from .charts import postprocess_charts

        _postprocess("图表", postprocess_charts, html, tmp_pptx)
        # 图标后处理：HTML 声明了 data-icon → 替换成 freeform 矢量图标（同 charts 架构）。
        # 惰性 import：icons.py 内部 import python-pptx，不拖慢无图标的路径。
        from .icons import postprocess_icons

        _postprocess("图标", postprocess_icons, html, tmp_pptx)
        yield tmp_pptx, final_out
    finally:
        # 任何路径（成功或失败）都清理临时文件；已存在的 final_out 不受影响。
        # convert 的 <out>_audit 审计目录跟着临时 .pptx 名字走，tmp 被替换/删除后
        # 就成了孤儿（charts 后处理在 os.replace 前已读完 measurements），一并清掉。
        if tmp_html_dir is not None:
            tmp_html_dir.cleanup()  # 注入副本随整目录删除
        if tmp_pptx and os.path.exists(tmp_pptx):
            os.unlink(tmp_pptx)
        if tmp_pptx:
            tmp_audit = os.path.join(os.path.dirname(tmp_pptx), f"{Path(tmp_pptx).stem}_audit")
            if os.path.isdir(tmp_audit):
                shutil.rmtree(tmp_audit, ignore_errors=True)


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
    with _render_tmp(
        html, out, only_slides, no_visual_audit, timeout, theme, apply_layouts, overwrite
    ) as (tmp_pptx, final_out):
        os.replace(tmp_pptx, final_out)
    return final_out


@dataclass
class RenderResult:
    """render_with_report 的产出：生成的 .pptx 路径 + 完整审计报告。"""

    output_path: str
    audit_report: PptxAuditReport

    def to_dict(self) -> dict:
        return {"output_path": self.output_path, "audit": self.audit_report.to_dict()}


class AuditGateError(ConversionError):
    """Deck strict 门禁未通过：审计报告保留在异常上，供调用方落盘/展示。

    ConversionError 子类 → CLI 侧既有错误处理路径（exit 1）原样生效。
    """

    def __init__(self, message: str, report: PptxAuditReport, fail_on: Severity):
        super().__init__(message)
        self.report = report
        self.fail_on = fail_on


def render_with_report(
    html: str,
    out: str | None = None,
    only_slides: list[int] | None = None,
    no_visual_audit: bool = False,
    timeout: int = 600,
    theme: str | None = None,
    apply_layouts: bool = False,
    overwrite: bool = False,
    audit_mode: Literal["report", "strict"] = "report",
    fail_on: Severity = Severity.HIGH,
    audit_config: AuditConfig | None = None,
) -> RenderResult:
    """render + 静态质量审计，按 audit_mode 决定放行策略。

    report（默认）：生成 → 审计 → 替换 → 返回 RenderResult（产出路径 + 报告）；
    strict：生成 → 审计 → 最高严重度 ≥ fail_on → 抛 AuditGateError（报告在异常上、
    临时文件已清理、旧目标不动）；未达门槛 → 替换 → 返回 RenderResult。

    audit_mode 之外显式重复 render() 的主要参数，不用无限制 **render_kw
    （保 IDE 补全与文档）。
    """
    from .audit import audit_pptx

    with _render_tmp(
        html, out, only_slides, no_visual_audit, timeout, theme, apply_layouts, overwrite
    ) as (tmp_pptx, final_out):
        audit_report = audit_pptx(tmp_pptx, audit_config)
        gate = audit_report.max_severity
        if audit_mode == "strict" and gate is not None and gate >= fail_on:
            raise AuditGateError(
                f"审计门槛未通过：最高严重度 {gate.name} ≥ fail_on={fail_on.name}，"
                f"{final_out} 未替换（旧文件保留）",
                report=audit_report,
                fail_on=fail_on,
            )
        os.replace(tmp_pptx, final_out)
    return RenderResult(output_path=final_out, audit_report=audit_report)


def open_live(pptx: str) -> str:
    """在真实 PowerPoint 里打开生成的 .pptx（实况展示），返回 doc_id 供后续绑定。"""
    ensure_server()
    return call("ppt", "open_pres", path=os.path.abspath(pptx))


def export_slides(
    out_dir: str,
    width: int = 1920,
    height: int = 1080,
    doc_id: str | None = None,
    overwrite: bool = False,
) -> list[str]:
    """把 doc_id 指定（缺省活动）的演示文稿逐页导出 PNG，供 Claude 视觉迭代。

    overwrite=False（默认）拒绝覆盖已有输出；True 时原子替换（不残留半成品）。
    """
    ensure_server()
    return call(
        "ppt",
        "export_slides",
        out_dir=os.path.abspath(out_dir),
        width=width,
        height=height,
        doc_id=doc_id,
        overwrite=overwrite,
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
        # 导出必须绑定本次渲染的 deck（P0-2）：open_live 返回其 doc_id，逐页导出
        # 显式传给它，绝不依赖「当前活动焦点」——防用户中途切到别的文稿。
        doc_id = open_live(pptx)
        export_slides(feedback_dir, doc_id=doc_id, overwrite=overwrite)
    elif open_live_flag:
        open_live(pptx)
    return pptx
