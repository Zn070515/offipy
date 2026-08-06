"""HTML → 可编辑 PPTX 落地管线。

Claude 写 16:9 HTML 幻灯片 → vendored 的 vector-first 转换器（Playwright
实测 DOM 坐标）转成原生可编辑 .pptx → 通过常驻 server 在真实 PowerPoint 里
打开实况展示（页面在动），并逐页导出 PNG 供 Claude 视觉迭代。

设计系统：render 支持 `theme=` 注入内置主题（见 design.py）。Claude 在 HTML
的 <head> 写 `<style data-theme="<name>"></style>` 占位，render 时替换成主题
CSS；同一份 HTML 换主题即换皮，内容与视觉解耦。
"""

from __future__ import annotations

import contextlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Literal

from .audit import AuditConfig, PptxAuditReport, Severity
from .client import call, ensure_server
from .design import inject_theme
from .exceptions import ConversionError, FileConflictError, InvalidArgumentError
from .layouts import inject_layouts
from .paths import converter_data_dir

if TYPE_CHECKING:
    # 注解用到的 art 类型（惰性：运行时已被 from __future__ import annotations
    # 字符串化，绝不触发包加载即拉 python-pptx 链）。
    from .art.models import ArtReport, DeckQualityReport

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


def _preserve_audit_dir(tmp_pptx: str, final_out: str) -> None:
    """把 convert 的 <tmp_stem>_audit 测量目录改名为 <final_stem>_audit。

    aesthetic.audit(html) 按 html stem 自动发现 measurements.json，feedback 回路
    （feedback.append → dimension_weights）依赖它；tmp 名是 mkstemp 随机名，必须
    换成最终输出名才能让 `deck render` 的产物被审计到。
    """
    tmp_audit = os.path.join(os.path.dirname(tmp_pptx), f"{Path(tmp_pptx).stem}_audit")
    if not os.path.isdir(tmp_audit):
        return
    final_audit = os.path.join(os.path.dirname(final_out), f"{Path(final_out).stem}_audit")
    if os.path.isdir(final_audit):
        shutil.rmtree(final_audit, ignore_errors=True)
    try:
        os.replace(tmp_audit, final_audit)
    except OSError:
        shutil.rmtree(tmp_audit, ignore_errors=True)  # 改名失败不残留孤儿


def _atomic_replace(tmp: str, final: str) -> None:
    """原子替换临时 .pptx 到最终路径，并给可操作错误（#22）。

    Windows 下目标文件被占用（PowerPoint 打开 / 杀软 / 资源管理器）时 os.replace
    抛 PermissionError [WinError 5]。裸异常对迭代工作流不友好——把最常见的
    「PowerPoint 实况演示锁住产物」翻译成可执行指引。
    """
    try:
        os.replace(tmp, final)
    except PermissionError as e:
        if os.name == "nt":
            raise ConversionError(
                f"无法替换 {final}：目标文件被占用（WinError 5）。最常见是 PowerPoint "
                f"实况演示仍打开着它——请先 deck.close_live() 或 offipy quit ppt 关闭，"
                f"或改用新的输出路径。"
            ) from e
        raise


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
    defer_audit_preserve: bool = False,
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
        # 保留 convert 的 <stem>_audit 测量目录（aesthetic/feedback 回路）：
        # 默认立即改到最终名（render / render_with_report 行为不变）；
        # defer_audit_preserve=True（render_with_quality_report）保持 tmp 名，
        # 由 RenderStage.commit() 双产物事务一起移动。
        if not defer_audit_preserve:
            _preserve_audit_dir(tmp_pptx, final_out)
        yield tmp_pptx, final_out
    finally:
        # 任何路径（成功或失败）都清理临时文件；已存在的 final_out 不受影响。
        # 审计目录在 yield 前已从 tmp 名改到最终输出名（保留给 aesthetic/feedback），
        # 这里只兜底清理改名失败残留的孤儿 tmp 目录。
        if tmp_html_dir is not None:
            tmp_html_dir.cleanup()  # 注入副本随整目录删除
        if tmp_pptx and os.path.exists(tmp_pptx):
            os.unlink(tmp_pptx)
        if tmp_pptx:
            tmp_audit = os.path.join(os.path.dirname(tmp_pptx), f"{Path(tmp_pptx).stem}_audit")
            if os.path.isdir(tmp_audit):
                shutil.rmtree(tmp_audit, ignore_errors=True)


@contextmanager
def _render_stage(
    html,
    out=None,
    only_slides=None,
    no_visual_audit=False,
    timeout=600,
    theme=None,
    apply_layouts=False,
    overwrite=False,
):
    """渲染 + 双产物原子发布阶段（defer audit preserve）。

    审计目录在 with 块内保持 tmp 名（defer_audit_preserve=True），commit() 由
    render_with_quality_report 在成功路径显式调用（context 内、finally unlink 之前）。
    这里不自动 commit，避免双产物二次提交。
    """
    with _render_tmp(
        html,
        out,
        only_slides,
        no_visual_audit,
        timeout,
        theme,
        apply_layouts,
        overwrite,
        defer_audit_preserve=True,
    ) as (tmp_pptx, final_out):
        stage = RenderStage(tmp_pptx=tmp_pptx, final_pptx=final_out)
        try:
            yield stage
        except BaseException:
            stage.rollback()
            raise


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
        _atomic_replace(tmp_pptx, final_out)
    return final_out


@dataclass
class RenderResult:
    """render_with_report 的产出：生成的 .pptx 路径 + 完整审计报告。"""

    output_path: str
    audit_report: PptxAuditReport

    def to_dict(self) -> dict:
        return {"output_path": self.output_path, "audit": self.audit_report.to_dict()}


@dataclass
class QualityRenderResult(RenderResult):
    """render_with_quality_report 产出：几何审计 + 艺术分析组合。"""

    art_report: ArtReport | None = None
    deck_quality: DeckQualityReport | None = None


@dataclass
class RenderStage:
    """双产物原子发布事务：PPTX + 审计目录一起提交。

    - tmp_pptx：渲染产物（_render_tmp 的 finally 会在 context 退出时 unlink）
    - final_pptx：最终 PPTX 路径
    - committed：commit 是否已执行（避免二次 replace）

    rev2.1.1：审计目录不再由 _render_tmp 提前改名到最终位置（那会让失败时
    新审计目录已落位、旧 PPTX 保留，双产物不一致）。改为 defer_audit_preserve：
    with 块内审计目录保持 tmp 名，commit() 才先换 PPTX 再移动审计目录。
    """

    tmp_pptx: str
    final_pptx: str
    committed: bool = False

    @property
    def tmp_audit_dir(self) -> Path:
        return Path(self.tmp_pptx).parent / f"{Path(self.tmp_pptx).stem}_audit"

    @property
    def final_audit_dir(self) -> Path:
        return Path(self.final_pptx).parent / f"{Path(self.final_pptx).stem}_audit"

    @property
    def measurements_path(self) -> Path | None:
        """with 块内审计目录仍在 tmp 名：从这里读 measurements（最终名还没生成）。"""
        m = self.tmp_audit_dir / "_cache" / "measurements.json"
        return m if m.is_file() else None

    def commit(self) -> None:
        """先原子替换 PPTX，成功后把 tmp 审计目录改到最终名（双产物一起落位）。"""
        if self.committed:
            return
        _atomic_replace(self.tmp_pptx, self.final_pptx)
        if self.tmp_audit_dir.is_dir():
            _preserve_audit_dir(self.tmp_pptx, self.final_pptx)
        self.committed = True

    def rollback(self) -> None:
        """失败回滚：不提交；tmp_pptx 与 tmp 审计目录由 _render_tmp finally 清理。"""
        self.committed = False


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

    非法 audit_mode 拼写会在进入渲染前抛 InvalidArgumentError（fail-open 防护：
    拼错不会静默退化成 report 绕过 strict 门禁）；fail_on 必须是 Severity。
    """
    from .audit import audit_pptx

    if audit_mode not in {"report", "strict"}:
        raise InvalidArgumentError("audit_mode must be 'report' or 'strict'")
    if not isinstance(fail_on, Severity):
        raise InvalidArgumentError("fail_on must be a Severity")

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
        _atomic_replace(tmp_pptx, final_out)
    return RenderResult(output_path=final_out, audit_report=audit_report)


def _run_art_analysis(
    measurements: dict | str,
    profile: str,
    pptx_report: object | None = None,
    slides_dir: str | None = None,
) -> ArtReport:
    """从保留的 measurements +（可选）几何审计 +（可选）像素 slides_dir 建场景并分析。"""
    from .art import analyze_scene, build_scene

    scene = build_scene(measurements=measurements, pptx_report=pptx_report, slides_dir=slides_dir)
    return analyze_scene(scene, profile=profile)


def _check_art_gate() -> None:
    """v0.12 占位：艺术层默认不阻断。strict 门禁仍归几何层。"""
    return None


def render_with_quality_report(
    html: str,
    out: str | None = None,
    only_slides: list[int] | None = None,
    no_visual_audit: bool = False,
    timeout: int = 600,
    theme: str | None = None,
    apply_layouts: bool = False,
    overwrite: bool = False,
    audit_mode: str = "report",
    fail_on: Severity = Severity.HIGH,
    audit_config: AuditConfig | None = None,
    profile: str = "balanced",
    pixel_analysis: Literal["off", "best_effort", "required"] = "off",
    preserve_pixel_slides: bool = False,
    slides_output_dir: str | None = None,
) -> QualityRenderResult:
    """render + 几何审计 + 艺术分析（原子发布）。

    与 render_with_report 契约一致（audit_mode / fail_on / audit_config 同语义）；
    艺术结果默认只建议不阻断。art 惰性 import。
    """
    from .art import ArtWarning, DeckQualityReport
    from .audit import audit_pptx

    if audit_mode not in {"report", "strict"}:
        raise InvalidArgumentError("audit_mode must be 'report' or 'strict'")
    if not isinstance(fail_on, Severity):
        raise InvalidArgumentError("fail_on must be a Severity")
    if pixel_analysis not in {"off", "best_effort", "required"}:
        raise InvalidArgumentError("pixel_analysis must be 'off', 'best_effort', or 'required'")

    with _render_stage(
        html, out, only_slides, no_visual_audit, timeout, theme, apply_layouts, overwrite
    ) as stage:
        audit_report = audit_pptx(stage.tmp_pptx, audit_config)
        gate = audit_report.max_severity
        if audit_mode == "strict" and gate is not None and gate >= fail_on:
            raise AuditGateError(
                f"审计门槛未通过：最高严重度 {gate.name} ≥ fail_on={fail_on.name}，"
                f"{stage.final_pptx} 未替换（旧文件保留）",
                report=audit_report,
                fail_on=fail_on,
            )
        m = stage.measurements_path
        art_report: ArtReport | None = None
        warnings: list[ArtWarning] = []
        staging_slides: str | None = None
        staging_dir: str | None = None
        if m is not None and pixel_analysis != "off":
            staging_dir = tempfile.mkdtemp(
                prefix="offipy-pixel-", dir=os.path.dirname(stage.final_pptx) or "."
            )
            staging_slides = os.path.join(staging_dir, "slides")
            os.makedirs(staging_slides, exist_ok=True)
            try:
                _export_pixel_slides(stage.tmp_pptx, staging_slides)
                _write_deck_info(staging_slides, stage.tmp_pptx)
            except Exception as exc:
                shutil.rmtree(staging_dir, ignore_errors=True)
                if pixel_analysis == "required":
                    raise ConversionError(f"像素分析导出失败（required）: {exc}") from exc
                warnings.append(
                    ArtWarning(
                        code="art.pixel.best_effort_failed",
                        message=f"像素分析导出失败，已跳过: {exc}",
                    )
                )
                staging_slides = None
                staging_dir = None
        if m is not None:
            # 双源融合：measurements 为主 + 刚审计的 pptx_report 作 secondary
            # +（可选）像素 slides_dir 作 tertiary
            art_report = _run_art_analysis(
                m, profile, pptx_report=audit_report, slides_dir=staging_slides
            )
        else:
            warnings.append(
                ArtWarning(
                    code="art.measurements_missing",
                    message="未找到 measurements.json，跳过艺术分析",
                )
            )
        _check_art_gate()
        # commit() 在这里的 context 内调用：_render_tmp finally unlink tmp_pptx 之前，
        # 同时把 tmp 审计目录改到最终名（双产物一起落位）
        stage.commit()
        if preserve_pixel_slides and staging_slides is not None:
            _move_slides_to_final(staging_slides, stage, slides_output_dir)
        elif staging_dir is not None:
            shutil.rmtree(staging_dir, ignore_errors=True)
    deck_quality = DeckQualityReport(geometry=audit_report, art=art_report, warnings=warnings)
    return QualityRenderResult(
        output_path=stage.final_pptx,
        audit_report=audit_report,
        art_report=art_report,
        deck_quality=deck_quality,
    )


def _export_pixel_slides(pptx: str, out_dir: str) -> list[str]:
    """在真实 PowerPoint 中打开 tmp_pptx 并逐页导出 PNG（PowerPoint 锁的是副本）。"""
    doc_id = open_live(pptx)
    try:
        return export_slides(out_dir, doc_id=doc_id, overwrite=True)
    finally:
        close_live(doc_id)


def _write_deck_info(out_dir: str, pptx: str) -> None:
    info = {"schema": 1, "pptx_sha256": _sha256_file(pptx), "run_id": None}
    Path(out_dir, "_deck_info.json").write_text(
        json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _sha256_file(path: str) -> str:
    import hashlib

    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _default_slides_dir(final_pptx: str) -> str:
    return str(Path(final_pptx).with_suffix("")) + "_slides"


def _move_slides_to_final(
    staging_slides: str, stage: RenderStage, slides_output_dir: str | None
) -> str:
    final_slides = slides_output_dir or _default_slides_dir(stage.final_pptx)
    os.makedirs(final_slides, exist_ok=True)
    for f in Path(staging_slides).iterdir():
        if f.is_file():
            shutil.copy2(str(f), os.path.join(final_slides, f.name))
    # 拷贝完成后清理 staging 目录（slides 的父目录即 mkdtemp 的 staging 根）
    shutil.rmtree(os.path.dirname(staging_slides), ignore_errors=True)
    return final_slides


# #22：open_live 前把 .pptx 复制到系统临时目录的 offipy-live-* 副本再让 PowerPoint
# 打开——PowerPoint 锁定的是副本，源产物路径永不被锁，同路径 re-render(overwrite=True)
# 不再 PermissionError。副本由 close_live 删除，废弃残留由 _cleanup_stale_live_tmp 兜底。
_LIVE_TMP_PREFIX = "offipy-live-"
# doc_id → 临时副本路径（open_live 登记，close_live 清理）。会话级尽力而为：
# server 重启 / 直接 ppt.close_pres 关闭时靠 _cleanup_stale_live_tmp 兜底。
_LIVE_TMP_PATHS: dict[str, str] = {}


def _cleanup_stale_live_tmp() -> None:
    """清理废弃的 offipy-live-* 副本（关闭/崩溃/重启后残留）。"""
    import time

    stale_before = time.time() - 3600  # 1 小时未动即视为废弃
    tmp_dir = Path(tempfile.gettempdir())
    for f in tmp_dir.glob(f"{_LIVE_TMP_PREFIX}*.pptx"):
        try:
            if f.stat().st_mtime < stale_before:
                f.unlink()
        except OSError:
            pass  # 仍被 PowerPoint 打开（无共享删除）时跳过


def _live_tmp_copy(pptx: str) -> str:
    """把 .pptx 复制成 offipy-live-* 临时副本，返回副本路径。"""
    src = os.path.abspath(pptx)
    if not os.path.exists(src):
        raise InvalidArgumentError(f"源 .pptx 不存在: {src}")
    _cleanup_stale_live_tmp()
    fd, tmp = tempfile.mkstemp(prefix=_LIVE_TMP_PREFIX, suffix=".pptx")
    os.close(fd)
    os.unlink(tmp)  # 先删占位，copyfile 以正常权限全新创建
    shutil.copyfile(src, tmp)
    return tmp


def open_live(pptx: str) -> str:
    """在真实 PowerPoint 里打开生成的 .pptx（实况展示），返回 doc_id 供后续绑定。

    #22：打开前复制到系统临时目录的 offipy-live-* 副本——PowerPoint 锁定副本而非
    产物，同路径 re-render(overwrite=True) 不再因源文件被锁而 PermissionError。
    关闭实况演示：deck.close_live(doc_id)；或 offipy quit ppt 整会话退出。
    """
    ensure_server()
    live = _live_tmp_copy(pptx)
    try:
        doc_id = call("ppt", "open_pres", path=live)
    except Exception:
        with contextlib.suppress(OSError):
            os.unlink(live)  # 打开失败不留孤儿副本
        raise
    _LIVE_TMP_PATHS[doc_id] = live
    return doc_id


def close_live(doc_id: str) -> None:
    """关闭 open_live 打开的实况演示，释放其占用的临时副本（#22）。

    save=False 直接关闭不保存——实况展示操作的是临时副本，回写无意义。
    配合 #26 的 Ppt.close_pres：关闭后同路径 re-render 不再 PermissionError。
    """
    ensure_server()
    call("ppt", "close_pres", doc_id=doc_id, save=False)
    path = _LIVE_TMP_PATHS.pop(doc_id, None)
    if path:
        with contextlib.suppress(OSError):
            os.unlink(path)


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
