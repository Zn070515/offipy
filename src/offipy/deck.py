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
import re
import shutil
import signal
import subprocess
import sys
import tempfile
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace
from typing import TYPE_CHECKING, Any, Literal, cast
from urllib.parse import unquote

from .assets.declarations import preprocess_asset_declarations
from .audit import AuditConfig, PptxAuditReport, Severity
from .client import call, ensure_server
from .design import inject_theme
from .exceptions import ConversionError, FileConflictError, InvalidArgumentError
from .layouts import chart_dominant_slide_indices, inject_layouts
from .paths import converter_data_dir

if TYPE_CHECKING:
    # 注解用到的 art 类型（惰性：运行时已被 from __future__ import annotations
    # 字符串化，绝不触发包加载即拉 python-pptx 链）。
    from collections.abc import Callable, Iterator

    from .art.models import ArtReport, DeckQualityReport

# vendored 转换器位于包内 _vendor/，site-packages 下经 __file__ 自定位
_CONVERT_DIR = Path(__file__).resolve().parent / "_vendor" / "html_to_editable_pptx"
CONVERT_PY = _CONVERT_DIR / "convert.py"


def _preflight_chart_layout(
    html_path: str, apply_layouts: bool, only_slides: list[int] | None
) -> None:
    """前置检查：chart-dominant 布局但未启用布局注入 → fail-fast。

    chart-dominant 的 .chart 可测尺寸由布局 CSS（flex/min-height）提供；
    apply_layouts=False（CLI 缺 --layouts）时 CSS 未注入，post-render 测量必然
    丢框 → 在启动 chromium / 跑 convert 之前就把用户拦下来，给可操作指引。
    only_slides 只拦指定页（对齐 convert 的 only-slides 语义）。
    """
    if apply_layouts:
        return
    with Path(html_path).open(encoding="utf-8") as f:
        content = f.read()
    indices = chart_dominant_slide_indices(content)
    if only_slides:
        indices = [i for i in indices if i in only_slides]
    if not indices:
        return
    raise InvalidArgumentError(
        f'第 {indices} 页声明了 data-layout="chart-dominant" 但未启用布局注入'
        "（apply_layouts=False）：chart-dominant 的图表尺寸由布局 CSS 提供，未注入时"
        "无法测量。请给 render/make 传 apply_layouts=True（CLI 用 --layouts），或改用"
        "自定义布局并为 .chart 提供显式可测尺寸。"
    )


def _preflight_browser() -> None:
    """渲染前置：确保 chromium 可用，否则给安装提示（P0-2）。"""
    from .envcheck import _check_browser

    check = _check_browser()
    if not check.ok:
        raise ConversionError(
            "HTML→PPTX 渲染需要 Chromium："
            f"{check.detail}。请运行: python -m playwright install chromium"
        )


_NOVA_DECLARATION_MARKERS = (
    ("data-chart", "图表(data-chart)"),
    ("data-icon", "图标(data-icon)"),
    ("data-asset", "资源(data-asset)"),
    ("data-primitive", "图元(data-primitive)"),
    ('class="mermaid"', "图示(mermaid)"),
    ("class='mermaid'", "图示(mermaid)"),  # 单引号变体，防 no_visual_audit fail-fast 漏检
    ('class="drawio"', "图示(drawio)"),
    ("class='drawio'", "图示(drawio)"),  # 单引号变体，同上
)


def _reject_no_visual_audit_declarations(content: str) -> None:
    """no_visual_audit 与声明注入不兼容 → fail-fast（chromium / convert 之前）。

    图表/图标/资源/图元的注入都依赖 visual audit 的 measurements.json；no_visual_audit
    不产出 → 命中任一声明类型就报错，信息列出具体类型。
    """
    found = [label for marker, label in _NOVA_DECLARATION_MARKERS if marker in content]
    if found:
        raise InvalidArgumentError(
            "no_visual_audit=True 但 HTML 声明了"
            + "、".join(found)
            + "：注入需要 measurements.json，请去掉 no_visual_audit 或用 visual audit 渲染"
        )


# data-* 自定义属性（data-icon/data-asset/data-chart-data 等）是逻辑值不是 URL；
# `data-chart-data=` 会从第二个 data 处误命中 data=。前置 (?<![\w-]) 只让真实
# 属性名（前导空格/< /"等）命中，属性名内含 - 或字母的自定义属性不重写。
# 白名单含 URL 承载属性（src/href/poster/data/background/cite/action/formaction/
# longdesc，#77）：漏一个属性就多一个注入副本下解析失败破图的点。
_ATTR_URL_RE = re.compile(
    r"((?<![\w-])(?:src|href|poster|data-drawio|data|background|cite|action|"
    r"formaction|longdesc)\s*=\s*[\"'])([^\"']*)([\"'])"
)
# srcset 值里可能带引号（data: URI 的 SVG 常用单引号属性），只以属性自身的结束
# 引号定界，不能用 [^\"']*（会把含引号的 data URI 截断）。
_SRCSET_RE = re.compile(r'(srcset\s*=\s*(["\']))((?:(?!\2).)*)(\2)')
_IMPORT_URL_RE = re.compile(r"(@import\s+[\"'])([^\"']*)([\"'])")
_CSS_URL_RE = re.compile(r"(url\(\s*[\"']?)([^\"')]*)([\"']?\s*\))")


def _abs_url(base_dir: Path, url: str) -> str:
    url = url.strip()
    if not url:
        return url
    if url.startswith(("//", "http://", "https://", "data:", "mailto:", "tel:", "file:")):
        return url
    frag = ""
    if "#" in url:
        url, frag = url.split("#", 1)
    if not url:
        return f"#{frag}"
    # unquote 再 resolve：URL 里的 %20 等编码对应文件系统里的真实名字，as_uri
    # 会按绝对路径重新编码，避免生成 file:///.../My%2520Photo.png 这类双编码。
    resolved = (base_dir / unquote(url)).resolve().as_uri()
    return resolved + (f"#{frag}" if frag else "")


_DATA_URI_RE = re.compile(r"data:[^\s]+")


def _rewrite_srcset(base_dir: Path, value: str) -> str:
    """重写 srcset 里的每个候选 URL，保留 1x/2x/300w 描述符（#63）。

    data: URI 不含空白但可能含逗号（如 `data:image/svg+xml;utf8,<svg>...`）——
    先整体藏起来再按逗号拆候选，否则会被误切成多段截断（#audit）。
    """
    stashed: list[str] = []

    def _stash(m: re.Match[str]) -> str:
        stashed.append(m.group(0))
        return f"__OFFIPY_DATA_{len(stashed) - 1}__"

    def _restore(token: str) -> str:
        return re.sub(
            r"__OFFIPY_DATA_(\d+)__",
            lambda m: stashed[int(m.group(1))],
            token,
        )

    out = []
    for candidate in _DATA_URI_RE.sub(_stash, value).split(","):
        parts = candidate.strip().split()
        if not parts:
            continue
        parts = [_restore(p) for p in parts]
        out.append(" ".join([_abs_url(base_dir, parts[0]), *parts[1:]]).rstrip())
    return ", ".join(out)


def _rewrite_relative_urls(content: str, base_dir: Path) -> str:
    """把相对 src/href/poster/data/srcset/CSS url()/@import 重写为绝对 file:// URL。

    注入副本落在 offipy-deck-* 临时目录，源 HTML 基于自身目录的相对引用
    （img/、样式、url()）在副本下会解析失败 → 资源加载不到（#57）。写入副本前
    把相对 URL 全转绝对 file://；scheme/协议相对/纯 fragment 原样保留。
    """
    content = _ATTR_URL_RE.sub(
        lambda m: m.group(1) + _abs_url(base_dir, m.group(2)) + m.group(3), content
    )
    content = _SRCSET_RE.sub(
        lambda m: m.group(1) + _rewrite_srcset(base_dir, m.group(3)) + m.group(4),
        content,
    )
    content = _CSS_URL_RE.sub(
        lambda m: m.group(1) + _abs_url(base_dir, m.group(2)) + m.group(3), content
    )
    return _IMPORT_URL_RE.sub(
        lambda m: m.group(1) + _abs_url(base_dir, m.group(2)) + m.group(3), content
    )


def _prepare_target(
    html: str, content: str, theme: str | None, apply_layouts: bool
) -> tuple[str, tempfile.TemporaryDirectory[str] | None]:
    """把源 HTML 整理成转换输入：layout → theme → asset 声明，写注入副本或原样返回。

    顺序固定：先注入布局/主题再解析 asset 声明，converter 才能量到最终 CSS token
    下的资源尺寸。任一变换发生就写临时注入副本（.audited.html 命名 → convert 跳过
    work-copy 分支），TemporaryDirectory 保证副本不落源目录、随 finally 清理；
    全无变换则返回源路径，不产临时文件。
    """
    if apply_layouts:
        content = inject_layouts(content)
    if theme:
        content = inject_theme(content, theme)
    content, asset_decls = preprocess_asset_declarations(content)
    if not (apply_layouts or theme is not None or asset_decls):
        return html, None
    content = _rewrite_relative_urls(content, Path(html).resolve().parent)
    tmp_html_dir = tempfile.TemporaryDirectory(prefix="offipy-deck-")
    tmp_html = str(Path(tmp_html_dir.name) / f"{Path(html).stem}.audited.html")
    with Path(tmp_html).open("w", encoding="utf-8") as f:
        f.write(content)
    return tmp_html, tmp_html_dir


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
    # --no-work-copy：deck 自己管理输入（源 html 或临时注入副本），转换器不得在
    # 源目录创建 .audited.html 工作副本（污染用户目录）。
    cmd += ["--no-work-copy"]
    # --fail-on-selfcheck：自检（Stage 5a 结构/像素校验）异常不再静默吞掉——deck
    # 管线里验证失败必须显式报错，否则 audit 报告宣称"已校验"而底层根本没跑。
    cmd += ["--fail-on-selfcheck"]
    return cmd


def _kill_process_tree(pid: int) -> None:
    """整树杀进程。Windows 用 taskkill /T——根进程还活着时才能顺带杀孙进程。

    不能等根进程死了再枚举子孙：Windows 上父进程死后子进程会被 reparent，
    按 PPID 事后枚举不到。POSIX 用进程组。
    """
    if os.name == "nt":
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True, text=True)
    else:
        # POSIX 分支在 CI（ubuntu 纯模块测试）可达。convert 子进程以独立会话启动
        # （见 _run_convert 的 start_new_session），是组长（pgid==pid），killpg(pid)
        # 只杀该组，绝不波及调用方进程组（否则会连 pytest/CI runner 一起 SIGKILL）。
        # mypy 在 Windows 平台类型上会把 killpg/SIGKILL 判为不存在，显式忽略 attr 检查。
        with contextlib.suppress(ProcessLookupError):
            os.killpg(pid, signal.SIGKILL)  # type: ignore[attr-defined]


def _run_convert(cmd: list[str], timeout: int, env: dict[str, str]) -> SimpleNamespace:
    """跑 convert 子进程，返回 returncode/stdout/stderr（SimpleNamespace）。

    超时用 Popen.communicate 手动处理而非 subprocess.run(timeout=)：run 超时只杀
    直接子进程，convert.py 派生的 chromium/渲染器孙进程会泄漏。这里超时先趁根
    还活着 taskkill /T 整树杀掉，再收割输出（消息与旧 run 分支格式一致）。
    """
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        # POSIX 下让 convert 子进程自成会话/进程组，超时整树杀才只杀该组；
        # 否则继承调用方进程组，killpg 会把 pytest/CI runner 一起杀掉。
        start_new_session=os.name != "nt",
    )
    try:
        out, err = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        _kill_process_tree(proc.pid)
        out, err = proc.communicate()
        raise ConversionError(f"convert.py 超时 ({timeout}s)\n{out}\n{err}") from None
    return SimpleNamespace(returncode=proc.returncode, stdout=out, stderr=err)


def _default_out(html: str) -> str:
    """对齐 convert.py 的默认命名：foo.audited.html → foo.pptx，否则 foo.pptx。"""
    p = Path(html)
    if p.name.endswith(".audited.html"):
        return str(p.with_name(p.name[: -len(".audited.html")] + ".pptx"))
    return str(p.with_suffix(".pptx"))


def _postprocess(label: str, fn: Callable[[str, str], Any], html: str, pptx: str) -> Any:
    """统一包装图表/资源后处理：解析/数据错 → InvalidArgumentError，其余 → ConversionError。

    ValueError（HTML 图表/资源声明非法、数据缺失/损坏）属用户输入问题；
    其余异常（measurements 缺失、python-pptx/XML/zip 损坏）属转换产物问题。
    均 from e 保留 __cause__。返回 fn 的返回值（postprocess_assets 的用量报告）。
    """
    try:
        return fn(html, pptx)
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
    tmp_audit = Path(tmp_pptx).parent / f"{Path(tmp_pptx).stem}_audit"
    if not tmp_audit.is_dir():
        return
    final_audit = Path(final_out).parent / f"{Path(final_out).stem}_audit"
    if final_audit.is_dir():
        shutil.rmtree(final_audit, ignore_errors=True)
    try:
        tmp_audit.replace(final_audit)
    except OSError:
        shutil.rmtree(tmp_audit, ignore_errors=True)  # 改名失败不残留孤儿


def _atomic_replace(tmp: str, final: str, *, overwrite: bool = True) -> None:
    """原子替换临时 .pptx 到最终路径，并给可操作错误（#22）。

    Windows 下目标文件被占用（PowerPoint 打开 / 杀软 / 资源管理器）时 os.replace
    抛 PermissionError [WinError 5]。裸异常对迭代工作流不友好——把最常见的
    「PowerPoint 实况演示锁住产物」翻译成可执行指引。

    overwrite=False 时落盘前再查一次 final 是否已存在：_render_tmp 前置检查与
    os.replace 之间有空窗（并发/外部进程可能已创建目标），复查把 TOCTOU 窗口压到
    最小。真正无竞争的原子不覆盖需要 O_EXCL 语义，Windows 上 os.replace 不提供，
    这里取「复查 + 拒绝」的最优近似。
    """
    if not overwrite and Path(final).exists():
        raise FileConflictError(f"输出 .pptx 已存在（overwrite=False）: {final}")
    try:
        Path(tmp).replace(final)
    except PermissionError as e:
        if os.name == "nt":
            # #90：目录目标也会 WinError 5——这是「out 是目录」参数错，不是 PowerPoint
            # 锁。先排除，避免误导 Agent 去做无意义的 close_live 排查。
            if Path(final).is_dir():
                raise InvalidArgumentError(
                    f"out 是一个目录: {final}（out 应为 .pptx 输出文件路径，不是目录）"
                ) from e
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
    html = str(Path(html).resolve())
    if not Path(html).exists():
        raise InvalidArgumentError(f"源 HTML 文件不存在: {html}")
    _preflight_chart_layout(html, apply_layouts, only_slides)
    with Path(html).open(encoding="utf-8") as f:
        content = f.read()
    if no_visual_audit:
        # 声明注入依赖 visual audit 的 measurements.json；no_visual_audit 不产出
        # → 启动 chromium / 跑 convert 之前 fail-fast，省一次白跑的渲染。
        _reject_no_visual_audit_declarations(content)
    _preflight_browser()
    final_out = str(Path(out).resolve()) if out else _default_out(html)
    # #90：out 前置校验——父目录不存在 / out 是目录时给可操作领域错误。否则 mkstemp
    # 在父目录缺失处抛裸 FileNotFoundError（泄漏临时文件名），os.replace 对目录目标
    # 抛 WinError 5 又会被误归因 PowerPoint 占用，两种都是低质量错误。
    parent = Path(final_out).parent or "."
    if not Path(parent).is_dir():
        raise InvalidArgumentError(f"输出目录不存在: {parent}（请先创建该目录再 render）")
    if Path(final_out).is_dir():
        raise InvalidArgumentError(
            f"out 是一个目录: {final_out}（out 应为 .pptx 输出文件路径，不是目录）"
        )
    if not overwrite and Path(final_out).exists():
        raise FileConflictError(
            f"输出 .pptx 已存在（overwrite=False，可传 overwrite=True 覆盖）: {final_out}"
        )
    target, tmp_html_dir = _prepare_target(html, content, theme, apply_layouts)
    tmp_pptx = None
    try:
        # 原子替换：mkstemp 生成与最终输出同目录的临时 .pptx（同卷、随机名）。
        # 占位文件删掉，让 convert.py 以正常权限全新创建（mkstemp 默认 0600
        # 不应泄漏给最终产物）；若转换失败没产出，finally 的清理是 no-op。
        fd, tmp_pptx = tempfile.mkstemp(
            prefix=f".{Path(final_out).stem}.",
            suffix=".pptx",
            dir=Path(final_out).parent or ".",
        )
        os.close(fd)
        Path(tmp_pptx).unlink()
        cmd = _convert_cmd(target, tmp_pptx, only_slides, no_visual_audit)
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"  # 中文 Windows 下 convert.py 输出才不会乱码
        # 转换器可变数据（配置/lessons-learned）落用户数据目录，不写包内
        env["OFFIPY_CONVERTER_DATA_DIR"] = str(converter_data_dir())
        r = _run_convert(cmd, timeout=timeout, env=env)
        if r.returncode != 0:
            raise ConversionError(f"convert.py 失败 (exit {r.returncode})\n{r.stdout}\n{r.stderr}")
        if not Path(tmp_pptx).exists():
            raise ConversionError(f"转换未产出 .pptx: {tmp_pptx}\n{r.stdout}\n{r.stderr}")
        # 图表后处理：HTML 声明了 data-chart → 读 measurements 替换成原生图表。
        # 惰性 import：charts.py 内部 import python-pptx，不拖慢无图表的路径。
        # 异常统一包装（_postprocess）：解析/数据错 → InvalidArgumentError，
        # 其余（measurements 缺失/XML/zip 损坏）→ ConversionError，均保留 __cause__。
        from .charts import postprocess_charts

        _postprocess("图表", postprocess_charts, target, tmp_pptx)
        # Mermaid 图后处理：<pre class="mermaid"> 块 → 可编辑形状（图在后，
        # 覆盖 charts 注入的同类 bbox 替换语义；先于 assets 避免资产形状干扰）。
        from .diagrams import postprocess_mermaid

        _postprocess("图示", postprocess_mermaid, target, tmp_pptx)
        # draw.io 图后处理：<div class="drawio" data-drawio="..."> → 可编辑形状
        # （图在后，覆盖 mermaid 注入的同类 bbox 替换语义；先于 assets 避免干扰）。
        from .drawio import postprocess_drawio

        _postprocess("图示(drawio)", postprocess_drawio, target, tmp_pptx)
        # 资源后处理：data-icon/data-asset/data-primitive → 统一 asset 管线（取代
        # postprocess_icons）。图表必须先于资源：图表用自己测量的占位符替换，不应受
        # 随后添加的 asset 形状影响。返回用量报告供 assets.json 清单。
        # 惰性 import：assets.render 内部才 import python-pptx，不拖慢无资产路径。
        from .assets.manifest import write_asset_manifest
        from .assets.render import postprocess_assets

        report = _postprocess("资源", postprocess_assets, target, tmp_pptx)
        # assets.json provenance 清单：只写进 visual-audit 的 <tmp>_audit 目录
        # （随 _preserve_audit_dir 整体改名进最终输出）；no_visual_audit 无审计
        # 目录 → 不写，资产已渲染进 PPTX、仅缺清单。
        tmp_audit_dir = Path(tmp_pptx).parent / f"{Path(tmp_pptx).stem}_audit"
        if tmp_audit_dir.is_dir():
            write_asset_manifest(str(tmp_audit_dir / "assets.json"), report)
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
        if tmp_pptx and Path(tmp_pptx).exists():
            Path(tmp_pptx).unlink()
        if tmp_pptx:
            tmp_audit = Path(tmp_pptx).parent / f"{Path(tmp_pptx).stem}_audit"
            if tmp_audit.is_dir():
                shutil.rmtree(tmp_audit, ignore_errors=True)


@contextmanager
def _render_stage(
    html: str,
    out: str | None = None,
    only_slides: list[int] | None = None,
    no_visual_audit: bool = False,
    timeout: int = 600,
    theme: str | None = None,
    apply_layouts: bool = False,
    overwrite: bool = False,
) -> Iterator[RenderStage]:
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
        stage = RenderStage(tmp_pptx=tmp_pptx, final_pptx=final_out, overwrite=overwrite)
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
    清理临时文件——已存在的 .pptx 绝不因一次失败的渲染被破坏。双产物
    （.pptx + 审计目录）由 RenderStage.commit 先换 pptx、后移审计目录
    （#audit-H1：审计目录绝不先于 pptx 落位，替换失败时新旧产物一致）。
    """
    with _render_stage(
        html, out, only_slides, no_visual_audit, timeout, theme, apply_layouts, overwrite
    ) as stage:
        stage.commit()
    return stage.final_pptx


@dataclass
class RenderResult:
    """render_with_report 的产出：生成的 .pptx 路径 + 完整审计报告。"""

    output_path: str
    audit_report: PptxAuditReport

    def to_dict(self) -> dict[str, Any]:
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
    overwrite: bool = True

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
        _atomic_replace(self.tmp_pptx, self.final_pptx, overwrite=self.overwrite)
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
        stage.commit()
    return RenderResult(output_path=stage.final_pptx, audit_report=audit_report)


def _run_art_analysis(
    measurements: dict[str, Any] | str,
    profile: str,
    pptx_report: object | None = None,
    slides_dir: str | None = None,
    pixel_required: bool = False,
    feedback_dir: str | None = None,
    include_experimental_score: bool = False,
) -> ArtReport:
    """从保留的 measurements +（可选）几何审计 +（可选）像素 slides_dir 建场景并分析。"""
    from .art import analyze_scene, build_scene

    scene = build_scene(measurements=measurements, pptx_report=pptx_report, slides_dir=slides_dir)
    if pixel_required and slides_dir is not None and "pixel" not in scene.sources:
        raise ConversionError("像素分析无有效页面（required）")
    return analyze_scene(
        scene,
        profile=profile,
        include_experimental_score=include_experimental_score,
        feedback=feedback_dir is not None,
        feedback_dir=feedback_dir,
    )


def _check_art_gate() -> None:
    """v0.12 占位：艺术层默认不阻断。strict 门禁仍归几何层。"""
    return


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
    feedback_dir: str | None = None,
    pixel_analysis: Literal["off", "best_effort", "required"] = "off",
    preserve_pixel_slides: bool = False,
    slides_output_dir: str | None = None,
    include_experimental_score: bool = False,
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
        try:
            if m is not None and pixel_analysis != "off":
                staging_dir = tempfile.mkdtemp(
                    prefix="offipy-pixel-", dir=Path(stage.final_pptx).parent or "."
                )
                staging_slides = str(Path(staging_dir) / "slides")
                try:
                    Path(staging_slides).mkdir(exist_ok=True, parents=True)
                    exported = _export_pixel_slides(stage.tmp_pptx, staging_slides)
                    _write_deck_info(staging_slides, stage.tmp_pptx)
                except Exception as exc:
                    if pixel_analysis == "required":
                        raise ConversionError(f"像素分析导出失败（required）: {exc}") from exc
                    warnings.append(
                        ArtWarning(
                            code="art.pixel.best_effort_failed",
                            message=f"像素分析导出失败，已跳过: {exc}",
                        )
                    )
                    staging_slides = None
                else:
                    if pixel_analysis == "required" and not exported:
                        raise ConversionError("像素分析未导出任何页面（required）")
            if m is not None:
                # 双源融合：measurements 为主 + pptx_report secondary + slides_dir tertiary
                art_report = _run_art_analysis(
                    str(m),
                    profile,
                    pptx_report=audit_report,
                    slides_dir=staging_slides,
                    pixel_required=(pixel_analysis == "required"),
                    feedback_dir=feedback_dir,
                    include_experimental_score=include_experimental_score,
                )
            else:
                if pixel_analysis == "required":
                    raise ConversionError("像素分析（required）需要 measurements.json")
                warnings.append(
                    ArtWarning(
                        code="art.measurements_missing",
                        message="未找到 measurements.json，跳过艺术分析",
                    )
                )
            _check_art_gate()
            # commit() 在 context 内调用：_render_tmp finally unlink tmp_pptx 之前，
            # 同时把 tmp 审计目录改到最终名（双产物一起落位）
            stage.commit()
            if preserve_pixel_slides and staging_slides is not None:
                _move_slides_to_final(staging_slides, stage, slides_output_dir)
        finally:
            if staging_dir is not None:
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
    info = {
        "schema": 1,
        "pptx": str(Path(pptx).resolve()),
        "pptx_sha256": _sha256_file(pptx),
        "run_id": None,
    }
    Path(out_dir, "_deck_info.json").write_text(
        json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _sha256_file(path: str) -> str:
    import hashlib

    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _default_slides_dir(final_pptx: str) -> str:
    return str(Path(final_pptx).with_suffix("")) + "_slides"


def _is_slide_png(name: str) -> bool:
    return name.startswith("slide_") and name.endswith(".png")


def _slides_dir_owned(final_slides: str, final_pptx: str) -> bool:
    """该目录的 slide_*.png 是否归本 deck？凭 _deck_info.json 的 pptx 路径判断。

    首次渲染目录里没有标记 → False（没有本 deck 旧产物可清，只做复制）；
    上次标记的 pptx 路径与当前一致 → 是 re-render，可安全清理旧页；
    标记指向别的 pptx / 标记缺失 / 标记损坏 → 目录是别人的，绝不删任何文件。
    """
    info_path = Path(final_slides, "_deck_info.json")
    if not info_path.is_file():
        return False
    try:
        info = json.loads(info_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False
    return str(info.get("pptx")) == str(Path(final_pptx).resolve())


def _move_slides_to_final(
    staging_slides: str, stage: RenderStage, slides_output_dir: str | None
) -> str:
    final_slides = slides_output_dir or _default_slides_dir(stage.final_pptx)
    Path(final_slides).mkdir(exist_ok=True, parents=True)
    # 只清理本 deck 先前落位的产物（slide_*.png + _deck_info.json），避免 re-render
    # 页数变少时残留旧页。归属校验确保绝不删用户其他文件——slides_output_dir 可能
    # 指向共享目录，里面同名的 slide_*.png 不是本 deck 的。
    if _slides_dir_owned(final_slides, stage.final_pptx):
        for f in Path(final_slides).iterdir():
            if f.is_file() and (f.name == "_deck_info.json" or _is_slide_png(f.name)):
                f.unlink()
    for f in Path(staging_slides).iterdir():
        if f.is_file():
            shutil.copy2(str(f), str(Path(final_slides) / f.name))
    _write_deck_info(final_slides, stage.final_pptx)
    return final_slides


# #22：open_live 前把 .pptx 复制到系统临时目录的 offipy-live-* 副本再让 PowerPoint
# 打开——PowerPoint 锁定的是副本，源产物路径永不被锁，同路径 re-render(overwrite=True)
# 不再 PermissionError。副本由 close_live 删除，废弃残留由 _cleanup_stale_live_tmp 兜底。
_LIVE_TMP_PREFIX = "offipy-live-"
# doc_id → 临时副本路径（open_live 登记，close_live 清理）。会话级尽力而为：
# server 重启 / 直接 ppt.close_pres 关闭时靠 _cleanup_stale_live_tmp 兜底。
_LIVE_TMP_PATHS: dict[str, str] = {}


def _remove_stale_live_tmp(f: Path, stale_before: float) -> None:
    """删单个过期副本；仍被 PowerPoint 打开（无共享删除）时跳过，不中断清理。"""
    try:
        if f.stat().st_mtime < stale_before:
            f.unlink()
    except OSError:
        pass


def _cleanup_stale_live_tmp() -> None:
    """清理废弃的 offipy-live-* 副本（关闭/崩溃/重启后残留）。"""
    import time

    stale_before = time.time() - 3600  # 1 小时未动即视为废弃
    tmp_dir = Path(tempfile.gettempdir())
    for f in tmp_dir.glob(f"{_LIVE_TMP_PREFIX}*.pptx"):
        _remove_stale_live_tmp(f, stale_before)


def _live_tmp_copy(pptx: str) -> str:
    """把 .pptx 复制成 offipy-live-* 临时副本，返回副本路径。"""
    src = Path(pptx).resolve()
    if not Path(src).exists():
        raise InvalidArgumentError(f"源 .pptx 不存在: {src}")
    _cleanup_stale_live_tmp()
    fd, tmp = tempfile.mkstemp(prefix=_LIVE_TMP_PREFIX, suffix=".pptx")
    os.close(fd)
    Path(tmp).unlink()  # 先删占位，copyfile 以正常权限全新创建
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
            Path(live).unlink()  # 打开失败不留孤儿副本
        raise
    _LIVE_TMP_PATHS[doc_id] = live
    return str(doc_id)


def close_live(doc_id: str) -> None:
    """关闭 open_live 打开的实况演示，释放其占用的临时副本（#22）。

    save=False 直接关闭不保存——实况展示操作的是临时副本，回写无意义。
    配合 #26 的 Ppt.close_pres：关闭后同路径 re-render 不再 PermissionError。
    """
    ensure_server()
    try:
        call("ppt", "close_pres", doc_id=doc_id, save=False)
    finally:
        # 即使 close_pres 异常（如 PowerPoint 已崩、server 已断），临时副本也要清理，
        # 不残留 offipy-live-* 孤儿。
        path = _LIVE_TMP_PATHS.pop(doc_id, None)
        if path:
            with contextlib.suppress(OSError):
                Path(path).unlink()


def export_slides(
    out_dir: str,
    width: int = 1920,
    height: int = 1080,
    doc_id: str | None = None,
    overwrite: bool = False,
) -> list[str]:
    """把 doc_id 指定的演示文稿逐页导出 PNG，供 Claude 视觉迭代。

    doc_id 必须显式传入（本 API 不隐式跟随活动文档，无 follow_active 参数）；
    overwrite=False（默认）拒绝覆盖已有输出；True 时原子替换（不残留半成品）。
    """
    ensure_server()
    return cast(
        "list[str]",
        call(
            "ppt",
            "export_slides",
            out_dir=str(Path(out_dir).resolve()),
            width=width,
            height=height,
            doc_id=doc_id,
            overwrite=overwrite,
        ),
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
        try:
            export_slides(feedback_dir, doc_id=doc_id, overwrite=overwrite)
        finally:
            close_live(doc_id)  # 导出后即释放实况窗口 + 临时副本，不留泄漏
    elif open_live_flag:
        open_live(pptx)
    return pptx
