"""审计/回归报告渲染：text / markdown / html。

json 由模型 to_json() 提供（完全 JSON 安全）；text/markdown/html 在此实现。
html 为单文件（内联 CSS/JS/SVG，不依赖外网 CDN）；PptxAuditReport 用 SVG
逐页画布，PptxDiffReport 用结构化表格。

Severity 比较一律按整数值（IntEnum），禁止字符串比较。
"""

from __future__ import annotations

import html as _html
from collections import defaultdict
from typing import Any

from .models import (
    AuditFinding,
    AuditWarning,
    PptxAuditReport,
    PptxDiffReport,
    Severity,
    SlideShapeSnapshot,
    SuppressedFinding,
)

_SEV_COLOR = {
    Severity.HIGH: "#d32f2f",
    Severity.MID: "#f9a825",
    Severity.LOW: "#1976d2",
}
_SUP_COLOR = "#7b1fa2"
_WARN_COLOR = "#5d4037"
_NEUTRAL_FILL = "#fafafa"
_NEUTRAL_STROKE = "#bdbdbd"


# ---------------------------------------------------------------- 公共入口


def render_text(report: PptxAuditReport | PptxDiffReport) -> str:
    if isinstance(report, PptxDiffReport):
        return _text_diff(report)
    return _text_audit(report)


def render_markdown(report: PptxAuditReport | PptxDiffReport) -> str:
    if isinstance(report, PptxDiffReport):
        return _markdown_diff(report)
    return _markdown_audit(report)


def render_html(report: PptxAuditReport | PptxDiffReport, *, slides_dir: str | None = None) -> str:
    if isinstance(report, PptxDiffReport):
        return _html_diff(report)
    return _html_audit(report, slides_dir)


# ---------------------------------------------------------------- 通用片段


def _fmt_num(v: Any) -> str:
    if isinstance(v, float):
        return f"{v:.3f}"
    return str(v)


def _details_text(details: dict[str, Any]) -> str:
    return " ".join(f"{k}={_fmt_num(v)}" for k, v in details.items())


def _md_cell(v: str) -> str:
    return v.replace("|", "\\|").replace("\n", " ")


def _finding_line(f: AuditFinding) -> str:
    p = f.primary
    name = f'"{p.name}"' if p.name else "(无名称)"
    sec = ""
    if f.secondary is not None:
        sec = f' vs #{f.secondary.shape_id} "{f.secondary.name}"'
    return f"[{f.severity.name}] {f.rule_id} #{p.shape_id} {name}{sec} {f.message}"


def _suppressed_line(s: SuppressedFinding) -> str:
    f = s.finding
    p = f.primary
    name = f'"{p.name}"' if p.name else "(无名称)"
    return f"[suppressed:{s.reason}] {f.rule_id} #{p.shape_id} {name} {f.message}"


def _warning_line(w: AuditWarning) -> str:
    sid = f"#{w.shape_id}" if w.shape_id is not None else ""
    return f"[warning] {w.code} {sid} {w.message}"


def _finding_table_rows(findings: list[AuditFinding]) -> list[str]:
    rows = ["| 严重度 | rule_id | 形状 | 描述 |", "|---|---|---|---|"]
    for f in findings:
        p = f.primary
        name = p.name or "(无名称)"
        sec = ""
        if f.secondary is not None:
            sec = f" vs #{f.secondary.shape_id}"
        rows.append(
            f"| {f.severity.name} | {f.rule_id} | #{p.shape_id} "
            f'"{_md_cell(name)}"{sec} | {_md_cell(f.message)} |'
        )
    return rows


# ---------------------------------------------------------------- text：审计


def _audit_sections(report: PptxAuditReport) -> list[list[str]]:
    by_slide: dict[int, list[AuditFinding]] = defaultdict(list)
    for f in report.findings:
        by_slide[f.primary.slide_index].append(f)
    sup_by_slide: dict[int, list[SuppressedFinding]] = defaultdict(list)
    for s in report.suppressed:
        sup_by_slide[s.finding.primary.slide_index].append(s)
    warn_by_slide: dict[int, list[AuditWarning]] = defaultdict(list)
    for wd in report.warnings:
        if wd.slide_index is not None:
            warn_by_slide[wd.slide_index].append(wd)

    sections: list[list[str]] = []
    for idx in range(1, report.slide_count + 1):
        fs = by_slide.get(idx, [])
        sups = sup_by_slide.get(idx, [])
        warns = warn_by_slide.get(idx, [])
        if not fs and not sups and not warns:
            continue
        sec = [f"第 {idx} 页 / {report.slide_count}"]
        sec += [_finding_line(f) for f in fs]
        sec += [_suppressed_line(s) for s in sups]
        sec += [_warning_line(w) for w in warns]
        sec.append("")
        sections.append(sec)
    return sections


def _text_audit(report: PptxAuditReport) -> str:
    w, h = report.slide_size
    out: list[str] = []
    out.append(f"审计报告: {report.path}")
    out.append(f"schema {report.schema_version} | offipy {report.offipy_version}")
    out.append(f"页面 {w:.2f} x {h:.2f} in | {report.slide_count} 页")
    out.append(f"sha256 {report.source_sha256}")
    out.append("")
    for section in _audit_sections(report):
        out.extend(section)
    counts = dict.fromkeys(Severity, 0)
    for f in report.findings:
        counts[f.severity] += 1
    out.append("概要")
    out.append(
        f"HIGH {counts[Severity.HIGH]} | MID {counts[Severity.MID]}"
        f" | LOW {counts[Severity.LOW]}"
        f" | suppressed {len(report.suppressed)}"
        f" | warnings {len(report.warnings)}"
    )
    return "\n".join(out).rstrip() + "\n"


# ---------------------------------------------------------------- text：对比


def _append_shape_changes(out: list[str], report: PptxDiffReport) -> None:
    groups = (
        ("新增", report.added_shapes),
        ("删除", report.removed_shapes),
        ("移动", report.moved_shapes),
        ("缩放", report.resized_shapes),
        ("文本", report.text_changes),
    )
    if not any(items for _, items in groups):
        out.append("  无")
        return
    for label, items in groups:
        for s in items:
            name = f'"{s.name}"' if s.name else "(无名称)"
            out.append(f"  [{label}] 第{s.slide_index}页 #{s.shape_id} {name}")
            det = _details_text(s.details)
            if det:
                out.append(f"    {det}")


def _append_finding_changes(out: list[str], report: PptxDiffReport) -> None:
    out.append("  新增:")
    if report.added_findings:
        out.extend(f"    {_finding_line(f)}" for f in report.added_findings)
    else:
        out.append("    无")
    out.append("  已解决:")
    if report.resolved_findings:
        out.extend(f"    {_finding_line(f)}" for f in report.resolved_findings)
    else:
        out.append("    无")
    out.append("  变化:")
    if report.changed_findings:
        for c in report.changed_findings:
            p = c.primary
            name = f'"{p.name}"' if p.name else "(无名称)"
            mark = " (恶化)" if c.worsened else ""
            out.append(
                f"    [{c.old_severity.name}->{c.new_severity.name}] "
                f"{c.rule_id} #{p.shape_id} {name}{mark}"
            )
    else:
        out.append("    无")


def _text_diff(report: PptxDiffReport) -> str:
    out: list[str] = []
    out.append("回归对比")
    out.append(f"基线: {report.baseline_path}")
    out.append(f"候选: {report.candidate_path}")
    out.append(f"页面 {report.baseline_slide_count} -> {report.candidate_slide_count}")
    out.append(f"sha256 基线 {report.baseline_sha256}")
    out.append(f"sha256 候选 {report.candidate_sha256}")
    out.append("")
    if report.added_slides or report.removed_slides:
        out.append(f"幻灯片: 新增 {report.added_slides} 页, 删除 {report.removed_slides} 页")
        out.append("")
    out.append("形状变化")
    _append_shape_changes(out, report)
    out.append("")
    out.append("问题变化")
    _append_finding_changes(out, report)
    if report.unmatched_baseline or report.unmatched_candidate:
        out.append("")
        out.append("无法可靠匹配")
        out.extend(
            f'  [基线] 第{r.slide_index}页 #{r.shape_id} "{r.name}"'
            for r in report.unmatched_baseline
        )
        out.extend(
            f'  [候选] 第{r.slide_index}页 #{r.shape_id} "{r.name}"'
            for r in report.unmatched_candidate
        )
    if report.warnings:
        out.append("")
        out.append("警告")
        out.extend(f"  {_warning_line(wd)}" for wd in report.warnings)
    gate = report.gate_severity()
    if gate is not None:
        out.append("")
        out.append(f"新增/恶化最高严重度: {gate.name}")
    return "\n".join(out).rstrip() + "\n"


# ---------------------------------------------------------------- markdown


def _markdown_audit(report: PptxAuditReport) -> str:
    w, h = report.slide_size
    out: list[str] = []
    out.append(f"# 审计报告: {report.path}")
    out.append("")
    out.append(f"- schema: {report.schema_version}")
    out.append(f"- offipy: {report.offipy_version}")
    out.append(f"- 页面: {w:.2f} x {h:.2f} in")
    out.append(f"- 页数: {report.slide_count}")
    out.append(f"- sha256: `{report.source_sha256}`")
    out.append("")
    for idx in range(1, report.slide_count + 1):
        fs = [f for f in report.findings if f.primary.slide_index == idx]
        sups = [s for s in report.suppressed if s.finding.primary.slide_index == idx]
        warns = [w for w in report.warnings if w.slide_index == idx]
        if not fs and not sups and not warns:
            continue
        out.append(f"## 第 {idx} 页 / {report.slide_count}")
        out.append("")
        if fs:
            out.extend(_finding_table_rows(fs))
            out.append("")
        if sups:
            out.append("### 豁免")
            for s in sups:
                p = s.finding.primary
                name = p.name or "(无名称)"
                out.append(
                    f"- `{s.reason}` {s.finding.rule_id} #{p.shape_id} "
                    f'"{_md_cell(name)}" {_md_cell(s.finding.message)}'
                )
            out.append("")
        if warns:
            out.append("### 警告")
            for wd in warns:
                sid = f"#{wd.shape_id}" if wd.shape_id is not None else ""
                out.append(f"- `{wd.code}` {sid} {_md_cell(wd.message)}")
            out.append("")
    counts = dict.fromkeys(Severity, 0)
    for f in report.findings:
        counts[f.severity] += 1
    out.append("## 概要")
    out.append("")
    out.append("| 严重度 | 数量 |")
    out.append("|---|---|")
    out.extend(
        f"| {sev.name} | {counts[sev]} |" for sev in (Severity.HIGH, Severity.MID, Severity.LOW)
    )
    out.append(f"| 豁免 | {len(report.suppressed)} |")
    out.append(f"| 警告 | {len(report.warnings)} |")
    out.append("")
    return "\n".join(out).rstrip() + "\n"


def _markdown_shape_table(out: list[str], report: PptxDiffReport) -> None:
    out.append("## 形状变化")
    out.append("")
    groups = (
        ("新增", report.added_shapes),
        ("删除", report.removed_shapes),
        ("移动", report.moved_shapes),
        ("缩放", report.resized_shapes),
        ("文本", report.text_changes),
    )
    if not any(items for _, items in groups):
        out.append("无变化")
        out.append("")
        return
    out.append("| 类型 | 页面 | shape_id | 名称 | 详情 |")
    out.append("|---|---|---|---|---|")
    for label, items in groups:
        for s in items:
            name = s.name or "(无名称)"
            det = _details_text(s.details)
            out.append(
                f"| {label} | {s.slide_index} | {s.shape_id} | {_md_cell(name)} | {_md_cell(det)} |"
            )
    out.append("")


def _markdown_finding_table(title: str, findings: list[AuditFinding]) -> list[str]:
    out = [f"## {title}", ""]
    out.extend(_finding_table_rows(findings))
    out.append("")
    return out


def _markdown_diff(report: PptxDiffReport) -> str:
    gate = report.gate_severity()
    out: list[str] = []
    out.append("# 回归对比")
    out.append("")
    out.append(f"- 基线: `{report.baseline_path}`")
    out.append(f"- 候选: `{report.candidate_path}`")
    out.append(f"- 页面: {report.baseline_slide_count} -> {report.candidate_slide_count}")
    out.append(f"- 基线 sha256: `{report.baseline_sha256}`")
    out.append(f"- 候选 sha256: `{report.candidate_sha256}`")
    out.append(f"- 新增/恶化最高严重度: {gate.name if gate is not None else '无'}")
    out.append("")
    out.append("## 幻灯片")
    out.append("")
    out.append(f"- 新增: {report.added_slides} 页")
    out.append(f"- 删除: {report.removed_slides} 页")
    out.append("")
    _markdown_shape_table(out, report)
    if report.added_findings:
        out.extend(_markdown_finding_table("新增问题", report.added_findings))
    if report.resolved_findings:
        out.extend(_markdown_finding_table("已解决问题", report.resolved_findings))
    if report.changed_findings:
        out.append("## 变化问题")
        out.append("")
        out.append("| rule_id | 严重度变化 | 恶化 | 形状 |")
        out.append("|---|---|---|---|")
        for c in report.changed_findings:
            p = c.primary
            name = p.name or "(无名称)"
            out.append(
                f"| {c.rule_id} | {c.old_severity.name} -> {c.new_severity.name} "
                f"| {'是' if c.worsened else '否'} | #{p.shape_id} {_md_cell(name)} |"
            )
        out.append("")
    if report.unmatched_baseline or report.unmatched_candidate:
        out.append("## 无法可靠匹配")
        out.append("")
        out.extend(
            f"- 基线: 第{r.slide_index}页 #{r.shape_id} {r.name}" for r in report.unmatched_baseline
        )
        out.extend(
            f"- 候选: 第{r.slide_index}页 #{r.shape_id} {r.name}"
            for r in report.unmatched_candidate
        )
        out.append("")
    if report.warnings:
        out.append("## 警告")
        out.append("")
        for wd in report.warnings:
            sid = f"#{wd.shape_id}" if wd.shape_id is not None else ""
            out.append(f"- `{wd.code}` {sid} {_md_cell(wd.message)}")
        out.append("")
    return "\n".join(out).rstrip() + "\n"


# ---------------------------------------------------------------- html


_AUDIT_CSS = """
body{font-family:system-ui,'Segoe UI',sans-serif;margin:16px;color:#212121}
h1{font-size:20px}h2{font-size:16px}h3{font-size:14px}
p{font-size:13px}
.filters button{padding:4px 10px;margin:0 6px 6px 0;cursor:pointer;border:1px solid #bbb;
border-radius:4px;background:#fff}
.filters button.active{background:#1976d2;color:#fff;border-color:#1976d2}
.slide{margin:28px 0}
.slide svg{border:1px solid #ddd;max-width:100%;height:auto;background:#fff}
.audit-detail{width:100%;margin-top:6px}
.shape-label{font-size:9px;fill:#616161;pointer-events:none}
.legend .sw{display:inline-block;width:14px;height:14px;margin:0 4px -2px 12px}
table{border-collapse:collapse;margin:8px 0}
th,td{border:1px solid #ddd;padding:4px 8px;text-align:left;font-size:13px}
th{background:#f5f5f5}
.high{color:#d32f2f;font-weight:600}
.mid{color:#f9a825;font-weight:600}
.low{color:#1976d2;font-weight:600}
"""

_AUDIT_JS = """
function filterAudit(kind){
  var el=document.querySelectorAll('.shape');
  for(var i=0;i<el.length;i++){
    var show=(kind==='all')||el[i].classList.contains(kind);
    el[i].style.display=show?'':'none';
  }
  var btns=document.querySelectorAll('.filters button');
  for(var j=0;j<btns.length;j++){
    btns[j].classList.toggle('active',btns[j].getAttribute('data-kind')===kind);
  }
}
"""


def _html_page_start(title: str, css: str) -> list[str]:
    return [
        "<!DOCTYPE html>",
        '<html lang="zh-CN"><head><meta charset="utf-8">',
        f"<title>{_html.escape(title)}</title>",
        f"<style>{css}</style>",
        "</head><body>",
        f"<h1>{_html.escape(title)}</h1>",
    ]


def _html_page_end(js: str) -> str:
    return f"<script>{js}</script></body></html>"


def _svg_coord(v: float | None, scale: float) -> float:
    return (v or 0.0) * scale


def _snap_intersection(
    a: SlideShapeSnapshot | None, b: SlideShapeSnapshot | None, scale: float
) -> tuple[float, float, float, float] | None:
    if a is None or b is None:
        return None
    if (
        a.left is None
        or a.top is None
        or a.width is None
        or a.height is None
        or b.left is None
        or b.top is None
        or b.width is None
        or b.height is None
    ):
        return None
    ax0, ay0 = a.left, a.top
    aw, ah = a.width, a.height
    bx0, by0 = b.left, b.top
    bw, bh = b.width, b.height
    x0 = max(ax0, bx0)
    y0 = max(ay0, by0)
    x1 = min(ax0 + aw, bx0 + bw)
    y1 = min(ay0 + ah, by0 + bh)
    if x1 - x0 <= 0 or y1 - y0 <= 0:
        return None
    return (x0 * scale, y0 * scale, (x1 - x0) * scale, (y1 - y0) * scale)


def _svg_slide(
    report: PptxAuditReport,
    idx: int,
    shapes: list[SlideShapeSnapshot],
    findings: list[AuditFinding],
    suppressed: list[SuppressedFinding],
    warnings: list[AuditWarning],
    slides_dir: str | None,
) -> str:
    w, h = report.slide_size
    scale = 100.0
    lines = [f'<svg viewBox="0 0 {w * scale:.0f} {h * scale:.0f}" role="img">']
    if slides_dir:
        img = f"{slides_dir.replace(chr(92), '/')}/slide-{idx}.png"
        lines.append(
            f'<image href="{_html.escape(img)}" x="0" y="0" '
            f'width="{w * scale:.0f}" height="{h * scale:.0f}" '
            'preserveAspectRatio="none"/>'
        )
    lines.append(
        f'<rect x="0" y="0" width="{w * scale:.0f}" height="{h * scale:.0f}" '
        'fill="none" stroke="#000" stroke-width="2"/>'
    )
    m = report.config.safe_margin_in * scale
    safe_w = (w - 2 * report.config.safe_margin_in) * scale
    safe_h = (h - 2 * report.config.safe_margin_in) * scale
    lines.append(
        f'<rect x="{m:.1f}" y="{m:.1f}" width="{safe_w:.1f}" '
        f'height="{safe_h:.1f}" fill="none" '
        'stroke="#90caf9" stroke-width="1" stroke-dasharray="4,4"/>'
    )

    by_id = {shp.shape_id: shp for shp in shapes}
    sev_by_id: dict[int, list[Severity]] = defaultdict(list)
    for f in findings:
        sev_by_id[f.primary.shape_id].append(f.severity)
        if f.secondary is not None:
            sev_by_id[f.secondary.shape_id].append(f.severity)
    sup_ids: set[int] = set()
    for sf in suppressed:
        sup_ids.add(sf.finding.primary.shape_id)
        if sf.finding.secondary is not None:
            sup_ids.add(sf.finding.secondary.shape_id)
    warn_ids = {wd.shape_id for wd in warnings if wd.shape_id is not None}

    for shp in shapes:
        if shp.left is None or shp.top is None or shp.width is None or shp.height is None:
            continue
        sevs = sev_by_id.get(shp.shape_id)
        cls = "shape"
        if sevs:
            sev = max(sevs)
            cls += f" sev-{sev.name.lower()}"
            fill = _SEV_COLOR[sev]
            stroke = _SEV_COLOR[sev]
            dash = ""
        elif shp.shape_id in sup_ids:
            cls += " sup"
            fill = "#f3e5f5"
            stroke = _SUP_COLOR
            dash = "5,3"
        elif shp.shape_id in warn_ids:
            cls += " warn"
            fill = "#fff8e1"
            stroke = _WARN_COLOR
            dash = "5,3"
        else:
            cls += " ok"
            fill = _NEUTRAL_FILL
            stroke = _NEUTRAL_STROKE
            dash = ""
        x = _svg_coord(shp.left, scale)
        y = _svg_coord(shp.top, scale)
        wd = _svg_coord(shp.width, scale)
        ht = _svg_coord(shp.height, scale)
        dash_attr = f' stroke-dasharray="{dash}"' if dash else ""
        lines.append(
            f'<rect class="{cls}" x="{x:.1f}" y="{y:.1f}" width="{wd:.1f}" '
            f'height="{ht:.1f}" fill="{fill}" stroke="{stroke}" '
            f'stroke-width="1.5"{dash_attr}/>'
        )
        label = f"#{shp.shape_id} {shp.name}"
        if shp.is_rotated:
            label += " (旋转)"
        if shp.geometry_unknown:
            label += " (几何未知)"
        lines.append(
            f'<text class="shape-label" x="{x + 1:.1f}" y="{y + 8:.1f}">'
            f"{_html.escape(label)}</text>"
        )

    for f in findings:
        if f.kind != "overlap" or f.secondary is None:
            continue
        inter = _snap_intersection(
            by_id.get(f.primary.shape_id), by_id.get(f.secondary.shape_id), scale
        )
        if inter is not None:
            ix, iy, iw, ih = inter
            lines.append(
                f'<rect class="shape overlap-area" x="{ix:.1f}" y="{iy:.1f}" '
                f'width="{iw:.1f}" height="{ih:.1f}" fill="{_SEV_COLOR[f.severity]}" '
                'fill-opacity="0.25" stroke="none"/>'
            )

    lines.append("</svg>")
    return "\n".join(lines)


def _html_detail_tables(
    parts: list[str],
    fs: list[AuditFinding],
    sups: list[SuppressedFinding],
    warns: list[AuditWarning],
) -> None:
    """每页 SVG 下方的可读问题详情表（行 class 与 SVG 一致，参与严重度筛选）。"""
    if fs:
        parts.append(
            "<table class='audit-detail'><tr><th>严重度</th><th>rule_id</th>"
            "<th>形状</th><th>关联形状</th><th>描述</th><th>置信度</th></tr>"
        )
        for f in fs:
            p = f.primary
            name = p.name or "(无名称)"
            sev = f.severity.name.lower()
            sec = f"#{f.secondary.shape_id}" if f.secondary is not None else ""
            parts.append(
                f'<tr class="shape sev-{sev}">'
                f'<td class="{sev}">{f.severity.name}</td>'
                f"<td>{f.rule_id}</td>"
                f'<td>#{p.shape_id} "{_html.escape(name)}"</td>'
                f"<td>{sec}</td>"
                f"<td>{_html.escape(f.message)}</td>"
                f"<td>{f.confidence:.2f}</td></tr>"
            )
        parts.append("</table>")
    if sups:
        parts.append(
            "<table class='audit-detail'><tr><th>严重度</th><th>rule_id</th>"
            "<th>形状</th><th>描述</th><th>豁免原因</th></tr>"
        )
        for s in sups:
            f = s.finding
            p = f.primary
            name = p.name or "(无名称)"
            sev = f.severity.name.lower()
            parts.append(
                f'<tr class="shape sup">'
                f'<td class="{sev}">{f.severity.name}</td>'
                f"<td>{f.rule_id}</td>"
                f'<td>#{p.shape_id} "{_html.escape(name)}"</td>'
                f"<td>{_html.escape(f.message)}</td>"
                f"<td>{s.reason}</td></tr>"
            )
        parts.append("</table>")
    if warns:
        parts.append(
            "<table class='audit-detail'><tr><th>code</th><th>页面</th>"
            "<th>形状</th><th>描述</th></tr>"
        )
        for wd in warns:
            sid = f"#{wd.shape_id}" if wd.shape_id is not None else "-"
            parts.append(
                f'<tr class="shape warn"><td>{wd.code}</td>'
                f"<td>{wd.slide_index if wd.slide_index is not None else '-'}</td>"
                f"<td>{sid}</td><td>{_html.escape(wd.message)}</td></tr>"
            )
        parts.append("</table>")


def _html_audit(report: PptxAuditReport, slides_dir: str | None) -> str:
    parts = _html_page_start(f"审计报告: {report.path}", _AUDIT_CSS)
    w, h = report.slide_size
    parts.append(
        f"<p>页面 {w:.2f} x {h:.2f} in | {report.slide_count} 页 | "
        f"sha256 <code>{report.source_sha256}</code></p>"
    )
    counts = dict.fromkeys(Severity, 0)
    for f in report.findings:
        counts[f.severity] += 1
    legend = [
        f'<span class="sw" style="background:{_SEV_COLOR[Severity.HIGH]}"></span>HIGH '
        f"{counts[Severity.HIGH]}",
        f'<span class="sw" style="background:{_SEV_COLOR[Severity.MID]}"></span>MID '
        f"{counts[Severity.MID]}",
        f'<span class="sw" style="background:{_SEV_COLOR[Severity.LOW]}"></span>LOW '
        f"{counts[Severity.LOW]}",
        f'<span class="sw" style="background:{_SUP_COLOR}"></span>豁免 {len(report.suppressed)}',
        f'<span class="sw" style="background:{_WARN_COLOR}"></span>警告 {len(report.warnings)}',
    ]
    parts.append('<p class="legend">' + "".join(legend) + "</p>")
    parts.append('<div class="filters">')
    for label, kind in (
        ("全部", "all"),
        ("HIGH", "sev-high"),
        ("MID", "sev-mid"),
        ("LOW", "sev-low"),
        ("豁免", "sup"),
        ("警告", "warn"),
    ):
        cls = ' class="active"' if kind == "all" else ""
        parts.append(
            f'<button data-kind="{kind}"{cls} onclick="filterAudit(\'{kind}\')">{label}</button>'
        )
    parts.append("</div>")

    by_slide: dict[int, list[AuditFinding]] = defaultdict(list)
    for f in report.findings:
        by_slide[f.primary.slide_index].append(f)
    sup_by_slide: dict[int, list[SuppressedFinding]] = defaultdict(list)
    for s in report.suppressed:
        sup_by_slide[s.finding.primary.slide_index].append(s)
    warn_by_slide: dict[int, list[AuditWarning]] = defaultdict(list)
    for wd in report.warnings:
        if wd.slide_index is not None:
            warn_by_slide[wd.slide_index].append(wd)

    for idx in range(1, report.slide_count + 1):
        shapes = [s for s in report.shapes if s.slide_index == idx]
        fs = by_slide.get(idx, [])
        sups = sup_by_slide.get(idx, [])
        warns = warn_by_slide.get(idx, [])
        parts.append('<div class="slide">')
        parts.append(f"<h2>第 {idx} 页 / {report.slide_count}</h2>")
        parts.append(_svg_slide(report, idx, shapes, fs, sups, warns, slides_dir))
        _html_detail_tables(parts, fs, sups, warns)
        parts.append("</div>")
    parts.append(_html_page_end(_AUDIT_JS))
    return "\n".join(parts)


# ---------------------------------------------------------------- html：对比


_DIFF_CSS = """
body{font-family:system-ui,'Segoe UI',sans-serif;margin:16px;color:#212121}
h1{font-size:20px}h2{font-size:16px}h3{font-size:14px}
p{font-size:13px}
table{border-collapse:collapse;margin:8px 0}
th,td{border:1px solid #ddd;padding:4px 8px;text-align:left;font-size:13px}
th{background:#f5f5f5}
.high{color:#d32f2f;font-weight:600}
.mid{color:#f9a825;font-weight:600}
.low{color:#1976d2;font-weight:600}
"""


def _html_shape_table(parts: list[str], report: PptxDiffReport) -> None:
    groups = (
        ("新增", report.added_shapes),
        ("删除", report.removed_shapes),
        ("移动", report.moved_shapes),
        ("缩放", report.resized_shapes),
        ("文本", report.text_changes),
    )
    if not any(items for _, items in groups):
        parts.append("<p>无变化</p>")
        return
    parts.append(
        "<table><tr><th>类型</th><th>页面</th><th>shape_id</th><th>名称</th><th>详情</th></tr>"
    )
    for label, items in groups:
        for s in items:
            name = s.name or "(无名称)"
            det = _details_text(s.details)
            parts.append(
                f"<tr><td>{label}</td><td>{s.slide_index}</td>"
                f"<td>{s.shape_id}</td><td>{_html.escape(name)}</td>"
                f"<td>{_html.escape(det)}</td></tr>"
            )
    parts.append("</table>")


def _html_finding_sections(parts: list[str], report: PptxDiffReport) -> None:
    if report.added_findings:
        parts.append("<h3>新增</h3>")
        parts.append("<table><tr><th>严重度</th><th>rule_id</th><th>形状</th><th>描述</th></tr>")
        for f in report.added_findings:
            p = f.primary
            name = p.name or "(无名称)"
            parts.append(
                f'<tr><td class="{f.severity.name.lower()}">{f.severity.name}</td>'
                f'<td>{f.rule_id}</td><td>#{p.shape_id} "{_html.escape(name)}"</td>'
                f"<td>{_html.escape(f.message)}</td></tr>"
            )
        parts.append("</table>")
    if report.resolved_findings:
        parts.append("<h3>已解决</h3>")
        parts.append("<table><tr><th>严重度</th><th>rule_id</th><th>形状</th><th>描述</th></tr>")
        for f in report.resolved_findings:
            p = f.primary
            name = p.name or "(无名称)"
            parts.append(
                f'<tr><td class="{f.severity.name.lower()}">{f.severity.name}</td>'
                f'<td>{f.rule_id}</td><td>#{p.shape_id} "{_html.escape(name)}"</td>'
                f"<td>{_html.escape(f.message)}</td></tr>"
            )
        parts.append("</table>")
    if report.changed_findings:
        parts.append("<h3>变化</h3>")
        parts.append(
            "<table><tr><th>rule_id</th><th>严重度变化</th><th>恶化</th><th>形状</th></tr>"
        )
        for c in report.changed_findings:
            p = c.primary
            name = p.name or "(无名称)"
            parts.append(
                f"<tr><td>{c.rule_id}</td>"
                f"<td>{c.old_severity.name} -> {c.new_severity.name}</td>"
                f"<td>{'是' if c.worsened else '否'}</td>"
                f'<td>#{p.shape_id} "{_html.escape(name)}"</td></tr>'
            )
        parts.append("</table>")
    if not (report.added_findings or report.resolved_findings or report.changed_findings):
        parts.append("<p>无问题变化</p>")


def _html_diff(report: PptxDiffReport) -> str:
    parts = _html_page_start(f"回归对比: {report.baseline_path}", _DIFF_CSS)
    gate = report.gate_severity()
    parts.append(
        f"<p>页面 {report.baseline_slide_count} -> {report.candidate_slide_count}"
        f" | 新增/恶化最高严重度: "
        f"<b>{gate.name if gate is not None else '无'}</b></p>"
    )
    parts.append("<h2>幻灯片</h2>")
    parts.append(f"<p>新增 {report.added_slides} 页, 删除 {report.removed_slides} 页</p>")
    parts.append("<h2>形状变化</h2>")
    _html_shape_table(parts, report)
    parts.append("<h2>问题变化</h2>")
    _html_finding_sections(parts, report)
    if report.unmatched_baseline or report.unmatched_candidate:
        parts.append("<h2>无法可靠匹配</h2>")
        for r in report.unmatched_baseline:
            parts.append(f'<p>基线: 第{r.slide_index}页 #{r.shape_id} "{_html.escape(r.name)}"</p>')
        for r in report.unmatched_candidate:
            parts.append(f'<p>候选: 第{r.slide_index}页 #{r.shape_id} "{_html.escape(r.name)}"</p>')
    if report.warnings:
        parts.append("<h2>警告</h2>")
        for wd in report.warnings:
            parts.append(f"<p>{_html.escape(_warning_line(wd))}</p>")
    parts.append(_html_page_end(""))
    return "\n".join(parts)
