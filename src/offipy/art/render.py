"""艺术报告渲染：JSON / Markdown / HTML（含每页维度雷达图）。"""

from __future__ import annotations

import html as html_lib
import math

from .models import ArtFinding, ArtReport, DimensionAssessment

_GRADE_LEVEL = {"excellent": 4, "good": 3, "attention": 2, "poor": 1}

_STATUS_LABEL = {
    "assessed": "已评估",
    "insufficient_evidence": "证据不足",
    "not_applicable": "不适用",
}


def report_to_json(report: ArtReport) -> dict:
    return report.to_dict()


def _finding_evidence(f: ArtFinding) -> str:
    """finding 的证据后缀：有像素/多源证据时展示来源/方法/可靠度。"""
    if not f.evidence_sources:
        return ""
    rel = f"{f.evidence_reliability:.2f}" if f.evidence_reliability is not None else "-"
    return (
        f" (Evidence: {', '.join(sorted(f.evidence_sources))} / "
        f"Method: {f.evidence_method or '-'} / Reliability: {rel})"
    )


def _severity_override_label(f: ArtFinding) -> str | None:
    """严重度调整来源标签：有调整才返回，默认 finding 不产生任何噪音。"""
    if not f.severity_override:
        return None
    if f.severity_override_source == "feedback":
        return "Severity adjusted: feedback"
    if f.severity_override_source == "user":
        return "Severity adjusted: user override"
    return None


def _finding_html(f: ArtFinding) -> str:
    """单个 finding 的 HTML 片段：发现行 + 严重度调整来源（若有）。"""
    text = (
        f"[{f.severity.name}] {html_lib.escape(f.rule_id)}: "
        f"{html_lib.escape(f.message)}{html_lib.escape(_finding_evidence(f))}"
    )
    prov = _severity_override_label(f)
    if prov:
        text += f"<br><span class='severity-adjust'>{prov}</span>"
    return text


def render_markdown(report: ArtReport) -> str:
    lines = [f"# 艺术分析报告 (profile: {report.profile})", ""]
    if report.experimental_score is not None:
        lines.append(f"综合指数 (experimental): {report.experimental_score}")
        lines.append("")
    for s in report.slides:
        lines.append(f"## Slide {s.slide_index}")
        for d in s.dimensions:
            if d.status != "assessed":
                lines.append(
                    f"- **{d.dimension}**: {_STATUS_LABEL[d.status]} "
                    f"(evidence {d.evidence_coverage:.2f})"
                )
                continue
            rel = f" / reliability {d.reliability:.2f}" if d.reliability is not None else ""
            lines.append(f"- **{d.dimension}**: {d.grade} (confidence {d.confidence:.2f}{rel})")
            for f in d.findings:
                lines.append(
                    f"  - [{f.severity.name}] {f.rule_id}: {f.message}{_finding_evidence(f)}"
                )
                prov = _severity_override_label(f)
                if prov:
                    lines.append(f"    - {prov}")
        lines.append("")
    if report.deck_findings:
        lines.append("## Deck 级")
        for f in report.deck_findings:
            lines.append(f"- [{f.severity.name}] {f.rule_id}: {f.message}{_finding_evidence(f)}")
            prov = _severity_override_label(f)
            if prov:
                lines.append(f"  - {prov}")
        lines.append("")
    return "\n".join(lines)


def _radar_svg(assessments: list[DimensionAssessment]) -> str:
    """雷达图。标签固定为「规则评级（非客观美学评分）」，避免被当客观美学分。"""
    n = len(assessments)
    if n < 3:
        return ""
    cx, cy, radius = 110.0, 100.0, 70.0

    def point(i: int, r: float) -> tuple[float, float]:
        ang = -math.pi / 2 + 2 * math.pi * i / n
        return (cx + r * math.cos(ang), cy + r * math.sin(ang))

    grid = ""
    axes = ""
    for lv in range(1, 5):
        pts = " ".join(f"{x:.1f},{y:.1f}" for x, y in (point(i, radius * lv / 4) for i in range(n)))
        grid += f'<polygon points="{pts}" fill="none" stroke="#ddd" stroke-width="1"/>'
    for i in range(n):
        x, y = point(i, radius)
        axes += f'<line x1="{cx}" y1="{cy}" x2="{x:.1f}" y2="{y:.1f}" stroke="#ddd"/>'
    data = " ".join(
        f"{x:.1f},{y:.1f}"
        for x, y in (
            point(i, radius * (_GRADE_LEVEL.get(a.grade, 0) if a.grade else 0) / 4)
            for i, a in enumerate(assessments)
        )
    )
    labels = ""
    for i, a in enumerate(assessments):
        x, y = point(i, radius + 20)
        labels += (
            f'<text x="{x:.1f}" y="{y:.1f}" text-anchor="middle" '
            f'font-size="10">{html_lib.escape(a.dimension)}</text>'
        )
    note = (
        f'<text x="{cx}" y="{cy + radius + 40}" text-anchor="middle" '
        f'font-size="11" fill="#666">规则评级（非客观美学评分）</text>'
    )
    return (
        f'<svg width="{cx * 2 + 44}" height="{cy * 2 + 54}" '
        f'viewBox="0 0 {cx * 2 + 44} {cy * 2 + 54}">'
        f"{grid}{axes}{labels}{note}"
        f'<polygon points="{data}" fill="rgba(59,130,246,0.35)" '
        f'stroke="#3b82f6" stroke-width="2"/>'
        f"</svg>"
    )


def render_html(report: ArtReport) -> str:
    head = (
        "<!DOCTYPE html><html lang='zh'><head><meta charset='utf-8'>"
        "<title>offipy 艺术分析报告</title>"
        "<style>body{font-family:system-ui,sans-serif;margin:24px;color:#111}"
        "table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:6px 10px}"
        ".dim-grade{display:inline-block;padding:2px 8px;border-radius:10px;font-size:12px}"
        ".excellent{background:#dcfce7}.good{background:#e0f2fe}"
        ".attention{background:#fef9c3}.poor{background:#fee2e2}"
        ".severity-adjust{display:block;font-size:11px;color:#b45309;margin-top:2px}"
        "</style></head><body>"
    )
    lines = [head, f"<h1>艺术分析报告 (profile: {html_lib.escape(report.profile)})</h1>"]
    if report.experimental_score is not None:
        lines.append(f"<p>综合指数 (experimental): {report.experimental_score}</p>")
    for s in report.slides:
        lines.append(f"<h2>Slide {s.slide_index}</h2>")
        lines.append(_radar_svg(s.dimensions))
        lines.append(
            "<table><tr><th>维度</th><th>评级</th><th>置信度</th><th>证据</th><th>发现</th></tr>"
        )
        for d in s.dimensions:
            if d.status != "assessed":
                lines.append(
                    f"<tr><td>{html_lib.escape(d.dimension)}</td>"
                    f"<td colspan='4'>{_STATUS_LABEL[d.status]} "
                    f"(evidence {d.evidence_coverage:.2f})</td></tr>"
                )
                continue
            findings = "<br>".join(_finding_html(f) for f in d.findings) or "-"
            conf = f"{d.confidence:.2f}"
            if d.reliability is not None:
                conf += f" / rel {d.reliability:.2f}"
            dim = html_lib.escape(d.dimension)
            grade = html_lib.escape(d.grade) if d.grade else ""
            lines.append(
                f"<tr><td>{dim}</td>"
                f"<td><span class='dim-grade {grade}'>{grade}</span></td>"
                f"<td>{conf}</td>"
                f"<td>{d.evidence_coverage:.2f}</td><td>{findings}</td></tr>"
            )
        lines.append("</table>")
    if report.deck_findings:
        lines.append("<h2>Deck 级</h2><ul>")
        for f in report.deck_findings:
            lines.append(f"<li>{_finding_html(f)}</li>")
        lines.append("</ul>")
    lines.append("</body></html>")
    return "\n".join(lines)
