"""建议投影：把 DeckQualityReport 投影成确定性建议记录（Task 5 / S3）。

建议策略：确定、不调 LLM。显式 remediation 只取 finding.details 里携带的
"suggestion"/"fix"/"remediation" 键（非空字符串）；今日规则消息绝大多数是描述
性的（如「标题字号过小（font_size_norm=0.018）。」），没有结构化修复指令 → 一律
"无自动建议"。details["suggestion"] 是未来规则可挂结构化修复钩子的预留位。

记录 schema（7 键，顺序固定，--json 稳定性依赖它）：
    dimension / slide_index / rule_id / element / severity / message / suggestion
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from .models import ArtElementRef, ArtFinding, DeckQualityReport

_NO_SUGGESTION = "无自动建议"

_REMEDIATION_KEYS = ("suggestion", "fix", "remediation")


def _element_label(primary: ArtElementRef | None) -> str:
    """主元素的可读标签：kind:role；无 role 回退 element_id；无引用空串。"""
    if primary is None:
        return ""
    if primary.role:
        return f"{primary.kind}:{primary.role}"
    return primary.element_id


def _remediation_for(finding: ArtFinding) -> str | None:
    """从 details 读取显式 remediation 提示（非空字符串）；无则 None。"""
    details = finding.details or {}
    for key in _REMEDIATION_KEYS:
        value = details.get(key)
        if isinstance(value, str) and value.strip():
            return value
    return None


def _record(finding: ArtFinding, *, fallback_slide: int | None) -> dict[str, Any]:
    """单条 finding → 建议记录（键序固定，投影确定）。"""
    return {
        "dimension": finding.dimension,
        "slide_index": finding.slide_index if finding.slide_index is not None else fallback_slide,
        "rule_id": finding.rule_id,
        "element": _element_label(finding.primary),
        "severity": finding.severity.name,
        "message": finding.message,
        "suggestion": _remediation_for(finding) or _NO_SUGGESTION,
    }


def project_suggestions(report: DeckQualityReport, *, source: str) -> list[dict[str, Any]]:  # noqa: ARG001 — source 是文档化的预留接口参数（见 docstring），当前记录不携带来源
    """把 DeckQualityReport 投影成确定性建议记录列表。

    art 为空（无艺术分析，如 PPTX-only 证据不足路径）→ 空列表。遍历顺序：
    每个 slide 的每个 dimension 的每个 finding，随后 deck_findings。source 预留
    供未来记录透传来源；当前记录本身不含 source。
    """
    records: list[dict[str, Any]] = []
    art = report.art
    if art is None:
        return records
    for slide in art.slides:
        for dim in slide.dimensions:
            records.extend(
                _record(finding, fallback_slide=slide.slide_index) for finding in dim.findings
            )
    records.extend(_record(finding, fallback_slide=None) for finding in art.deck_findings)
    return records
