"""规则框架：RuleSpec / RuleContext / RuleEvaluation / 三分离维度聚合。

rev2.1：
- RuleSpec 显式绑定 rule_id + dimension + run，替代动态函数属性；
- 规则函数返回 RuleEvaluation(findings, covered_count, eligible_count, warnings)；
- assess_dimension：先 active 规则 → not_applicable；再 coverage → insufficient_evidence；
- grade 只由 quality_penalty 决定；confidence = coverage × applicability × reliability。
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass, field

from offipy.audit import Severity

from .models import (
    ArtElement,
    ArtFinding,
    ArtScene,
    ArtSlide,
    ArtWarning,
    DimensionAssessment,
    Grade,
)
from .profiles import ArtProfile

_SEVERITY_WEIGHT = {Severity.LOW: 0.5, Severity.MID: 1.5, Severity.HIGH: 3.0}
# 契约：finding_confidence < 0.35 不驱动降级（低置信发现不进 penalty 求和）
_GRADE_CONFIDENCE_FLOOR = 0.35
_COVERAGE_MIN = 0.5

_GRADE_THRESHOLDS = (0.0, 1.0, 2.5)  # ≤0 excellent, ≤1 good, ≤2.5 attention, else poor


@dataclass(frozen=True)
class RuleSpec:
    rule_id: str
    dimension: str
    run: Callable[[ArtSlide, RuleContext], RuleEvaluation]
    experimental: bool = False


@dataclass
class RuleContext:
    profile: ArtProfile
    slide: ArtSlide
    slide_index: int
    features: dict
    deck: ArtScene
    sources: frozenset[str] = frozenset({"measurement"})
    # 页面级规则可能在 slide 维度需要看整个 scene 的其它页；默认当前页
    page_scope: list[ArtSlide] = field(default_factory=list)


@dataclass
class RuleEvaluation:
    findings: list[ArtFinding] = field(default_factory=list)
    covered_count: int = 0
    eligible_count: int = 0
    warnings: list[ArtWarning] = field(default_factory=list)
    reliability: float | None = None


def make_finding(
    rule_id: str,
    dimension: str,
    severity: Severity,
    message: str,
    confidence: float,
    slide_index: int | None = None,
    primary: object | None = None,
    related: Iterable[object] | None = None,
    details: dict | None = None,
    evidence_sources: Iterable[str] | None = None,
    evidence_reliability: float | None = None,
    evidence_method: str | None = None,
) -> ArtFinding:
    from .models import ArtElementRef

    def ref(e) -> ArtElementRef | None:
        if e is None:
            return None
        if isinstance(e, ArtElementRef):
            return e
        if isinstance(e, ArtElement):
            return ArtElementRef(e.slide_index, e.element_id, e.kind, e.role)
        return ArtElementRef(e.slide_index, e.element_id, e.kind, e.role)

    return ArtFinding(
        rule_id=rule_id,
        dimension=dimension,
        severity=severity,
        message=message,
        confidence=confidence,
        slide_index=slide_index,
        primary=ref(primary),
        related=[r for r in (ref(x) for x in (related or [])) if r is not None],
        details=details or {},
        evidence_sources=frozenset(evidence_sources or ()),
        evidence_reliability=evidence_reliability,
        evidence_method=evidence_method,
    )


def _effective_penalty(severity: Severity, confidence: float) -> float:
    return _SEVERITY_WEIGHT[severity] * confidence


def grade_from_findings(findings: list[ArtFinding]) -> Grade:
    """质量评级：只由 quality_penalty 决定，与判断置信度无关。"""
    penalty = 0.0
    for f in findings:
        if f.confidence < _GRADE_CONFIDENCE_FLOOR:
            continue  # 低置信 finding 不驱动降级
        penalty += _effective_penalty(f.severity, f.confidence)
    if penalty <= _GRADE_THRESHOLDS[0]:
        return "excellent"
    if penalty <= _GRADE_THRESHOLDS[1]:
        return "good"
    if penalty <= _GRADE_THRESHOLDS[2]:
        return "attention"
    return "poor"


def _is_active(spec: RuleSpec, profile: ArtProfile) -> bool:
    """规则开关：enabled_rules 正向白名单 + disabled_rules 负向开关，双重过滤。"""
    return spec.rule_id in profile.enabled_rules and spec.rule_id not in profile.disabled_rules


def _apply_profile(
    finding: ArtFinding, profile: ArtProfile, experimental: bool = False
) -> ArtFinding:
    sev = profile.severity_overrides.get(finding.rule_id)
    conf = profile.confidence_overrides.get(finding.rule_id)
    if sev is not None:
        finding.severity = sev
    if conf is not None:
        finding.confidence = float(conf)
    # RuleSpec.experimental 是唯一权威来源；profile.experimental_rules 兼容叠加
    if experimental or finding.rule_id in profile.experimental_rules:
        finding.confidence = min(finding.confidence, 0.3)
    return finding


def _scene_reliability(ctx: RuleContext) -> float:
    """场景证据源可靠度：measurement 1.0；pixel 0.55；仅几何 0.6。"""
    if "measurement" in ctx.sources:
        return 1.0
    if "pixel" in ctx.sources:
        return 0.55
    return 0.6


def assess_dimension(
    dimension: str,
    rule_specs: list[RuleSpec],
    ctx: RuleContext,
) -> DimensionAssessment:
    """聚合一个维度：先判 active 规则，再算 coverage，最后 grade/confidence 分离。"""
    active = [rs for rs in rule_specs if _is_active(rs, ctx.profile)]
    if not active:
        return DimensionAssessment(dimension=dimension, status="not_applicable")
    findings: list[ArtFinding] = []
    covered = 0
    eligible = 0
    warnings: list[ArtWarning] = []
    applicable = 0
    reliability_terms: list[float] = []
    reliability_weights: list[int] = []
    for rs in active:
        ev = rs.run(ctx.slide, ctx)
        covered += ev.covered_count
        eligible += ev.eligible_count
        warnings.extend(ev.warnings)
        if ev.eligible_count > 0:
            applicable += 1
        if (
            ev.reliability is not None
            and ev.covered_count > 0
            and not rs.experimental
            and rs.rule_id not in ctx.profile.experimental_rules
        ):
            reliability_terms.append(ev.reliability)
            reliability_weights.append(ev.covered_count)
        for f in ev.findings:
            findings.append(_apply_profile(f, ctx.profile, experimental=rs.experimental))
    coverage = (covered / eligible) if eligible else 0.0
    if coverage < _COVERAGE_MIN:
        # 证据不足 → 不误报：丢弃低置信 finding，只保留 coverage 状态与 warnings
        return DimensionAssessment(
            dimension=dimension,
            status="insufficient_evidence",
            evidence_coverage=round(coverage, 4),
            findings=[],
            warnings=warnings,
        )
    grade = grade_from_findings(findings)
    applicability = (applicable / len(active)) if active else 0.0
    weight_total = sum(reliability_weights)
    if weight_total > 0:
        reliability = (
            sum(t * w for t, w in zip(reliability_terms, reliability_weights, strict=True))
            / weight_total
        )
        minimum_reliability = min(reliability_terms)
    else:
        reliability = _scene_reliability(ctx)
        minimum_reliability = reliability
    conf = round(coverage * applicability * reliability, 4)
    return DimensionAssessment(
        dimension=dimension,
        status="assessed",
        grade=grade,
        confidence=conf,
        evidence_coverage=round(coverage, 4),
        findings=findings,
        warnings=warnings,
        reliability=round(reliability, 4),
        minimum_reliability=round(minimum_reliability, 4),
    )
