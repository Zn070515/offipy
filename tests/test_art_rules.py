from offipy.art.models import ArtScene, ArtSlide
from offipy.art.profiles import RULE_TITLE_TOO_SMALL, get_profile
from offipy.art.rules import (
    RuleContext,
    RuleEvaluation,
    RuleSpec,
    assess_dimension,
    grade_from_findings,
    make_finding,
)
from offipy.audit import Severity


def _ctx(slide, profile="balanced", features=None, sources=None):
    return RuleContext(
        profile=get_profile(profile),
        slide=slide,
        slide_index=slide.index,
        features=features or {},
        deck=ArtScene(slides=[slide]),
        sources=frozenset(sources or {"measurement"}),
    )


def _rule(
    rule_id,
    dimension="hierarchy",
    experimental=False,
    findings=None,
    covered=0,
    eligible=0,
    warnings=None,
    reliability=None,
):
    def run(slide, ctx):
        return RuleEvaluation(
            findings=findings or [],
            covered_count=covered,
            eligible_count=eligible,
            warnings=warnings or [],
            reliability=reliability,
        )

    return RuleSpec(rule_id=rule_id, dimension=dimension, run=run, experimental=experimental)


def test_rule_spec_binds_rule_id():
    rs = RuleSpec(rule_id="art.color.low_contrast", dimension="color", run=lambda s, c: None)
    assert rs.rule_id == "art.color.low_contrast"
    assert rs.dimension == "color"


def test_make_finding_sets_slide_index():
    f = make_finding("a.h", "hierarchy", Severity.MID, "m", 0.6, slide_index=1)
    assert f.slide_index == 1
    assert f.dimension == "hierarchy"


def test_grade_from_findings_none():
    assert grade_from_findings([]) == "excellent"


def test_grade_confidence_floor():
    # 契约：finding_confidence < 0.35 不驱动降级（低置信发现不进 penalty 求和）
    f = make_finding("a.h", "hierarchy", Severity.LOW, "m", 0.2, slide_index=1)
    assert grade_from_findings([f]) == "excellent"


def test_grade_high_finding_but_confidence_independent():
    # HIGH finding → grade=poor；但 finding_confidence 0.95 依然高（质量≠置信度）
    f = make_finding("a.h", "hierarchy", Severity.HIGH, "m", 0.95, slide_index=1)
    assert grade_from_findings([f]) == "poor"
    assert f.confidence == 0.95


def test_assess_dimension_not_applicable_when_no_rules():
    slide = ArtSlide(index=1, width=1920, height=1080)
    d = assess_dimension("hierarchy", [], _ctx(slide))
    assert d.status == "not_applicable"
    assert d.grade is None


def test_assess_dimension_insufficient_evidence_low_coverage():
    slide = ArtSlide(index=1, width=1920, height=1080)
    specs = [_rule(RULE_TITLE_TOO_SMALL, covered=2, eligible=10)]
    d = assess_dimension("hierarchy", specs, _ctx(slide))
    assert d.status == "insufficient_evidence"
    assert d.grade is None
    assert abs(d.evidence_coverage - 0.2) < 1e-9


def test_assess_dimension_respects_disabled_rules():
    slide = ArtSlide(index=1, width=1920, height=1080)
    from offipy.art.profiles import RULE_OFF_BALANCE, ArtProfile

    prof = ArtProfile(name="x", disabled_rules=frozenset({RULE_OFF_BALANCE}))
    ctx = RuleContext(
        profile=prof,
        slide=slide,
        slide_index=1,
        features={},
        deck=ArtScene(slides=[slide]),
        sources=frozenset({"measurement"}),
    )
    specs = [_rule(RULE_OFF_BALANCE, dimension="composition", covered=1, eligible=1)]
    d = assess_dimension("composition", specs, ctx)
    assert d.status == "not_applicable"


def test_assess_dimension_inactive_when_not_in_enabled_rules():
    # P0 门禁回归：rule_id 不在 enabled_rules（默认 ALL_RULES）→ 维度 not_applicable
    slide = ArtSlide(index=1, width=1920, height=1080)
    specs = [_rule("synthetic.not_in_registry", covered=1, eligible=1)]
    d = assess_dimension("hierarchy", specs, _ctx(slide))
    assert d.status == "not_applicable"


def test_assess_dimension_confidence_override_and_grade():
    slide = ArtSlide(index=1, width=1920, height=1080)
    from offipy.art.profiles import ArtProfile

    prof = ArtProfile(name="x", confidence_overrides={RULE_TITLE_TOO_SMALL: 0.1})
    ctx = RuleContext(
        profile=prof,
        slide=slide,
        slide_index=1,
        features={},
        deck=ArtScene(slides=[slide]),
        sources=frozenset({"measurement"}),
    )
    f = make_finding(RULE_TITLE_TOO_SMALL, "hierarchy", Severity.MID, "m", 0.6, slide_index=1)
    specs = [_rule(RULE_TITLE_TOO_SMALL, findings=[f], covered=5, eligible=5)]
    d = assess_dimension("hierarchy", specs, ctx)
    assert d.status == "assessed"
    # override 后 confidence=0.1 → 低于 floor，grade=excellent
    assert d.grade == "excellent"


def test_assess_dimension_experimental_caps_confidence():
    slide = ArtSlide(index=1, width=1920, height=1080)
    f = make_finding(
        "art.composition.off_balance", "composition", Severity.MID, "m", 0.8, slide_index=1
    )
    specs = [
        _rule(
            "art.composition.off_balance",
            dimension="composition",
            experimental=True,
            findings=[f],
            covered=3,
            eligible=3,
        )
    ]
    d = assess_dimension("composition", specs, _ctx(slide))
    assert d.status == "assessed"
    assert d.findings[0].confidence <= 0.3  # experimental cap


def test_dimension_confidence_reliability_pptx_only():
    slide = ArtSlide(index=1, width=1920, height=1080)
    specs = [_rule(RULE_TITLE_TOO_SMALL, covered=5, eligible=5)]
    # 仅 pptx 源 → reliability 0.6
    ctx = _ctx(slide, sources={"pptx"})
    d = assess_dimension("hierarchy", specs, ctx)
    assert d.status == "assessed"
    expected = 1.0 * 1.0 * 0.6  # coverage 1.0 × applicability 1.0 × reliability 0.6
    assert abs(d.confidence - expected) < 1e-9


def test_make_finding_evidence_fields():
    f = make_finding(
        "a.h",
        "hierarchy",
        Severity.MID,
        "m",
        0.6,
        slide_index=1,
        evidence_sources={"pixel"},
        evidence_reliability=0.85,
        evidence_method="declared_verified",
    )
    assert f.evidence_sources == frozenset({"pixel"})
    assert f.evidence_reliability == 0.85
    assert f.evidence_method == "declared_verified"


def test_dimension_reliability_weighted_mean_skips_experimental():
    slide = ArtSlide(index=1, width=1920, height=1080)
    specs = [
        _rule("art.hierarchy.no_focus", covered=1, eligible=1, reliability=1.0),
        _rule(
            "art.composition.off_balance",
            experimental=True,
            covered=4,
            eligible=4,
            reliability=0.3,
        ),
    ]
    d = assess_dimension("hierarchy", specs, _ctx(slide))
    assert d.status == "assessed"
    # experimental 不参与聚合 → 仅 deterministic 规则
    assert abs(d.reliability - 1.0) < 1e-9
    assert abs(d.minimum_reliability - 1.0) < 1e-9


def test_dimension_reliability_weighted_mean_mixed():
    slide = ArtSlide(index=1, width=1920, height=1080)
    specs = [
        _rule("art.typography.many_families", covered=4, eligible=4, reliability=1.0),
        _rule("art.typography.tiny_text", covered=1, eligible=1, reliability=0.5),
    ]
    d = assess_dimension("typography", specs, _ctx(slide))
    expected = (1.0 * 4 + 0.5 * 1) / 5  # 0.9
    assert abs(d.reliability - expected) < 1e-9
    assert abs(d.minimum_reliability - 0.5) < 1e-9


def test_dimension_reliability_zero_coverage_rule_excluded():
    slide = ArtSlide(index=1, width=1920, height=1080)
    specs = [
        _rule("art.typography.many_families", covered=5, eligible=5, reliability=0.8),
        _rule("art.typography.tiny_text", covered=0, eligible=5, reliability=0.2),
    ]
    d = assess_dimension("typography", specs, _ctx(slide))
    assert abs(d.reliability - 0.8) < 1e-9


def test_dimension_reliability_fallback_when_no_rule_reliability():
    slide = ArtSlide(index=1, width=1920, height=1080)
    specs = [_rule("art.typography.tiny_text", covered=5, eligible=5)]
    ctx = _ctx(slide, sources={"pptx"})
    d = assess_dimension("typography", specs, ctx)
    assert abs(d.reliability - 0.6) < 1e-9  # 回退 _scene_reliability({pptx})
    assert abs(d.minimum_reliability - 0.6) < 1e-9
