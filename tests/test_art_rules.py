from dataclasses import replace

from art_helpers import make_slide as _hs
from art_helpers import make_text_element as _hte
from offipy.art.color import RULES as COLOR_RULES
from offipy.art.features import compute_features
from offipy.art.models import ArtColor, ArtScene, ArtSlide
from offipy.art.profiles import (
    RULE_LOW_CONTRAST,
    RULE_TITLE_TOO_SMALL,
    ArtProfile,
    get_profile,
)
from offipy.art.rules import (
    RuleContext,
    RuleEvaluation,
    RuleSpec,
    _step_severity,
    apply_profile_to_finding,
    assess_dimension,
    grade_from_findings,
    make_finding,
)
from offipy.audit import Severity


def _ctx(slide, profile="balanced", features=None, sources=None):
    prof = profile if isinstance(profile, ArtProfile) else get_profile(profile)
    return RuleContext(
        profile=prof,
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


def test_assess_dimension_keeps_well_covered_finding_when_aggregate_low():
    # #155：一条规则高覆盖（含 finding）+ 一条低覆盖 → 高覆盖 finding 必须保留
    slide = ArtSlide(index=1, width=1920, height=1080)
    f = make_finding(RULE_TITLE_TOO_SMALL, "hierarchy", Severity.MID, "m", 0.6, slide_index=1)
    specs = [
        _rule("art.hierarchy.no_focus", covered=1, eligible=1, findings=[f]),
        _rule("art.typography.many_families", covered=1, eligible=10),
    ]
    d = assess_dimension("hierarchy", specs, _ctx(slide))
    assert d.status == "assessed"  # 有 ≥1 条 assessable 规则
    assert d.findings == [f]
    assert any(w.code == "art.rule.insufficient_coverage" for w in d.warnings)


def test_assess_dimension_gated_finding_dropped_when_rule_under_covered():
    slide = ArtSlide(index=1, width=1920, height=1080)
    f = make_finding(RULE_TITLE_TOO_SMALL, "hierarchy", Severity.MID, "m", 0.6, slide_index=1)
    specs = [_rule(RULE_TITLE_TOO_SMALL, covered=1, eligible=10, findings=[f])]
    d = assess_dimension("hierarchy", specs, _ctx(slide))
    assert d.status == "insufficient_evidence"
    assert d.findings == []


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


def test_dimension_reliability_fallback_measurement_source():
    slide = ArtSlide(index=1, width=1920, height=1080)
    specs = [_rule("art.typography.tiny_text", covered=5, eligible=5)]
    ctx = _ctx(slide)  # default sources={"measurement"} per _ctx helper
    d = assess_dimension("typography", specs, ctx)
    assert abs(d.reliability - 1.0) < 1e-9


def test_dimension_reliability_fallback_pixel_source():
    slide = ArtSlide(index=1, width=1920, height=1080)
    specs = [_rule("art.typography.tiny_text", covered=5, eligible=5)]
    ctx = _ctx(slide, sources={"pixel"})
    d = assess_dimension("typography", specs, ctx)
    assert abs(d.reliability - 0.55) < 1e-9


def test_dimension_reliability_excludes_profile_experimental():
    slide = ArtSlide(index=1, width=1920, height=1080)
    spec = _rule("art.typography.many_families", covered=5, eligible=5, reliability=0.9)
    profile = ArtProfile(name="x", experimental_rules={"art.typography.many_families"})
    ctx = _ctx(slide, profile=profile)
    d = assess_dimension("typography", [spec], ctx)
    # profile-experimental 规则同样不参与聚合 → 无权重 → 回退场景可靠度
    assert abs(d.reliability - 1.0) < 1e-9  # sources 默认 measurement


# ---------------------------------------------------------------------------
# #153：规则无 ev.reliability 时，从 finding.evidence_reliability 取 min 派生
# ---------------------------------------------------------------------------


def test_dimension_reliability_derived_from_finding_evidence():
    slide = ArtSlide(index=1, width=1920, height=1080)
    f = make_finding(
        RULE_LOW_CONTRAST,
        "color",
        Severity.LOW,
        "m",
        0.25,
        slide_index=1,
        evidence_reliability=0.5,
    )
    specs = [_rule(RULE_LOW_CONTRAST, dimension="color", findings=[f], covered=3, eligible=3)]
    d = assess_dimension("color", specs, _ctx(slide))
    assert d.status == "assessed"
    assert abs(d.reliability - 0.5) < 1e-9
    assert abs(d.minimum_reliability - 0.5) < 1e-9


def test_dimension_reliability_derived_min_of_findings():
    slide = ArtSlide(index=1, width=1920, height=1080)
    f1 = make_finding(
        RULE_LOW_CONTRAST,
        "color",
        Severity.LOW,
        "m",
        0.25,
        slide_index=1,
        evidence_reliability=0.5,
    )
    f2 = make_finding(
        RULE_LOW_CONTRAST,
        "color",
        Severity.MID,
        "m",
        0.9,
        slide_index=1,
        evidence_reliability=0.8,
    )
    specs = [_rule(RULE_LOW_CONTRAST, dimension="color", findings=[f1, f2], covered=3, eligible=3)]
    d = assess_dimension("color", specs, _ctx(slide))
    assert abs(d.reliability - 0.5) < 1e-9  # min(0.5, 0.8)
    assert abs(d.minimum_reliability - 0.5) < 1e-9


# ---------------------------------------------------------------------------
# _step_severity 边界
# ---------------------------------------------------------------------------


def test_step_severity_low_up():
    assert _step_severity(Severity.LOW, 1) == Severity.MID


def test_step_severity_mid_up():
    assert _step_severity(Severity.MID, 1) == Severity.HIGH


def test_step_severity_high_up_saturates():
    assert _step_severity(Severity.HIGH, 1) == Severity.HIGH


def test_step_severity_high_down():
    assert _step_severity(Severity.HIGH, -1) == Severity.MID


def test_step_severity_mid_down():
    assert _step_severity(Severity.MID, -1) == Severity.LOW


def test_step_severity_low_down_saturates():
    assert _step_severity(Severity.LOW, -1) == Severity.LOW


# ---------------------------------------------------------------------------
# feedback_severity_adjustments via assess_dimension
# ---------------------------------------------------------------------------


def test_feedback_delta_applied_via_assess_dimension():
    slide = ArtSlide(index=1, width=1920, height=1080)
    f = make_finding(RULE_TITLE_TOO_SMALL, "hierarchy", Severity.MID, "m", 0.7, slide_index=1)
    specs = [_rule(RULE_TITLE_TOO_SMALL, findings=[f], covered=5, eligible=5)]
    prof = ArtProfile(name="x", feedback_severity_adjustments={RULE_TITLE_TOO_SMALL: +1})
    ctx = _ctx(slide, profile=prof)
    d = assess_dimension("hierarchy", specs, ctx)
    assert d.status == "assessed"
    assert d.findings[0].severity == Severity.HIGH
    assert d.findings[0].severity_override is True
    assert d.findings[0].severity_override_source == "feedback"


def test_feedback_delta_high_up_no_false_provenance():
    slide = ArtSlide(index=1, width=1920, height=1080)
    f = make_finding(RULE_TITLE_TOO_SMALL, "hierarchy", Severity.HIGH, "m", 0.7, slide_index=1)
    specs = [_rule(RULE_TITLE_TOO_SMALL, findings=[f], covered=5, eligible=5)]
    prof = ArtProfile(name="x", feedback_severity_adjustments={RULE_TITLE_TOO_SMALL: +1})
    ctx = _ctx(slide, profile=prof)
    d = assess_dimension("hierarchy", specs, ctx)
    assert d.status == "assessed"
    assert d.findings[0].severity == Severity.HIGH
    assert d.findings[0].severity_override is False
    assert d.findings[0].severity_override_source is None


def test_user_override_beats_feedback():
    slide = ArtSlide(index=1, width=1920, height=1080)
    f = make_finding(RULE_TITLE_TOO_SMALL, "hierarchy", Severity.MID, "m", 0.7, slide_index=1)
    specs = [_rule(RULE_TITLE_TOO_SMALL, findings=[f], covered=5, eligible=5)]
    prof = ArtProfile(
        name="x",
        severity_overrides={RULE_TITLE_TOO_SMALL: Severity.HIGH},
        feedback_severity_adjustments={RULE_TITLE_TOO_SMALL: -1},
    )
    ctx = _ctx(slide, profile=prof)
    d = assess_dimension("hierarchy", specs, ctx)
    assert d.status == "assessed"
    assert d.findings[0].severity == Severity.HIGH
    assert d.findings[0].severity_override is True
    assert d.findings[0].severity_override_source == "user"


# ---------------------------------------------------------------------------
# feedback delta on dynamically-computed severity (color rules)
# ---------------------------------------------------------------------------


def _color_ctx(slide, profile="balanced"):
    prof = profile if isinstance(profile, ArtProfile) else get_profile(profile)
    return RuleContext(
        profile=prof,
        slide=slide,
        slide_index=slide.index,
        features=compute_features(slide),
        deck=ArtScene(slides=[slide]),
        sources=frozenset({"measurement"}),
    )


def test_feedback_delta_color_high_to_mid():
    """contrast ratio < 2.0 → HIGH, with -1 delta → MID, source=feedback."""
    slide = _hs(
        1,
        elements=[
            _hte("t", "Gray on white", font_size=24.0, foreground=ArtColor(200, 200, 200)),
        ],
        background_color=ArtColor(255, 255, 255),
    )
    prof = replace(
        get_profile("balanced"),
        feedback_severity_adjustments={RULE_LOW_CONTRAST: -1},
    )
    d = assess_dimension("color", COLOR_RULES, _color_ctx(slide, profile=prof))
    assert d.status == "assessed"
    findings = [f for f in d.findings if f.rule_id == RULE_LOW_CONTRAST]
    assert len(findings) == 1
    f = findings[0]
    assert f.severity == Severity.MID
    assert f.severity_override is True
    assert f.severity_override_source == "feedback"


def test_feedback_delta_color_mid_to_low():
    """contrast ratio in (2.0, 3.0) → MID, with -1 delta → LOW, source=feedback."""
    slide = _hs(
        1,
        elements=[
            _hte("t", "Gray on white", font_size=24.0, foreground=ArtColor(160, 160, 160)),
        ],
        background_color=ArtColor(255, 255, 255),
    )
    prof = replace(
        get_profile("balanced"),
        feedback_severity_adjustments={RULE_LOW_CONTRAST: -1},
    )
    d = assess_dimension("color", COLOR_RULES, _color_ctx(slide, profile=prof))
    assert d.status == "assessed"
    findings = [f for f in d.findings if f.rule_id == RULE_LOW_CONTRAST]
    assert len(findings) == 1
    f = findings[0]
    assert f.severity == Severity.LOW
    assert f.severity_override is True
    assert f.severity_override_source == "feedback"


# ---------------------------------------------------------------------------
# experimental confidence cap + feedback delta coexist
# ---------------------------------------------------------------------------


def test_experimental_confidence_cap_with_feedback_delta():
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
    prof = ArtProfile(
        name="x",
        feedback_severity_adjustments={"art.composition.off_balance": +1},
    )
    ctx = _ctx(slide, profile=prof)
    d = assess_dimension("composition", specs, ctx)
    assert d.status == "assessed"
    found = d.findings[0]
    # severity bumped by feedback
    assert found.severity == Severity.HIGH
    assert found.severity_override is True
    assert found.severity_override_source == "feedback"
    # confidence still capped by experimental
    assert found.confidence <= 0.3


# ---------------------------------------------------------------------------
# apply_profile_to_finding direct (unit)
# ---------------------------------------------------------------------------


def test_apply_profile_to_finding_feedback_delta():
    f = make_finding("art.color.low_contrast", "color", Severity.LOW, "m", 0.6, slide_index=1)
    prof = ArtProfile(
        name="x",
        feedback_severity_adjustments={"art.color.low_contrast": +1},
    )
    result = apply_profile_to_finding(f, prof)
    assert result.severity == Severity.MID
    assert result.severity_override is True
    assert result.severity_override_source == "feedback"


def test_apply_profile_to_finding_user_override():
    f = make_finding("art.color.low_contrast", "color", Severity.LOW, "m", 0.6, slide_index=1)
    prof = ArtProfile(
        name="x",
        severity_overrides={"art.color.low_contrast": Severity.HIGH},
        feedback_severity_adjustments={"art.color.low_contrast": -1},
    )
    result = apply_profile_to_finding(f, prof)
    assert result.severity == Severity.HIGH
    assert result.severity_override is True
    assert result.severity_override_source == "user"


def test_apply_profile_to_finding_confidence_override():
    f = make_finding("art.color.low_contrast", "color", Severity.LOW, "m", 0.6, slide_index=1)
    prof = ArtProfile(name="x", confidence_overrides={"art.color.low_contrast": 0.9})
    result = apply_profile_to_finding(f, prof)
    assert result.confidence == 0.9
