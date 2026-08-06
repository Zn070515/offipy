from offipy.art.compare import compare_reports
from offipy.art.models import (
    ArtElementRef,
    ArtFinding,
    ArtReport,
    ArtSlideReport,
    DimensionAssessment,
)
from offipy.art.rules import grade_from_findings
from offipy.audit import Severity


def _finding(
    rule_id,
    slide_index,
    sev=Severity.MID,
    dim="hierarchy",
    primary_id="a",
    related=(),
    confidence=0.6,
):
    return ArtFinding(
        rule_id=rule_id,
        dimension=dim,
        severity=sev,
        message=rule_id,
        confidence=confidence,
        slide_index=slide_index,
        primary=ArtElementRef(slide_index, primary_id, "text", "body"),
        related=[ArtElementRef(slide_index, r, "text", "body") for r in related],
    )


def _report(slide_dims, deck=None):
    slides = []
    for i in sorted(slide_dims):
        dims = []
        for dim, fs in slide_dims[i].items():
            grade = grade_from_findings(fs)  # rev2.1：只返回 grade
            dims.append(
                DimensionAssessment(dimension=dim, grade=grade, confidence=0.8, findings=fs)
            )
        slides.append(ArtSlideReport(slide_index=i, dimensions=dims))
    return ArtReport(slides=slides, deck_findings=deck or [])


def test_new_and_unchanged():
    before = _report({1: {"hierarchy": [_finding("a.h", 1)]}})
    after = _report({1: {"hierarchy": [_finding("a.h", 1), _finding("b.h", 1)]}})
    diff = compare_reports(before, after)
    assert [c.rule_id for c in diff.new_findings] == ["b.h"]
    assert diff.resolved_findings == []
    assert any(c.rule_id == "a.h" and c.status == "unchanged" for c in diff.changes)


def test_resolved():
    before = _report({1: {"hierarchy": [_finding("a.h", 1), _finding("b.h", 1)]}})
    after = _report({1: {"hierarchy": [_finding("a.h", 1)]}})
    diff = compare_reports(before, after)
    assert [c.rule_id for c in diff.resolved_findings] == ["b.h"]


def test_occurrence_distinguishes_same_rule_different_related():
    a = _finding("a.h", 1, related=("x",))
    b = _finding("a.h", 1, related=("y",))
    assert a.rule_id == b.rule_id and a.slide_index == b.slide_index
    from offipy.art.compare import _stable_key

    assert _stable_key(a) != _stable_key(b)


def test_changed_detected():
    before = _report({1: {"hierarchy": [_finding("a.h", 1, confidence=0.6)]}})
    after = _report({1: {"hierarchy": [_finding("a.h", 1, confidence=0.6)]}})
    after.slides[0].dimensions[0].findings[0].details = {"x": 1}
    diff = compare_reports(before, after)
    assert any(c.rule_id == "a.h" and c.status == "changed" for c in diff.changes)


def test_improved_detected():
    before = _report({1: {"hierarchy": [_finding("a.h", 1, sev=Severity.HIGH)]}})
    after = _report({1: {"hierarchy": [_finding("a.h", 1, sev=Severity.MID)]}})
    diff = compare_reports(before, after)
    assert any(c.rule_id == "a.h" and c.status == "improved" for c in diff.changes)


def test_worsened_detected():
    before = _report({1: {"hierarchy": [_finding("a.h", 1, sev=Severity.MID)]}})
    after = _report({1: {"hierarchy": [_finding("a.h", 1, sev=Severity.HIGH)]}})
    diff = compare_reports(before, after)
    assert any(c.rule_id == "a.h" and c.status == "worsened" for c in diff.changes)


def test_deck_finding_no_primary_matches():
    # deck 级 finding（无 primary、slide_index=None）按 message+details 匹配，不误报 resolved
    a = ArtFinding(
        rule_id="art.color.no_accent",
        dimension="color",
        severity=Severity.LOW,
        message="no accent",
        confidence=0.3,
        slide_index=None,
        primary=None,
    )
    diff = compare_reports(
        ArtReport(slides=[], deck_findings=[a]), ArtReport(slides=[], deck_findings=[a])
    )
    assert any(c.rule_id == "art.color.no_accent" and c.status == "unchanged" for c in diff.changes)
    assert diff.resolved_findings == []


def test_grade_change_detected():
    before = _report({1: {"hierarchy": [_finding("a.h", 1)]}})
    after = _report({1: {"hierarchy": [_finding("a.h", 1), _finding("b.h", 1)]}})
    diff = compare_reports(before, after)
    assert any(
        g.dimension == "hierarchy" and g.slide_index == 1 and g.before != g.after
        for g in diff.grade_changes
    )


def test_schema_mismatch_warns():
    before = _report({1: {}})
    before.schema_version = "0.1"
    after = _report({1: {}})
    after.schema_version = "0.2"
    diff = compare_reports(before, after)
    assert any(w.code == "art.compare.schema_mismatch" for w in diff.warnings)


def test_profile_mismatch_warns():
    before = _report({1: {}})
    before.profile = "balanced"
    after = _report({1: {}})
    after.profile = "consulting"
    diff = compare_reports(before, after)
    assert any(w.code == "art.compare.profile_mismatch" for w in diff.warnings)
