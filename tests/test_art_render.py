import json

from offipy.art.models import (
    ArtElementRef,
    ArtFinding,
    ArtReport,
    ArtSlideReport,
    DimensionAssessment,
)
from offipy.art.render import render_html, render_markdown, report_to_json
from offipy.audit import Severity


def _finding():
    return ArtFinding(
        rule_id="a.h",
        dimension="hierarchy",
        severity=Severity.MID,
        message="m",
        confidence=0.6,
        slide_index=1,
        primary=ArtElementRef(1, "t", "text", "title"),
    )


def _report():
    # 3 个 assessed 维度：保证 _radar_svg(n>=3) 真正渲染雷达（见 plan-bug 修正说明）
    return ArtReport(
        profile="balanced",
        experimental_score=66.7,
        slides=[
            ArtSlideReport(
                slide_index=1,
                dimensions=[
                    DimensionAssessment(
                        dimension="hierarchy",
                        status="assessed",
                        grade="good",
                        confidence=0.8,
                        findings=[_finding()],
                    ),
                    DimensionAssessment(
                        dimension="composition",
                        status="assessed",
                        grade="good",
                        confidence=0.7,
                        findings=[_finding()],
                    ),
                    DimensionAssessment(
                        dimension="typography",
                        status="assessed",
                        grade="good",
                        confidence=0.6,
                        findings=[_finding()],
                    ),
                ],
            )
        ],
    )


def test_report_to_json_indent_and_ascii():
    d = report_to_json(_report())
    json.dumps(d, ensure_ascii=False)
    assert d["profile"] == "balanced"
    assert d["experimental_score"] == 66.7
    assert d["slides"][0]["slide_index"] == 1  # 1-based
    assert d["slides"][0]["dimensions"][0]["status"] == "assessed"
    assert d["slides"][0]["dimensions"][0]["findings"][0]["rule_id"] == "a.h"


def test_render_markdown_one_based_slide():
    md = render_markdown(_report())
    assert "## Slide 1" in md
    assert "a.h" in md
    assert "hierarchy" in md


def test_render_markdown_status_labels():
    rep = ArtReport(
        slides=[
            ArtSlideReport(
                slide_index=1,
                dimensions=[
                    DimensionAssessment(
                        dimension="hierarchy", status="insufficient_evidence", evidence_coverage=0.3
                    ),
                ],
            )
        ],
    )
    md = render_markdown(rep)
    assert "证据不足" in md


def test_render_html_contains_dimensions_and_radar_label():
    html = render_html(_report())
    assert "hierarchy" in html
    assert "good" in html
    assert "规则评级" in html
    assert "<svg" in html


def test_render_markdown_shows_evidence_and_reliability():
    f = ArtFinding(
        rule_id="a.h",
        dimension="hierarchy",
        severity=Severity.MID,
        message="m",
        confidence=0.6,
        slide_index=1,
        evidence_sources=frozenset({"pixel"}),
        evidence_reliability=0.85,
        evidence_method="declared_verified",
    )
    d = DimensionAssessment(
        dimension="hierarchy",
        status="assessed",
        grade="good",
        confidence=0.8,
        reliability=0.85,
        findings=[f],
    )
    rep = ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])])
    md = render_markdown(rep)
    assert "Evidence: pixel" in md
    assert "Method: declared_verified" in md
    assert "Reliability: 0.85" in md
    assert "reliability 0.85" in md


def test_render_html_shows_evidence():
    f = ArtFinding(
        rule_id="a.h",
        dimension="hierarchy",
        severity=Severity.MID,
        message="m",
        confidence=0.6,
        slide_index=1,
        evidence_sources=frozenset({"pixel"}),
        evidence_reliability=0.85,
        evidence_method="declared_verified",
    )
    d = DimensionAssessment(
        dimension="hierarchy",
        status="assessed",
        grade="good",
        confidence=0.8,
        reliability=0.85,
        findings=[f],
    )
    rep = ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])])
    html = render_html(rep)
    assert "Evidence: pixel" in html
    assert "/ rel 0.85" in html


def test_render_markdown_no_evidence_no_suffix():
    f = ArtFinding(
        rule_id="a.h",
        dimension="hierarchy",
        severity=Severity.MID,
        message="m",
        confidence=0.6,
        slide_index=1,
    )
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    md = render_markdown(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "Evidence:" not in md


def test_render_html_no_evidence_no_suffix():
    f = ArtFinding(
        rule_id="a.h",
        dimension="hierarchy",
        severity=Severity.MID,
        message="m",
        confidence=0.6,
        slide_index=1,
    )
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    html = render_html(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "Evidence:" not in html


def test_render_html_escapes_profile_dimension_grade():
    # profile/dimension/grade 本应来自枚举，但 ArtReport 是纯 dataclass——
    # 敌意调用方能注入 HTML/JS。转义后原文不出现在输出里。
    d = DimensionAssessment(
        dimension="<script>dim</script>",
        status="assessed",
        grade='good">x<',
        confidence=0.8,
        findings=[_finding()],
    )
    d2 = DimensionAssessment(
        dimension="composition", status="assessed", grade="good", confidence=0.7
    )
    d3 = DimensionAssessment(
        dimension="typography", status="assessed", grade="good", confidence=0.6
    )
    rep = ArtReport(
        profile="<img src=x onerror=alert(1)>",
        slides=[ArtSlideReport(slide_index=1, dimensions=[d, d2, d3])],
    )
    html = render_html(rep)
    assert "<img src=x onerror=alert(1)>" not in html
    assert "&lt;img src=x onerror=alert(1)&gt;" in html
    assert "<script>dim</script>" not in html
    assert "&lt;script&gt;dim&lt;/script&gt;" in html
    assert 'good">x<' not in html
    assert "good&quot;&gt;x&lt;" in html


def test_render_html_escapes_evidence_method():
    f = ArtFinding(
        rule_id="a.h",
        dimension="hierarchy",
        severity=Severity.MID,
        message="m",
        confidence=0.6,
        slide_index=1,
        evidence_sources=frozenset({"pixel"}),
        evidence_reliability=0.85,
        evidence_method="<b>bad</b>",
    )
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    html = render_html(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "&lt;b&gt;bad&lt;/b&gt;" in html
    assert "<b>bad</b>" not in html


def test_render_markdown_shows_feedback_provenance():
    f = _finding()
    f.severity_override = True
    f.severity_override_source = "feedback"
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    md = render_markdown(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "Severity adjusted: feedback" in md
    assert "Severity adjusted: user override" not in md


def test_render_markdown_shows_user_override_provenance():
    f = _finding()
    f.severity_override = True
    f.severity_override_source = "user"
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    md = render_markdown(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "Severity adjusted: user override" in md


def test_render_markdown_deck_provenance():
    f = _finding()
    f.severity_override = True
    f.severity_override_source = "feedback"
    md = render_markdown(ArtReport(deck_findings=[f]))
    assert "Severity adjusted: feedback" in md


def test_render_markdown_no_provenance_for_untouched_finding():
    f = _finding()
    assert f.severity_override is False
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    md = render_markdown(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "Severity adjusted:" not in md


def test_render_html_shows_feedback_provenance():
    f = _finding()
    f.severity_override = True
    f.severity_override_source = "feedback"
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    html = render_html(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "Severity adjusted: feedback" in html


def test_render_html_deck_provenance():
    f = _finding()
    f.severity_override = True
    f.severity_override_source = "user"
    html = render_html(ArtReport(deck_findings=[f]))
    assert "Severity adjusted: user override" in html


def test_render_html_no_provenance_for_untouched_finding():
    f = _finding()
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    html = render_html(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "Severity adjusted:" not in html


def test_report_to_json_schema_version_0_3():
    from offipy.art.models import ART_REPORT_SCHEMA_VERSION

    assert report_to_json(_report())["schema_version"] == ART_REPORT_SCHEMA_VERSION


def test_report_to_json_emits_override_provenance():
    f = _finding()
    f.severity_override = True
    f.severity_override_source = "feedback"
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    rep = ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])])
    fd = report_to_json(rep)["slides"][0]["dimensions"][0]["findings"][0]
    assert fd["severity_override"] is True
    assert fd["severity_override_source"] == "feedback"


def test_render_marks_experimental_findings():
    f = _finding()
    f.experimental = True
    d = DimensionAssessment(
        dimension="hierarchy", status="assessed", grade="good", confidence=0.8, findings=[f]
    )
    md = render_markdown(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "[experimental]" in md
    html = render_html(ArtReport(slides=[ArtSlideReport(slide_index=1, dimensions=[d])]))
    assert "experimental" in html
