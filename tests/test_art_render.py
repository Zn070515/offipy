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
