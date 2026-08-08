from offipy.art.models import (
    ART_REPORT_SCHEMA_VERSION,
    ART_SCHEMA_VERSION,
    ArtColor,
    ArtElement,
    ArtElementRef,
    ArtFinding,
    ArtReport,
    ArtScene,
    ArtTextRun,
    DimensionAssessment,
)
from offipy.audit import Severity


def _el(**kw):
    defaults = dict(
        element_id="a", kind="text", role="body", x=0.1, y=0.1, width=0.2, height=0.1, slide_index=1
    )
    defaults.update(kw)
    return ArtElement(**defaults)


def test_schema_versions():
    assert ART_SCHEMA_VERSION == "0.2"
    assert ART_REPORT_SCHEMA_VERSION == "0.3"


def test_alpha_composite():
    top = ArtColor(0, 0, 0, a=1.0)
    bottom = ArtColor(255, 255, 255, a=1.0)
    out = top.with_alpha_over(bottom)
    assert (out.r, out.g, out.b) == (0, 0, 0)

    # 半透明黑盖白 → 128
    top = ArtColor(0, 0, 0, a=0.5)
    out = top.with_alpha_over(bottom)
    assert (out.r, out.g, out.b) == (128, 128, 128)
    assert abs(out.a - 1.0) < 1e-9


def test_color_from_dict_rgb_string():
    c = ArtColor.from_dict({"r": 30, "g": 60, "b": 90})
    assert (c.r, c.g, c.b) == (30, 60, 90)


def test_color_from_dict_malformed_returns_none():
    # 未受信 color dict：r/g/b 缺失或非数字 → None（无颜色证据），不抛 ValueError
    assert ArtColor.from_dict({"r": "x", "g": 0, "b": 0}) is None
    assert ArtColor.from_dict({"r": 0, "g": 0}) is None  # 缺 b
    assert ArtColor.from_dict({"r": None, "g": 0, "b": 0}) is None


def test_coordinate_overflow_allowed():
    el = _el(x=-0.05, width=1.3)
    assert el.x == -0.05
    assert el.area == 1.3 * 0.1


def test_font_size_norm_is_plain_field():
    # 契约：font_size_norm 是适配器填的普通字段，模型不自动算
    el = _el(font_size=48.0, font_size_unit="px")
    assert el.font_size_norm is None
    el2 = _el(font_size=48.0, font_size_unit="px", font_size_norm=48.0 / 1080.0)
    assert abs(el2.font_size_norm - 48.0 / 1080.0) < 1e-9


def test_colors_three_way_split():
    fg = ArtColor(0, 0, 0)
    bg = ArtColor(255, 255, 255)
    bd = ArtColor(200, 200, 200)
    el = _el(foreground=fg, background=bg, border=bd, is_background=False)
    assert el.foreground == fg
    assert el.background == bg
    assert el.border == bd
    assert el.is_background is False


def test_run_has_font_family():
    r = ArtTextRun(text="x", font_family="Microsoft YaHei")
    assert r.font_family == "Microsoft YaHei"


def test_element_ref_frozen():
    r = ArtElementRef(slide_index=1, element_id="a", kind="text", role="title")
    try:
        r.element_id = "b"
    except Exception:
        pass
    else:
        raise AssertionError("ArtElementRef must be frozen")


def test_finding_slide_index_optional_and_serialize():
    primary = ArtElementRef(1, "t", "text", "title")
    related = [ArtElementRef(1, "b", "text", "body")]
    f = ArtFinding(
        rule_id="a.h",
        dimension="hierarchy",
        severity=Severity.MID,
        message="m",
        confidence=0.6,
        slide_index=1,
        primary=primary,
        related=related,
    )
    d = f.to_dict()
    assert d["primary"]["element_id"] == "t"
    assert d["related"][0]["role"] == "body"
    deck_f = ArtFinding(
        rule_id="d.r", dimension="consistency", severity=Severity.LOW, message="m", confidence=0.5
    )
    assert deck_f.slide_index is None  # deck 级 finding 可无 slide


def test_dimension_assessment_requires_status():
    d = DimensionAssessment(
        dimension="hierarchy",
        status="assessed",
        grade="good",
        confidence=0.8,
        evidence_coverage=0.9,
    )
    assert d.grade == "good"
    assert d.status == "assessed"


def test_scene_warnings_and_sources():
    from offipy.art.models import ArtWarning

    scene = ArtScene(
        slides=[], width_unit="px", warnings=[ArtWarning("w", "m")], sources={"measurement"}
    )
    assert scene.sources == {"measurement"}
    assert scene.warnings[0].code == "w"


def test_scene_from_dict_hand_written():
    data = {
        "slides": [
            {
                "index": 1,
                "width": 1920.0,
                "height": 1080.0,
                "background_color": {"r": 255, "g": 255, "b": 255, "a": 1.0},
                "elements": [
                    {
                        "element_id": "t",
                        "kind": "text",
                        "role": "title",
                        "x": 0.1,
                        "y": 0.05,
                        "width": 0.6,
                        "height": 0.1,
                        "font_size": 48.0,
                        "font_size_unit": "px",
                        "text": "Hi",
                        "foreground": {"r": 0, "g": 0, "b": 0, "a": 1.0},
                    }
                ],
            }
        ],
    }
    scene = ArtScene.from_dict(data)
    assert scene.slides[0].index == 1
    el = scene.slides[0].elements[0]
    assert el.kind == "text" and el.role == "title"
    assert el.foreground == ArtColor(0, 0, 0)
    assert abs(el.font_size_norm - 48.0 / 1080.0) < 1e-9


def test_report_experimental_score_default_none():
    r = ArtReport(profile="balanced")
    assert r.experimental_score is None


def test_report_to_dict_schema():
    r = ArtReport(profile="balanced", experimental_score=66.7)
    d = r.to_dict()
    assert d["schema_version"] == ART_REPORT_SCHEMA_VERSION
    assert d["profile"] == "balanced"
    assert d["experimental_score"] == 66.7


def test_element_pixel_evidence_round_trip():
    from offipy.art.models import ElementPixelEvidence

    pe = ElementPixelEvidence(
        foreground=ArtColor(30, 30, 30),
        foreground_match_ratio=0.8,
        background_complexity=0.1,
        color_confidence=0.85,
        method="declared_verified",
    )
    d = pe.to_dict()
    assert d["method"] == "declared_verified"
    assert d["foreground"]["r"] == 30
    assert ElementPixelEvidence.from_dict(d) == pe


def test_slide_pixel_evidence_round_trip():
    from offipy.art.models import PixelColorShare, SlidePixelEvidence

    sp = SlidePixelEvidence(
        background=ArtColor(255, 255, 255),
        background_confidence=0.9,
        background_uniformity=0.95,
        palette=[PixelColorShare(ArtColor(255, 255, 255), 0.8)],
        background_like_ratio=0.2,
    )
    d = sp.to_dict()
    assert d["palette"][0]["ratio"] == 0.8
    assert SlidePixelEvidence.from_dict(d) == sp


def test_finding_evidence_serialization():
    f = ArtFinding(
        rule_id="art.color.low_contrast",
        dimension="color",
        severity=Severity.MID,
        message="m",
        confidence=0.85,
        slide_index=1,
        evidence_sources=frozenset({"pixel", "measurement"}),
        evidence_reliability=0.85,
        evidence_method="declared_verified",
    )
    d = f.to_dict()
    assert d["evidence_sources"] == ["measurement", "pixel"]
    assert d["evidence_reliability"] == 0.85
    assert d["evidence_method"] == "declared_verified"


def test_finding_default_omits_severity_override():
    f = ArtFinding(
        rule_id="a.h",
        dimension="hierarchy",
        severity=Severity.MID,
        message="m",
        confidence=0.6,
    )
    d = f.to_dict()
    assert "severity_override" not in d
    assert "severity_override_source" not in d
    # 0.2 报告字典可被旧解析路径读取：默认 finding 序列化须与 0.2 逐字节一致
    assert d == {
        "rule_id": "a.h",
        "dimension": "hierarchy",
        "severity": "MID",
        "message": "m",
        "confidence": 0.6,
        "slide_index": None,
        "primary": None,
        "related": [],
        "details": {},
        "evidence_sources": [],
        "evidence_reliability": None,
        "evidence_method": None,
    }


def test_finding_severity_override_serialized():
    for src in ("user", "feedback"):
        f = ArtFinding(
            rule_id="a.h",
            dimension="hierarchy",
            severity=Severity.LOW,
            message="m",
            confidence=0.6,
            severity_override=True,
            severity_override_source=src,
        )
        d = f.to_dict()
        assert d["severity_override"] is True
        assert d["severity_override_source"] == src


def test_finding_severity_override_source_omitted_when_none():
    # override=True 但无来源时，source 仍按契约省略
    f = ArtFinding(
        rule_id="a.h",
        dimension="hierarchy",
        severity=Severity.MID,
        message="m",
        confidence=0.6,
        severity_override=True,
    )
    d = f.to_dict()
    assert d["severity_override"] is True
    assert "severity_override_source" not in d


def test_dimension_reliability_serialization():
    d = DimensionAssessment(
        dimension="hierarchy",
        status="assessed",
        grade="good",
        confidence=0.8,
        evidence_coverage=0.9,
        reliability=0.85,
        minimum_reliability=0.55,
    )
    dd = d.to_dict()
    assert dd["reliability"] == 0.85
    assert dd["minimum_reliability"] == 0.55
    # None 不序列化（0.1 报告可被 0.2 读取）
    d2 = DimensionAssessment(dimension="hierarchy", status="assessed")
    assert "reliability" not in d2.to_dict()


def test_scene_from_dict_parses_pixel_evidence_and_metadata():
    data = {
        "width_unit": "px",
        "metadata": {"pixel_pages_covered": 1, "pixel_pages_total": 1},
        "slides": [
            {
                "index": 1,
                "width": 1920.0,
                "height": 1080.0,
                "pixel_evidence": {"background_confidence": 0.9, "method": "unsupported"},
                "elements": [
                    {
                        "element_id": "t",
                        "kind": "text",
                        "role": "title",
                        "x": 0.1,
                        "y": 0.05,
                        "width": 0.6,
                        "height": 0.1,
                        "pixel_evidence": {
                            "method": "declared_verified",
                            "color_confidence": 0.8,
                        },
                    }
                ],
            }
        ],
    }
    scene = ArtScene.from_dict(data)
    assert scene.metadata["pixel_pages_covered"] == 1
    assert scene.slides[0].pixel_evidence.background_confidence == 0.9
    assert scene.slides[0].elements[0].pixel_evidence.method == "declared_verified"
