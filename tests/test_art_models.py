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
    assert ART_SCHEMA_VERSION == "0.1"
    assert ART_REPORT_SCHEMA_VERSION == "0.1"


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
