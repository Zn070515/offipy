import json

import pytest

from art_helpers import make_scene, make_slide, make_text_element
from offipy.art.analyze import analyze_deck, analyze_scene
from offipy.exceptions import InvalidArgumentError


def test_analyze_scene_builds_report():
    scene = make_scene(
        [
            make_slide(
                1,
                height=1080.0,
                elements=[
                    make_text_element(
                        "title", "Title", x=0.1, y=0.05, w=0.6, h=0.08, font_size=52.0, role="title"
                    ),
                    make_text_element(
                        "body", "Body", x=0.1, y=0.2, w=0.5, h=0.06, font_size=24.0, role="body"
                    ),
                    make_text_element(
                        "cap",
                        "Caption",
                        x=0.1,
                        y=0.4,
                        w=0.4,
                        h=0.05,
                        font_size=18.0,
                        role="caption",
                    ),
                ],
            ),
        ]
    )
    report = analyze_scene(scene, profile="balanced")
    assert report.profile == "balanced"
    assert len(report.slides) == 1
    s = report.slides[0]
    assert s.slide_index == 1
    assert len(s.dimensions) == 5
    for dim in ("hierarchy", "composition", "typography", "color", "media"):
        d = s.by_dimension(dim)
        assert d is not None
        assert d.status in ("assessed", "insufficient_evidence", "not_applicable")
        if d.status == "assessed":
            assert d.grade in ("excellent", "good", "attention", "poor")
    assert report.experimental_score is None  # 默认不计算


def test_analyze_scene_experimental_score_when_requested():
    scene = make_scene(
        [
            make_slide(
                1,
                height=1080.0,
                elements=[
                    make_text_element(
                        "t", "Title", x=0.1, y=0.05, w=0.6, h=0.08, font_size=52.0, role="title"
                    ),
                    make_text_element(
                        "b", "Body", x=0.1, y=0.2, w=0.5, h=0.06, font_size=24.0, role="body"
                    ),
                    make_text_element(
                        "c", "Cap", x=0.1, y=0.4, w=0.4, h=0.05, font_size=18.0, role="caption"
                    ),
                ],
            ),
        ]
    )
    report = analyze_scene(scene, include_experimental_score=True)
    assert report.experimental_score is not None


def test_analyze_deck_measurements_only(tmp_path):
    m = tmp_path / "m.json"
    m.write_text(json.dumps({"slides": []}), encoding="utf-8")
    result = analyze_deck(measurements=str(m), profile="balanced")
    assert result.geometry is None
    assert result.art is not None
    assert result.art.profile == "balanced"
    assert result.art.slides == []


def test_analyze_deck_geometry_only(tmp_path, monkeypatch):
    import offipy.art.analyze as A
    from offipy.audit.models import AuditConfig, PptxAuditReport

    def fake_audit(p):
        return PptxAuditReport(
            schema_version="0.1",
            offipy_version="0.11.6",
            path=p,
            source_sha256="abc",
            slide_size=(10.0, 7.5),
            slide_count=1,
            config=AuditConfig(),
        )

    monkeypatch.setattr(A, "audit_pptx", fake_audit)
    result = A.analyze_deck(pptx="x.pptx")
    # rev2.1 正式契约：PPTX-only 可产出艺术报告
    assert result.geometry is not None
    assert result.art is not None
    assert any(w.code == "art.evidence.limited" for w in result.warnings)


def test_analyze_deck_dual_source_no_double_audit(tmp_path, monkeypatch):
    import offipy.art.analyze as A
    from offipy.audit.models import AuditConfig, PptxAuditReport

    calls = {"n": 0}

    def fake_audit(p):
        calls["n"] += 1
        return PptxAuditReport(
            schema_version="0.1",
            offipy_version="0.11.6",
            path=p,
            source_sha256="abc",
            slide_size=(10.0, 7.5),
            slide_count=1,
            config=AuditConfig(),
        )

    m = tmp_path / "m.json"
    m.write_text(json.dumps({"slides": []}), encoding="utf-8")
    monkeypatch.setattr(A, "audit_pptx", fake_audit)
    A.analyze_deck(pptx="x.pptx", measurements=str(m))
    assert calls["n"] == 1  # 只审计一次，build_scene 复用 report


def test_analyze_deck_no_source_raises():
    with pytest.raises(InvalidArgumentError):
        analyze_deck()


def test_analyze_deck_slides_dir_empty_dir_raises(tmp_path):
    with pytest.raises(InvalidArgumentError):
        analyze_deck(slides_dir=str(tmp_path / "nope"))


def test_analyze_deck_slides_dir_only(tmp_path):
    pytest.importorskip("PIL")
    from PIL import Image

    d = tmp_path / "slides"
    d.mkdir()
    Image.new("RGB", (100, 100), (255, 255, 255)).save(d / "slide_1.png")
    result = analyze_deck(slides_dir=str(d))
    assert result.geometry is None
    assert result.art is not None
    assert result.art.slides[0].slide_index == 1
    # slides_dir-only 有像素证据 → 不触发 art.evidence.limited
    assert not any(w.code == "art.evidence.limited" for w in result.warnings)
