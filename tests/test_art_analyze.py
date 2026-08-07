import dataclasses
import json

import pytest

from art_helpers import make_scene, make_slide, make_text_element
from offipy.art.analyze import analyze_deck, analyze_scene
from offipy.art.feedback import append
from offipy.art.profiles import RULE_CORNER_CLUSTER, RULE_TITLE_DRIFT, get_profile
from offipy.audit import Severity
from offipy.exceptions import InvalidArgumentError


def _build_scene():
    """单页场景：corner_cluster 规则在 balanced 下触发（LOW）。"""
    return make_scene(
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


def _finding(report, rule_id):
    for s in report.slides:
        for d in s.dimensions:
            for f in d.findings:
                if f.rule_id == rule_id:
                    return f
    return None


def _content_slide_measurement(title_x=0.1, body_x=0.1):
    def el(rec_id, x, y, w, h, font_size, cls, text):
        return {
            "id": rec_id,
            "kind": "text",
            "rect": {"x": x, "y": y, "w": w, "h": h},
            "className": cls,
            "style": {"fontSize": font_size},
            "runs": [{"text": text, "fontSize": font_size}],
            "text": text,
        }

    return {
        "width": 1920.0,
        "height": 1080.0,
        "records": [
            el("t", title_x * 1920, 0.05 * 1080, 0.4 * 1920, 0.08 * 1080, 48, "title", "Title"),
            el("b", body_x * 1920, 0.2 * 1080, 0.5 * 1920, 0.06 * 1080, 24, "body", "Body"),
            el("b2", body_x * 1920, 0.3 * 1080, 0.5 * 1920, 0.06 * 1080, 20, "body", "Body2"),
        ],
    }


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


def test_rule_dimensions_agree_with_dimension_rules():
    import offipy.art.analyze as A
    from offipy.art.profiles import RULE_DIMENSIONS, RULE_MARGIN_DRIFT, RULE_TITLE_DRIFT

    for dimension, specs in A._DIMENSION_RULES.items():
        for spec in specs:
            assert RULE_DIMENSIONS[spec.rule_id] == spec.dimension == dimension
    # deck 级一致性规则不在 _DIMENSION_RULES 列表里，单独断言
    assert RULE_DIMENSIONS[RULE_TITLE_DRIFT] == "consistency"
    assert RULE_DIMENSIONS[RULE_MARGIN_DRIFT] == "consistency"


# ---------------------------------------------------------------------------
# 分析入口的反馈接入（feedback / feedback_dir 关键字参数）
# ---------------------------------------------------------------------------


def test_analyze_scene_feedback_false_matches_default(tmp_path):
    """feedback=False 与不传反馈参数逐字段一致（含 feedback_dir 传空库）。"""
    scene = _build_scene()
    default = analyze_scene(scene, profile="balanced")
    assert dataclasses.asdict(default) == dataclasses.asdict(
        analyze_scene(scene, profile="balanced", feedback=False)
    )
    assert dataclasses.asdict(default) == dataclasses.asdict(
        analyze_scene(scene, profile="balanced", feedback=False, feedback_dir=tmp_path)
    )


def test_analyze_scene_feedback_true_steps_severity(tmp_path):
    """corner_cluster +1 → LOW 升 MID，severity_override_source == feedback。"""
    for _ in range(3):
        append("balanced", RULE_CORNER_CLUSTER, "fixed", Severity.MID, feedback_dir=tmp_path)
    baseline = analyze_scene(_build_scene(), profile="balanced")
    base = _finding(baseline, RULE_CORNER_CLUSTER)
    assert base is not None
    assert base.severity == Severity.LOW
    report = analyze_scene(_build_scene(), profile="balanced", feedback=True, feedback_dir=tmp_path)
    f = _finding(report, RULE_CORNER_CLUSTER)
    assert f is not None
    assert f.severity == Severity.MID
    assert f.severity_override is True
    assert f.severity_override_source == "feedback"


def test_analyze_scene_custom_profile_with_feedback(tmp_path):
    """自定义 ArtProfile + feedback=True：profile 名保留、override 生效。"""
    prof = dataclasses.replace(
        get_profile("balanced"),
        name="custom",
        severity_overrides={RULE_CORNER_CLUSTER: Severity.HIGH},
    )
    report = analyze_scene(_build_scene(), profile=prof, feedback=True, feedback_dir=tmp_path)
    assert report.profile == "custom"
    f = _finding(report, RULE_CORNER_CLUSTER)
    assert f is not None
    assert f.severity == Severity.HIGH
    assert f.severity_override_source == "user"


def test_analyze_scene_user_override_beats_feedback(tmp_path):
    """user severity_override 压制反馈 delta（source == user，不升两级）。"""
    for _ in range(3):
        append("balanced", RULE_CORNER_CLUSTER, "fixed", Severity.MID, feedback_dir=tmp_path)
    prof = dataclasses.replace(
        get_profile("balanced"),
        severity_overrides={RULE_CORNER_CLUSTER: Severity.HIGH},
    )
    report = analyze_scene(_build_scene(), profile=prof, feedback=True, feedback_dir=tmp_path)
    f = _finding(report, RULE_CORNER_CLUSTER)
    assert f is not None
    assert f.severity == Severity.HIGH
    assert f.severity_override is True
    assert f.severity_override_source == "user"


def test_analyze_scene_feedback_empty_store_no_change(tmp_path):
    """feedback=True 但反馈库为空 → 报告与 feedback=False 完全一致。"""
    scene = _build_scene()
    default = analyze_scene(scene, profile="balanced")
    assert dataclasses.asdict(default) == dataclasses.asdict(
        analyze_scene(scene, profile="balanced", feedback=True, feedback_dir=tmp_path / "nope")
    )


def test_analyze_deck_feedback_adjusts_consistency_rule(tmp_path):
    """deck 级 title_drift +1 → LOW 升 MID，severity_override_source == feedback。"""
    for _ in range(3):
        append("balanced", RULE_TITLE_DRIFT, "fixed", Severity.MID, feedback_dir=tmp_path)
    m = {
        "slides": [
            _content_slide_measurement(title_x=0.1),
            _content_slide_measurement(title_x=0.5),
            _content_slide_measurement(title_x=0.1),
        ]
    }
    baseline = analyze_deck(measurements=m, profile="balanced")
    base = [f for f in baseline.art.deck_findings if f.rule_id == RULE_TITLE_DRIFT]
    assert len(base) == 1
    assert base[0].severity == Severity.LOW
    report = analyze_deck(measurements=m, profile="balanced", feedback=True, feedback_dir=tmp_path)
    fs = [f for f in report.art.deck_findings if f.rule_id == RULE_TITLE_DRIFT]
    assert len(fs) == 1
    assert fs[0].severity == Severity.MID
    assert fs[0].severity_override is True
    assert fs[0].severity_override_source == "feedback"
