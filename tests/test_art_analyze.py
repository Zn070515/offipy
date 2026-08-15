import dataclasses
import json
from pathlib import Path

import pytest

from art_helpers import make_scene, make_slide, make_text_element
from offipy.art.adapters import build_scene
from offipy.art.analyze import analyze_deck, analyze_scene
from offipy.art.feedback import ART_FEEDBACK_FILE, append, apply_feedback
from offipy.art.profiles import (
    RULE_CORNER_CLUSTER,
    RULE_NO_ACCENT,
    RULE_TITLE_DRIFT,
    get_profile,
)
from offipy.audit import Severity
from offipy.exceptions import InvalidArgumentError

FIXTURES = Path(__file__).parent / "fixtures" / "art"


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


def _findings_map(report):
    """{(slide_index, dimension, rule_id): finding}，用于跨报告逐条对比。"""
    return {
        (s.slide_index, d.dimension, f.rule_id): f
        for s in report.slides
        for d in s.dimensions
        for f in d.findings
    }


def test_analyze_scene_real_fixture_feedback_flow(tmp_path):
    """S2 端到端：真实 v0.12.2 报告 + 反馈闭环。

    - feedback 默认不加载（baseline 无任何 override 标记）；
    - corner_cluster 3×fixed → +1 → LOW→MID（source=feedback）；
    - no_accent 3×accepted → −1 → LOW 饱和不变，且不产生虚假 provenance 标记；
    - 仅目标规则变化一级，其余 finding 逐条与 baseline 一致；
    - 用户显式 override 压过反馈 delta（绝对 severity，source=user）；
    - 内置 profile 共享对象从未被 mutate。
    """
    raw = (FIXTURES / "real_measurements.json").read_text(encoding="utf-8")
    scene = build_scene(measurements=raw)

    baseline = analyze_scene(scene, profile="balanced")  # feedback 默认 False
    assert baseline.schema_version == "0.3"  # 报告 schema 0.3
    assert _finding(baseline, RULE_CORNER_CLUSTER).severity == Severity.LOW
    assert _finding(baseline, RULE_NO_ACCENT).severity == Severity.LOW

    for _ in range(3):
        append("balanced", RULE_CORNER_CLUSTER, "fixed", Severity.MID, feedback_dir=tmp_path)
        append("balanced", RULE_NO_ACCENT, "accepted", Severity.MID, feedback_dir=tmp_path)

    report = analyze_scene(scene, profile="balanced", feedback=True, feedback_dir=tmp_path)
    cc = _finding(report, RULE_CORNER_CLUSTER)
    assert cc.severity == Severity.MID  # LOW +1，只升一级
    assert cc.severity_override is True
    assert cc.severity_override_source == "feedback"
    na = _finding(report, RULE_NO_ACCENT)
    assert na.severity == Severity.LOW  # −1 在 LOW 饱和
    assert na.severity_override is False  # 饱和不改值 → 无虚假标记
    assert na.severity_override_source is None

    # 只有 corner_cluster LOW→MID，其余 finding 逐条与 baseline 一致
    assert len(_findings_map(baseline)) == len(_findings_map(report))
    for key, b in _findings_map(baseline).items():
        r = _findings_map(report)[key]
        if b.rule_id == RULE_CORNER_CLUSTER:
            assert b.severity == Severity.LOW and r.severity == Severity.MID
        else:
            assert r.severity == b.severity
            assert r.severity_override == b.severity_override
            assert r.severity_override_source == b.severity_override_source

    # 用户显式 override 压过反馈 delta
    prof = dataclasses.replace(
        get_profile("balanced"),
        name="custom",
        severity_overrides={RULE_CORNER_CLUSTER: Severity.HIGH},
    )
    report2 = analyze_scene(scene, profile=prof, feedback=True, feedback_dir=tmp_path)
    cc2 = _finding(report2, RULE_CORNER_CLUSTER)
    assert cc2.severity == Severity.HIGH  # 绝对 user override，不叠加反馈 +1
    assert cc2.severity_override is True
    assert cc2.severity_override_source == "user"

    # 内置 profile 共享对象从未被 mutate；store 聚合喂给了本轮运行
    p = get_profile("balanced")
    assert p.feedback_severity_adjustments == {}
    assert p.severity_overrides == {}
    assert apply_feedback("balanced", feedback_dir=tmp_path).feedback_severity_adjustments == {
        RULE_CORNER_CLUSTER: 1,
        RULE_NO_ACCENT: -1,
    }


def test_analyze_scene_corrupt_feedback_line_skipped(tmp_path):
    """脏反馈行不破坏分析：坏行跳过，有效记录仍生效。"""
    raw = (FIXTURES / "real_measurements.json").read_text(encoding="utf-8")
    scene = build_scene(measurements=raw)
    for _ in range(3):
        append("balanced", RULE_CORNER_CLUSTER, "fixed", Severity.MID, feedback_dir=tmp_path)
    f = tmp_path / ART_FEEDBACK_FILE
    lines = f.read_text(encoding="utf-8").splitlines()
    f.write_text("not-json\n" + "\n".join(lines), encoding="utf-8")
    report = analyze_scene(scene, profile="balanced", feedback=True, feedback_dir=tmp_path)
    cc = _finding(report, RULE_CORNER_CLUSTER)
    assert cc.severity == Severity.MID
    assert cc.severity_override is True
    assert cc.severity_override_source == "feedback"


# ---------------------------------------------------------------------------
# v0.17 学习路径集成
# ---------------------------------------------------------------------------


def test_analyze_feedback_no_model_falls_back_to_v2(tmp_path):
    # 无 model.json → 冷启动 = v2（recommend_adjustments 照旧）
    for _ in range(3):
        append("balanced", RULE_CORNER_CLUSTER, "fixed", Severity.MID, feedback_dir=tmp_path)
    report = analyze_scene(_build_scene(), profile="balanced", feedback=True, feedback_dir=tmp_path)
    f = _finding(report, RULE_CORNER_CLUSTER)
    assert f is not None
    assert f.severity == Severity.MID
    assert f.severity_override_source == "feedback"


def test_analyze_learning_pass_applies_severity_shift(monkeypatch, tmp_path):
    """有有效模型 → rule-computed finding 被 severity_shift；override 的跳过。"""
    from offipy.art import features_registry
    from offipy.feedback import infer
    from offipy.feedback.mlp import MLP
    from offipy.feedback.model import model_file, save_model

    save_model(
        MLP(input_dim=len(features_registry.feature_keys()), hidden_dims=(4,), seed=0),
        input_schema_version=features_registry.feature_schema_version(),
        output_schema_version="1",
        seed=0,
        hidden_dims=(4,),
        stats={},
        path=model_file(tmp_path),
    )

    # 确定性：monkeypatch model_worth 让 corner_cluster finding 命中 +0.8
    def fake_worth(features, mlp=None):
        if features.get("finding.rule_id.art.composition.corner_cluster") == 1.0:
            return 0.8
        return 0.0

    monkeypatch.setattr(infer, "model_worth", fake_worth)
    report = analyze_scene(_build_scene(), profile="balanced", feedback=True, feedback_dir=tmp_path)
    f = _finding(report, RULE_CORNER_CLUSTER)
    assert f is not None
    # 基线 corner_cluster = LOW；+0.8 → round(1+0.8)=2 → MID（未被 override，shift 生效）
    assert f.severity == Severity.MID
    assert f.severity_override is False  # severity_shift 不标 override


def test_analyze_learning_quality_score_replaces_formula(monkeypatch, tmp_path):
    from offipy.art import features_registry
    from offipy.feedback import infer
    from offipy.feedback.mlp import MLP
    from offipy.feedback.model import model_file, save_model

    save_model(
        MLP(input_dim=len(features_registry.feature_keys()), hidden_dims=(4,), seed=0),
        input_schema_version=features_registry.feature_schema_version(),
        output_schema_version="1",
        seed=0,
        hidden_dims=(4,),
        stats={},
        path=model_file(tmp_path),
    )
    monkeypatch.setattr(infer, "model_worth", lambda feats, mlp=None: -0.5)  # 可接受 → 高分
    report = analyze_scene(
        _build_scene(),
        profile="balanced",
        feedback=True,
        feedback_dir=tmp_path,
        include_experimental_score=True,
    )
    assert report.experimental_score == 73.1


def test_analyze_corrupt_model_falls_back_to_v2(tmp_path):
    """损坏但 schema 匹配的 model.json → analyze_scene 不抛异常，回退 v2。"""
    from offipy.art import features_registry
    from offipy.feedback.mlp import MLP
    from offipy.feedback.model import model_file, save_model

    save_model(
        MLP(input_dim=len(features_registry.feature_keys()), hidden_dims=(4,), seed=0),
        input_schema_version=features_registry.feature_schema_version(),
        output_schema_version="1",
        seed=0,
        hidden_dims=(4,),
        stats={},
        path=model_file(tmp_path),
    )
    p = model_file(tmp_path)
    data = json.loads(p.read_text(encoding="utf-8"))
    data["hidden_dims"] = [8]  # 与真实权重形状不一致 → weights_from_dict 抛 ValueError
    p.write_text(json.dumps(data), encoding="utf-8")

    for _ in range(3):
        append("balanced", RULE_CORNER_CLUSTER, "fixed", Severity.MID, feedback_dir=tmp_path)
    report = analyze_scene(_build_scene(), profile="balanced", feedback=True, feedback_dir=tmp_path)
    f = _finding(report, RULE_CORNER_CLUSTER)
    assert f is not None
    assert f.severity == Severity.MID  # v2 回退：apply_feedback 后仍是 MID
    assert f.severity_override_source == "feedback"


def test_analyze_scene_feedback_requires_dir():
    """#113：feedback=True 无 feedback_dir → InvalidArgumentError（禁静默全局模型）。"""
    with pytest.raises(InvalidArgumentError, match="feedback_dir"):
        analyze_scene(_build_scene(), profile="balanced", feedback=True)


def test_analyze_deck_feedback_requires_dir():
    """#113：analyze_deck 同契约——feedback=True 必须显式 feedback_dir。"""
    m = {"width": 1920.0, "height": 1080.0, "records": []}
    with pytest.raises(InvalidArgumentError, match="feedback_dir"):
        analyze_deck(measurements=m, profile="balanced", feedback=True)
