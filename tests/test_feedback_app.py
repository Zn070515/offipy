"""FeedbackApp：train/status 可用、has_com_root=False、顶层 numpy-free（F2-F）。"""

import json
import subprocess
import sys
from pathlib import Path

import pytest

from art_helpers import make_scene, make_slide, make_text_element
from offipy.art.profiles import RULE_TITLE_TOO_SMALL
from offipy.audit import Severity
from offipy.exceptions import InvalidArgumentError
from offipy.feedback.app import FeedbackApp


def test_has_no_com_root():
    assert FeedbackApp().has_com_root is False


def test_status_returns_dict(tmp_path):
    res = FeedbackApp().status(feedback_dir=str(tmp_path))
    assert res["samples"] == 0
    assert res["model"] == "none"


def test_train_insufficient_pairs_returns_status(tmp_path):
    res = FeedbackApp().train(feedback_dir=str(tmp_path))
    assert res["trained"] is False


def test_feedback_app_lazy_import_no_numpy():
    """F2-F：从深路径 import FeedbackApp 也不拖 numpy（app.py 顶层仅标准库）。"""
    code = (
        "import sys\n"
        "from offipy.feedback.app import FeedbackApp\n"
        "assert 'numpy' not in sys.modules, 'app.py 顶层不得 import numpy'\n"
    )
    subprocess.run([sys.executable, "-c", code], check=True)


def test_import_offipy_no_numpy():
    code = (
        "import sys\n"
        "import offipy\n"
        "assert 'numpy' not in sys.modules, 'import offipy 不得拖 numpy'\n"
    )
    subprocess.run([sys.executable, "-c", code], check=True)


def test_train_requires_numpy_in_subprocess():
    code = (
        "import sys\n"
        "sys.modules['numpy'] = None\n"
        "from offipy.feedback.app import FeedbackApp\n"
        "try:\n"
        "    FeedbackApp().train()\n"
        "    raise SystemExit('should raise')\n"
        "except RuntimeError as e:\n"
        "    assert 'offipy[feedback]' in str(e)\n"
        "    print('ok')\n"
    )
    subprocess.run([sys.executable, "-c", code], check=True)


def test_train_rejects_non_dir(tmp_path):
    f = tmp_path / "not_a_dir"
    f.write_text("x")
    with pytest.raises(InvalidArgumentError):
        FeedbackApp().train(feedback_dir=str(f))


# ---------------------------------------------------------------- append


def test_append_writes_record(tmp_path):
    """FeedbackApp.append：severity 字符串 + features JSON 字符串 → 落记录。"""
    res = FeedbackApp().append(
        "balanced",
        RULE_TITLE_TOO_SMALL,
        "fixed",
        "MID",
        feedback_dir=str(tmp_path),
        features='{"finding.confidence": 0.5}',
    )
    assert Path(res["record"]).exists()
    from offipy.art.feedback import load_records

    recs = load_records(tmp_path)
    assert len(recs) == 1
    assert recs[0].rule_id == RULE_TITLE_TOO_SMALL
    assert recs[0].action == "fixed"
    assert recs[0].severity == Severity.MID
    assert recs[0].features == {"finding.confidence": 0.5}


def test_append_bad_features_json_raises(tmp_path):
    with pytest.raises(InvalidArgumentError, match="features"):
        FeedbackApp().append(
            "balanced",
            RULE_TITLE_TOO_SMALL,
            "fixed",
            "MID",
            feedback_dir=str(tmp_path),
            features="not-json",
        )


def test_append_bad_severity_raises(tmp_path):
    with pytest.raises(InvalidArgumentError, match="severity"):
        FeedbackApp().append(
            "balanced",
            RULE_TITLE_TOO_SMALL,
            "fixed",
            "bogus",
            feedback_dir=str(tmp_path),
        )


def test_append_auto_fills_schema_version(tmp_path):
    """#143：append 带 features 缺 version → 自动填当前 feature_schema_version。"""
    from offipy.art.features_registry import feature_schema_version as current
    from offipy.art.feedback import load_records

    app = FeedbackApp()
    app.append(
        "balanced",
        RULE_TITLE_TOO_SMALL,
        "fixed",
        "MID",
        feedback_dir=str(tmp_path),
        features={"finding.confidence": 0.5},
    )
    rec = load_records(tmp_path)[0]
    assert rec.feature_schema_version == current()
    assert rec.features == {"finding.confidence": 0.5}


def test_reschema_records_rewrites_expired(tmp_path):
    """#144：过期但有 features 的记录重写为当前 version；无 features 跳过计数。"""
    from offipy.art import append as art_append
    from offipy.art.features_registry import feature_schema_version as current
    from offipy.art.feedback import ART_FEEDBACK_FILE, load_records, reschema_records
    from offipy.audit import Severity

    art_append(
        "balanced",
        "art.hierarchy.title_too_small",
        "fixed",
        Severity.MID,
        feedback_dir=tmp_path,
        features={"finding.confidence": 0.5},
        feature_schema_version="999",
    )
    art_append(
        "balanced", "art.hierarchy.title_too_small", "accepted", Severity.MID, feedback_dir=tmp_path
    )  # 无 features
    res = reschema_records(tmp_path)
    assert res["rewritten"] == 1
    assert res["skipped_no_features"] == 1
    recs = load_records(tmp_path)
    assert recs[0].feature_schema_version == current()
    # 无 features 的 skipped 行必须字节级保留（原 JSON，无 feature_schema_version key）
    lines = (tmp_path / ART_FEEDBACK_FILE).read_text(encoding="utf-8").splitlines()
    skipped_line = [ln for ln in lines if '"action": "accepted"' in ln]
    assert skipped_line and "feature_schema_version" not in skipped_line[0]


def test_reschema_records_empty_dir(tmp_path):
    """#144：无记录文件 → 全零计数，不创建文件。"""
    from offipy.art.feedback import reschema_records

    res = reschema_records(tmp_path)
    assert res == {"rewritten": 0, "skipped_no_features": 0, "already_current": 0}


def test_app_reschema_rewrites_expired(tmp_path):
    """#144：FeedbackApp.reschema 包装 reschema_records——旧记录原地重写 + 计数正确。"""
    from offipy.art import append as art_append
    from offipy.art.features_registry import feature_schema_version as current
    from offipy.art.feedback import load_records
    from offipy.audit import Severity

    art_append(
        "balanced",
        "art.hierarchy.title_too_small",
        "fixed",
        Severity.MID,
        feedback_dir=tmp_path,
        features={"finding.confidence": 0.5},
        feature_schema_version="999",
    )
    art_append(
        "balanced",
        "art.hierarchy.title_too_small",
        "accepted",
        Severity.MID,
        feedback_dir=tmp_path,
        features={"finding.confidence": 0.5},
        feature_schema_version=current(),
    )  # 已当前版本（直接 art_append 不自动补版本，须显式传）
    res = FeedbackApp().reschema(str(tmp_path))
    assert res == {"rewritten": 1, "skipped_no_features": 0, "already_current": 1}
    recs = load_records(tmp_path)
    assert all(r.feature_schema_version == current() for r in recs)


def test_app_reschema_empty_dir(tmp_path):
    """#144：无记录 → 全零计数。"""
    assert FeedbackApp().reschema(str(tmp_path)) == {
        "rewritten": 0,
        "skipped_no_features": 0,
        "already_current": 0,
    }


# ---------------------------------------------------------------- recommend / apply (#160)


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


def test_recommend_no_model_raises(tmp_path):
    """#160：feedback recommend 无有效模型 → 显式报错（不回退 v2 静默推荐）。"""
    with pytest.raises(InvalidArgumentError, match="模型"):
        FeedbackApp().recommend("x.pptx", str(tmp_path))


def test_recommend_projects_adjusted_findings(monkeypatch, tmp_path):
    """#160：有模型 → 只读分析 + 投影 adjusted_findings/suggestions。"""
    import offipy.art.analyze as A
    import offipy.feedback.infer as infer
    from offipy.art.models import (
        ArtFinding,
        ArtReport,
        ArtSlideReport,
        DeckQualityReport,
        DimensionAssessment,
    )

    f = ArtFinding(
        rule_id="art.composition.corner_cluster",
        dimension="composition",
        severity=Severity.MID,
        message="m",
        confidence=0.5,
        slide_index=1,
        details={
            "feedback": {
                "head": "severity_shift",
                "worth": 0.8,
                "shift": 0.8,
                "before": "LOW",
                "after": "MID",
            }
        },
        severity_override=True,
        severity_override_source="feedback",
    )
    rep = DeckQualityReport(
        geometry=None,
        art=ArtReport(
            profile="balanced",
            slides=[
                ArtSlideReport(
                    slide_index=1,
                    dimensions=[
                        DimensionAssessment(
                            dimension="composition", status="assessed", grade="good", findings=[f]
                        ),
                    ],
                )
            ],
            deck_findings=[],
            feedback_adjustments={"art.composition.corner_cluster": 1},
            experimental_score=73.1,
            experimental_score_mode="worth_sigmoid",
        ),
        warnings=[],
    )

    class _Bundle:
        pass

    monkeypatch.setattr(infer.ModelBundle, "load", lambda _d: _Bundle())
    monkeypatch.setattr(A, "analyze_deck", lambda **kw: rep)
    res = FeedbackApp().recommend("x.pptx", str(tmp_path), profile="balanced")
    assert res["profile"] == "balanced"
    assert res["feedback_adjustments"] == {"art.composition.corner_cluster": 1}
    assert res["experimental_score"] == 73.1
    assert res["adjusted_findings"] == [
        {
            "dimension": "composition",
            "slide_index": 1,
            "rule_id": "art.composition.corner_cluster",
            "severity_before": "LOW",
            "severity_after": "MID",
            "worth": 0.8,
            "shift": 0.8,
        }
    ]
    assert res["suggestions"]  # project_suggestions 正常投影


def test_apply_persists_and_deck_uses(monkeypatch, tmp_path):
    """#160：apply 把 rule.delta 持久化到 profile 存储；feedback=False 的 deck 分析读它。"""
    import offipy.art.profiles as P
    import offipy.feedback.infer as infer
    from offipy.art.analyze import analyze_scene
    from offipy.art.profiles import RULE_CORNER_CLUSTER

    monkeypatch.setattr(P, "PROFILE_STORE_DIR", tmp_path)  # 默认存储指到 tmp，不碰真实 home
    monkeypatch.setattr(infer, "learned_adjustments", lambda *a, **k: {RULE_CORNER_CLUSTER: 1})
    res = FeedbackApp().apply("balanced", str(tmp_path))
    assert res["profile"] == "balanced"
    assert res["adjustments"] == {RULE_CORNER_CLUSTER: 1}
    store = tmp_path / P.PROFILE_STORE_FILE
    assert store.exists()

    # feedback=False 的 analyze_scene 现在也吃到持久化的 rule.delta（LOW → MID）
    report = analyze_scene(_build_scene(), profile="balanced")
    f = _finding(report, RULE_CORNER_CLUSTER)
    assert f is not None
    assert f.severity == Severity.MID
    assert f.severity_override_source == "feedback"


def test_apply_unknown_profile_raises(tmp_path):
    with pytest.raises(InvalidArgumentError, match="profile"):
        FeedbackApp().apply("bogus", str(tmp_path))


def test_apply_no_model_raises(tmp_path):
    with pytest.raises(InvalidArgumentError, match="模型"):
        FeedbackApp().apply("balanced", str(tmp_path))


def test_persisted_store_corrupt_and_stale(tmp_path):
    """#160：损坏/非 ±1/NaN 值 → 容错读取 {}；save 原子写可重复。"""
    from offipy.art.profiles import (
        load_persisted_adjustments,
        save_persisted_adjustments,
    )

    p = tmp_path / "art_profiles.json"
    p.write_text("{ bad json", encoding="utf-8")
    assert load_persisted_adjustments(tmp_path) == {}  # 损坏 → {}

    p.write_text(json.dumps({"balanced": {"r1": 1, "r2": 0.5, "r3": "x"}}), encoding="utf-8")
    assert load_persisted_adjustments(tmp_path) == {"balanced": {"r1": 1}}  # 只留 ±1 int

    save_persisted_adjustments({"balanced": {"r1": -1}}, tmp_path)
    assert load_persisted_adjustments(tmp_path) == {"balanced": {"r1": -1}}  # 原子写可读回
