"""推理：rule.delta（历史聚合）/ severity_shift（当前 finding）/ quality.score。"""

from unittest import mock

from offipy.art.features_registry import feature_keys, feature_schema_version
from offipy.art.profiles import (
    RULE_LOW_CONTRAST,
    RULE_OFF_BALANCE,
    RULE_TITLE_TOO_SMALL,
)
from offipy.audit import Severity
from offipy.feedback import infer
from offipy.feedback.infer import learned_adjustments, quality_score_for_report
from offipy.feedback.mlp import MLP
from offipy.feedback.model import model_file, save_model


def _fake_mlp():
    return MLP(input_dim=len(feature_keys()), hidden_dims=(4,), seed=0)


def _write_model(tmp_path):
    n = len(feature_keys())
    # 恒等预处理：not-yet-applied transform 为 no-op（A6 ModelBundle 接入真实 transform）
    return save_model(
        members=[(0, _fake_mlp())],
        input_schema_version=feature_schema_version(),
        output_schema_version="1",
        seed=0,
        stats={},
        preprocessing={"kept": list(range(n)), "mean": [0.0] * n, "scale": [1.0] * n},
        calibration={"worth_scale": 1.0},
        abstain={"worth_margin_p25": 0.0, "std_p80": 0.0},
        path=model_file(tmp_path),
    )


def test_no_model_returns_none(tmp_path):
    assert learned_adjustments("balanced", feedback_dir=tmp_path) is None


def test_expired_model_returns_none(tmp_path):
    n = len(feature_keys())
    # 恒等预处理：not-yet-applied transform 为 no-op（A6 ModelBundle 接入真实 transform）
    save_model(
        members=[(0, _fake_mlp())],
        input_schema_version="999",
        output_schema_version="1",
        seed=0,
        stats={},
        preprocessing={"kept": list(range(n)), "mean": [0.0] * n, "scale": [1.0] * n},
        calibration={"worth_scale": 1.0},
        abstain={"worth_margin_p25": 0.0, "std_p80": 0.0},
        path=model_file(tmp_path),
    )
    assert learned_adjustments("balanced", feedback_dir=tmp_path) is None


def test_learned_adjustments_quantizes_and_omits_zero(monkeypatch, tmp_path):
    _write_model(tmp_path)  # 先落有效模型，learned_adjustments 才能越过 _load_valid_model
    monkeypatch.setattr(
        infer,
        "model_worth",
        mock.Mock(side_effect=[0.8, 0.8, 0.8, -0.9, -0.9, -0.9, 0.1, 0.1, 0.1]),
    )
    # monkeypatch 让 load_records 返回我们造的记录
    recs = []
    from offipy.art.feedback import ArtFeedbackRecord

    def mk(rule_id, action):
        return ArtFeedbackRecord(
            ts="t",
            profile="balanced",
            rule_id=rule_id,
            dimension="x",
            severity=Severity.MID,
            action=action,
            features={"f": 1.0},
            feature_schema_version=feature_schema_version(),
        )

    for rule in (RULE_TITLE_TOO_SMALL, RULE_LOW_CONTRAST, RULE_OFF_BALANCE):
        recs += [mk(rule, "fixed"), mk(rule, "fixed"), mk(rule, "accepted")]
    monkeypatch.setattr(infer, "load_records", lambda feedback_dir: recs)
    adj = learned_adjustments("balanced", feedback_dir=tmp_path)
    assert adj == {RULE_TITLE_TOO_SMALL: 1, RULE_LOW_CONTRAST: -1}  # off_balance 均值 0.1 → omit
    assert RULE_OFF_BALANCE not in adj


def test_quality_score_from_worth_mean():
    assert quality_score_for_report(0.0) == 50.0
    assert 0.0 <= quality_score_for_report(1.0) <= quality_score_for_report(-1.0) <= 100.0


def test_learned_adjustments_no_dir_returns_none():
    """#113：无 feedback_dir 不碰全局 ~/.offipy，直接回退（None）。"""
    assert learned_adjustments("balanced") is None


def test_model_bundle_load_rejects_oob_kept(tmp_path):
    """#150：kept 下标越界（schema bump 忘重训）→ load 返回 None，不 IndexError。"""
    from offipy.art import append as art_append
    from offipy.art.features_registry import feature_schema_version
    from offipy.audit import Severity
    from offipy.feedback.infer import ModelBundle
    from offipy.feedback.train import run_training

    disc = {
        "fixed": {"finding.severity_ordinal": 3.0, "finding.confidence": 1.0},
        "accepted": {"finding.severity_ordinal": 1.0, "finding.confidence": 0.2},
    }
    for _ in range(12):
        art_append(
            "balanced",
            "art.hierarchy.title_too_small",
            "fixed",
            Severity.MID,
            feedback_dir=tmp_path,
            features=disc["fixed"],
            feature_schema_version=feature_schema_version(),
        )
    for _ in range(4):
        art_append(
            "balanced",
            "art.hierarchy.title_too_small",
            "accepted",
            Severity.MID,
            feedback_dir=tmp_path,
            features=disc["accepted"],
            feature_schema_version=feature_schema_version(),
        )
    run_training(tmp_path, min_pairs=0)
    import json

    from offipy.feedback.model import model_file

    path = model_file(tmp_path)
    data = json.loads(path.read_text(encoding="utf-8"))
    data["preprocessing"]["kept"][0] = (
        9999  # 保长度改写：input_dim 与权重仍匹配，逼出 kept 越界检查
    )
    path.write_text(json.dumps(data), encoding="utf-8")
    assert ModelBundle.load(tmp_path) is None
