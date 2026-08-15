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
    return save_model(
        _fake_mlp(),
        input_schema_version=feature_schema_version(),
        output_schema_version="1",
        seed=0,
        hidden_dims=(4,),
        stats={},
        path=model_file(tmp_path),
    )


def test_no_model_returns_none(tmp_path):
    assert learned_adjustments("balanced", feedback_dir=tmp_path) is None


def test_expired_model_returns_none(tmp_path):
    save_model(
        _fake_mlp(),
        input_schema_version="999",
        output_schema_version="1",
        seed=0,
        hidden_dims=(4,),
        stats={},
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
