"""样本级 repeated stratified CV（#119）：分层 5 折 × 3 repeats 的离线验证。

Kajimura 2022：按 rule_id 对独立样本分层，绝不 pair-level split（防 leakage）。
每折在训练样本上重拟合预处理 + 重建 pairs + 训练单 member，在留出折样本的
pairs 上评估 pairwise accuracy 与 separation。输出 95% 下界触 0.5 →
random_indistinguishable（soft flag，不 hard reject）。
"""

import numpy as np

from offipy.art.features_registry import feature_schema_version
from offipy.art.feedback import ArtFeedbackRecord
from offipy.art.profiles import RULE_DIMENSIONS, RULE_LOW_CONTRAST, RULE_TITLE_TOO_SMALL
from offipy.audit import Severity
from offipy.feedback.validation import _fold_splits, repeated_stratified_cv

_FIXED_FEATS = {"finding.confidence": 0.9, "finding.severity_ordinal": 3.0}
_ACCEPTED_FEATS = {"finding.confidence": 0.1, "finding.severity_ordinal": 1.0}


def _rec(rule_id: str, action: str, features: dict[str, float]) -> ArtFeedbackRecord:
    return ArtFeedbackRecord(
        ts="2026-08-15T00:00:00+00:00",
        profile="balanced",
        rule_id=rule_id,
        dimension=RULE_DIMENSIONS[rule_id],
        severity=Severity.MID,
        action=action,
        features=features,
        feature_schema_version=feature_schema_version(),
    )


def _separable_records() -> list[ArtFeedbackRecord]:
    """2 rule × (15 fixed 高置信 + 15 accepted 低置信)：模型可完美分离。"""
    recs = []
    for rule_id in (RULE_TITLE_TOO_SMALL, RULE_LOW_CONTRAST):
        for _ in range(15):
            recs.append(_rec(rule_id, "fixed", _FIXED_FEATS))
            recs.append(_rec(rule_id, "accepted", _ACCEPTED_FEATS))
    return recs


def test_separable_data_scores_high_accuracy():
    """可分离数据 → pairwise accuracy≈1、separation>0、random_indistinguishable=False。"""
    res = repeated_stratified_cv(_separable_records(), hidden_dims=(8,), seed=42)
    assert res["folds"] > 0
    assert res["accuracy_pooled"] > 0.8
    assert res["accuracy_mean"] > 0.8
    assert res["separation"] > 0.0
    assert res["random_indistinguishable"] is False


def test_identical_features_indistinguishable():
    """全部样本特征相同 → 模型无法区分 → accuracy≈0.5 → random_indistinguishable=True。

    注意：特征全等时预处理走 kept 空回退（零方差 drop 全列），不应崩。
    """
    recs = [
        _rec(RULE_TITLE_TOO_SMALL, action, {"finding.confidence": 0.5})
        for _ in range(20)
        for action in ("fixed", "accepted")
    ]
    res = repeated_stratified_cv(recs, hidden_dims=(8,), seed=42)
    assert res["folds"] > 0
    assert res["accuracy_pooled"] <= 0.6
    assert res["separation"] == 0.0
    assert res["random_indistinguishable"] is True


def test_fold_splits_partition_same_rule_no_leak():
    """样本级防泄露：同 rule 记录分层成 k 折后，每折记录不跨折泄露（partition）。"""
    records = [
        _rec(
            RULE_TITLE_TOO_SMALL,
            "fixed" if i % 2 == 0 else "accepted",
            {"finding.confidence": 0.5},
        )
        for i in range(10)
    ]
    rng = np.random.default_rng(0)
    folds = _fold_splits(records, k=5, rng=rng)
    assert len(folds) == 5
    seen: set[int] = set()
    for fold in folds:
        ids = {id(r) for r in fold}
        assert len(ids) == len(fold)  # 折内无重复记录
        assert not (ids & seen)  # 记录不跨折重复
        seen |= ids
    assert len(seen) == len(records)  # 所有记录恰好进入某一折
    for fi, fold in enumerate(folds):
        others = {id(r) for i, f in enumerate(folds) if i != fi for r in f}
        for r in fold:
            assert id(r) not in others  # 留出折记录不进其余折（train 不漏 val）


def test_cv_deterministic_same_seed():
    """同 seed 两次调用结果完全一致（rng 固定 + 逐折 seed+fi）。"""
    records = _separable_records()
    a = repeated_stratified_cv(records, hidden_dims=(8,), seed=42)
    b = repeated_stratified_cv(records, hidden_dims=(8,), seed=42)
    assert a == b
