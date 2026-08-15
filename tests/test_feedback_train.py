"""离线训练编排：样本过滤 / 配对阈值 / F2-E 保留旧模型 / 训练成功写模型。"""

from offipy.art import append as art_append
from offipy.art.features_registry import feature_schema_version
from offipy.art.feedback import ART_FEEDBACK_FILE
from offipy.art.profiles import RULE_TITLE_TOO_SMALL
from offipy.audit import Severity
from offipy.feedback.model import load_model, model_file
from offipy.feedback.train import run_training


def _add(tmp_path, rule_id, action, n, *, profile="balanced", features=None, version=None):
    for _ in range(n):
        art_append(
            profile,
            rule_id,
            action,
            Severity.MID,
            feedback_dir=tmp_path,
            features=features or {"finding.confidence": 0.5},
            feature_schema_version=version or feature_schema_version(),
        )


def _features(ordinal, conf):
    return {"finding.severity_ordinal": ordinal, "finding.confidence": conf}


def test_insufficient_pairs_reports_status(tmp_path):
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "fixed", 2)
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "accepted", 2)  # 4 pairs < _MIN_PAIRS
    res = run_training(tmp_path)
    assert res["trained"] is False
    assert res["reason"] == "insufficient_pairs"
    assert res["pairs"] == 4
    assert not model_file(tmp_path).exists()


def test_empty_pairs_with_min_pairs_zero_reports_insufficient(tmp_path):
    """min_pairs=0 + 空 pairs（全 fixed 无 accepted）→ insufficient_pairs，不抛 IndexError。"""
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "fixed", 4)
    res = run_training(tmp_path, min_pairs=0)
    assert res["trained"] is False
    assert res["reason"] == "insufficient_pairs"
    assert res["pairs"] == 0
    assert res["samples"] == 4
    assert not model_file(tmp_path).exists()


def test_insufficient_pairs_includes_per_rule_diagnosis(tmp_path):
    """insufficient_pairs 时带 per_rule 逐规则诊断（#117），key 完整、供用户行动。"""
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "fixed", 2)
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "accepted", 2)  # 4 pairs < _MIN_PAIRS
    res = run_training(tmp_path)
    assert res["trained"] is False
    assert res["reason"] == "insufficient_pairs"
    per_rule = res["per_rule"]
    assert set(per_rule) == {RULE_TITLE_TOO_SMALL}
    d = per_rule[RULE_TITLE_TOO_SMALL]
    assert set(d) == {"fixed", "accepted", "pairs", "single_direction", "suggest"}
    assert d["fixed"] == 2
    assert d["accepted"] == 2
    assert d["pairs"] == 4
    assert d["single_direction"] is False
    assert d["suggest"] > 0


def test_no_valid_samples_reports_reason(tmp_path):
    art_append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    res = run_training(tmp_path)
    assert res["trained"] is False
    assert res["reason"] == "no_valid_samples"


def test_old_schema_samples_skipped(tmp_path):
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "fixed", 6, version="0")
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "accepted", 6, version="0")
    res = run_training(tmp_path)  # 版本不符 → 无有效样本
    assert res["trained"] is False
    assert res["reason"] == "no_valid_samples"


def test_train_success_writes_model(tmp_path):
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "fixed", 12, features=_features(3.0, 1.0))
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "accepted", 4, features=_features(1.0, 0.2))
    res = run_training(tmp_path, min_pairs=0)
    assert res["trained"] is True
    assert res["pairs"] == 48
    data = load_model(model_file(tmp_path))
    assert data is not None
    assert data["input_schema_version"] == feature_schema_version()


def test_train_failure_keeps_old_model(tmp_path):
    """F2-E：train 失败（样本不足）不删除/覆盖已有 model.json。"""
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "fixed", 12, features=_features(3.0, 1.0))
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "accepted", 4, features=_features(1.0, 0.2))
    first = run_training(tmp_path, min_pairs=0)
    assert first["trained"] is True
    old_bytes = model_file(tmp_path).read_bytes()
    # 现在清空记录，再 train → 样本不足 → 旧模型保留
    (tmp_path / ART_FEEDBACK_FILE).write_text("", encoding="utf-8")
    res = run_training(tmp_path)
    assert res["trained"] is False
    assert model_file(tmp_path).read_bytes() == old_bytes


def test_collapsed_constant_output_rejected(tmp_path):
    """特征全零 → 模型对任何输入输出常数 → 判别力门禁拒绝写模型（#112）。"""
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "fixed", 12, features={"finding.confidence": 0.0})
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "accepted", 12, features={"finding.confidence": 0.0})
    res = run_training(tmp_path, min_pairs=0)
    assert res["trained"] is False
    assert res["reason"] == "model_collapsed"
    assert res["output_std"] == 0.0  # 恒定输出 → 判别力门禁拒绝（#112）
    assert not model_file(tmp_path).exists()


def test_diverged_loss_reports_status_no_model(tmp_path):
    """loss 非有限（NaN 输入）→ 立即中止，返回状态且不写模型（#112）。"""
    _add(tmp_path, RULE_TITLE_TOO_SMALL, "fixed", 12, features={"finding.confidence": float("nan")})
    _add(
        tmp_path,
        RULE_TITLE_TOO_SMALL,
        "accepted",
        12,
        features={"finding.confidence": float("nan")},
    )
    res = run_training(tmp_path, min_pairs=0)
    assert res["trained"] is False
    assert res["reason"] == "training_diverged"
    assert not model_file(tmp_path).exists()
