"""status：样本数 / 配对潜力 / 模型状态（顶层 numpy-free）。"""

import json

from offipy.art import append as art_append
from offipy.art.features_registry import feature_schema_version
from offipy.art.profiles import RULE_TITLE_TOO_SMALL
from offipy.audit import Severity
from offipy.feedback.model import model_file
from offipy.feedback.status import report_status
from offipy.feedback.train import run_training


def _add(tmp_path, action, n, *, features=None):
    for _ in range(n):
        art_append(
            "balanced",
            RULE_TITLE_TOO_SMALL,
            action,
            Severity.MID,
            feedback_dir=tmp_path,
            features=features or {"finding.confidence": 0.5},
            feature_schema_version=feature_schema_version(),
        )


def _discriminative(action):
    """可判别特征：#112 训练门禁下，特征全等会导致模型坍缩成常数被拒。"""
    return {
        "fixed": {"finding.severity_ordinal": 3.0, "finding.confidence": 1.0},
        "accepted": {"finding.severity_ordinal": 1.0, "finding.confidence": 0.2},
    }[action]


def test_status_empty_dir(tmp_path):
    s = report_status(tmp_path)
    assert s["samples"] == 0
    assert s["pair_potential"] == 0
    assert s["model"] == "none"
    assert "effective_dims" not in s
    assert "samples_per_param" not in s
    assert "poor_generalization" not in s


def test_status_counts_samples_and_pairs(tmp_path):
    _add(tmp_path, "fixed", 3)
    _add(tmp_path, "accepted", 2)
    _add(tmp_path, "ignored", 5)
    s = report_status(tmp_path)
    assert s["samples"] == 10  # ignored 也计入样本总数
    assert s["valid_samples"] == 5  # 有 features + 当前 schema 且非 ignored
    assert s["pair_potential"] == 6  # 3 fixed × 2 accepted


def test_status_model_missing_no_file(tmp_path):
    s = report_status(tmp_path)
    assert s["model"] == "none"


def test_status_model_valid(tmp_path):
    _add(tmp_path, "fixed", 12, features=_discriminative("fixed"))
    _add(tmp_path, "accepted", 4, features=_discriminative("accepted"))
    run_training(tmp_path, min_pairs=0)
    s = report_status(tmp_path)
    assert s["model"] == "valid"
    model = json.loads(model_file(tmp_path).read_text(encoding="utf-8"))
    assert s["effective_dims"] == len(model["preprocessing"]["kept"])
    assert s["effective_dims"] > 0
    assert s["samples_per_param"] == model["stats"]["capacity"]["samples_per_param"]
    assert isinstance(s["samples_per_param"], float) and s["samples_per_param"] > 0
    assert isinstance(s["poor_generalization"], bool)


def test_status_model_valid_missing_stats_keys(tmp_path):
    """valid 模型 stats 缺键/异常类型 → 兜底 None 不抛；kept 缺失 → #150 stale。"""
    _add(tmp_path, "fixed", 12, features=_discriminative("fixed"))
    _add(tmp_path, "accepted", 4, features=_discriminative("accepted"))
    run_training(tmp_path, min_pairs=0)
    path = model_file(tmp_path)
    data = json.loads(path.read_text(encoding="utf-8"))

    # capacity 缺 → samples_per_param None；kept 仍在 → effective_dims > 0
    del data["stats"]["capacity"]
    path.write_text(json.dumps(data), encoding="utf-8")
    s = report_status(tmp_path)
    assert s["model"] == "valid"
    assert s["effective_dims"] > 0
    assert s["samples_per_param"] is None

    # capacity 非 dict（病理，kept 仍在）→ samples_per_param None，仍 valid 不抛
    data = json.loads(path.read_text(encoding="utf-8"))
    data["stats"]["capacity"] = "not-a-dict"
    path.write_text(json.dumps(data), encoding="utf-8")
    s = report_status(tmp_path)
    assert s["model"] == "valid"
    assert s["effective_dims"] > 0
    assert s["samples_per_param"] is None

    # kept 缺失 → #150 stale（预处理不完整，不冒充 valid）
    data = json.loads(path.read_text(encoding="utf-8"))
    del data["preprocessing"]["kept"]
    path.write_text(json.dumps(data), encoding="utf-8")
    s = report_status(tmp_path)
    assert s["model"] == "stale"
    assert "effective_dims" not in s
    assert "samples_per_param" not in s


def test_status_model_expired(tmp_path):
    _add(tmp_path, "fixed", 12, features=_discriminative("fixed"))
    _add(tmp_path, "accepted", 4, features=_discriminative("accepted"))
    run_training(tmp_path, min_pairs=0)
    path = model_file(tmp_path)
    data = json.loads(path.read_text(encoding="utf-8"))
    data["input_schema_version"] = "999"
    path.write_text(json.dumps(data), encoding="utf-8")
    s = report_status(tmp_path)
    assert s["model"] == "expired"
    assert "effective_dims" not in s
    assert "samples_per_param" not in s
    assert "poor_generalization" not in s


def test_status_model_stale_when_kept_oob(tmp_path):
    """#150：kept 越界 → status 报 stale 而非 valid。"""
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
    path = model_file(tmp_path)
    data = json.loads(path.read_text(encoding="utf-8"))
    data["preprocessing"]["kept"][0] = (
        9999  # 保长度改写（input_dim 与权重匹配），逼出 kept 越界 → stale
    )
    path.write_text(json.dumps(data), encoding="utf-8")
    s = report_status(tmp_path)
    assert s["model"] == "stale"
    assert "effective_dims" not in s


def test_status_excluded_breakdown(tmp_path):
    """#144：schema 过期/无特征记录在 status 里显式报出 excluded。"""
    from offipy.art import append as art_append
    from offipy.art.features_registry import feature_schema_version as current
    from offipy.audit import Severity

    art_append(
        "balanced",
        "art.hierarchy.title_too_small",
        "fixed",
        Severity.MID,
        feedback_dir=tmp_path,
        features={"finding.confidence": 0.5},
        feature_schema_version=current(),
    )
    art_append(
        "balanced",
        "art.hierarchy.title_too_small",
        "fixed",
        Severity.MID,
        feedback_dir=tmp_path,
        features={"finding.confidence": 0.5},
        feature_schema_version="999",
    )
    s = report_status(tmp_path)
    assert s["excluded"] == {"schema_mismatch": 1, "no_features": 0, "ignored": 0, "other": 0}


def test_status_per_rule_diagnosis(tmp_path):
    """#152：status 输出逐规则样本诊断（fixed/accepted/pairs/single_direction/suggest）。"""
    _add(tmp_path, "fixed", 12, features=_discriminative("fixed"))
    _add(tmp_path, "accepted", 4, features=_discriminative("accepted"))
    run_training(tmp_path, min_pairs=0)
    s = report_status(tmp_path)
    per = s["per_rule"]
    assert set(per) == {RULE_TITLE_TOO_SMALL}
    row = per[RULE_TITLE_TOO_SMALL]
    assert row["fixed"] == 12
    assert row["accepted"] == 4
    assert row["pairs"] == 48
    assert row["single_direction"] is False
    assert "suggest" in row


def test_status_surfaces_saturation(tmp_path):
    """#151：饱和检测结果经 stats 持久化，status 透出（valid 分支）。"""
    _add(tmp_path, "fixed", 12, features=_discriminative("fixed"))
    _add(tmp_path, "accepted", 4, features=_discriminative("accepted"))
    run_training(tmp_path, min_pairs=0)
    assert report_status(tmp_path)["saturation"] is False  # 判别性样本 → 未饱和
    path = model_file(tmp_path)
    data = json.loads(path.read_text(encoding="utf-8"))
    data["stats"]["saturation"] = True
    path.write_text(json.dumps(data), encoding="utf-8")
    assert report_status(tmp_path)["saturation"] is True


def test_status_model_corrupt_weights(tmp_path):
    """#147：schema 匹配但权重形状损坏 → status 报 corrupt 而非 valid。"""
    _add(tmp_path, "fixed", 12, features=_discriminative("fixed"))
    _add(tmp_path, "accepted", 4, features=_discriminative("accepted"))
    run_training(tmp_path, min_pairs=0)
    path = model_file(tmp_path)
    data = json.loads(path.read_text(encoding="utf-8"))
    data["members"][0]["hidden_dims"] = [8]  # 与权重形状不一致 → weights_from_dict 抛 ValueError
    path.write_text(json.dumps(data), encoding="utf-8")
    s = report_status(tmp_path)
    assert s["model"] == "corrupt"
    assert "effective_dims" not in s
    assert "samples_per_param" not in s
