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
    assert report_status(tmp_path)["model"] == "valid"


def test_status_model_expired(tmp_path):
    _add(tmp_path, "fixed", 12, features=_discriminative("fixed"))
    _add(tmp_path, "accepted", 4, features=_discriminative("accepted"))
    run_training(tmp_path, min_pairs=0)
    path = model_file(tmp_path)
    data = json.loads(path.read_text(encoding="utf-8"))
    data["input_schema_version"] = "999"
    path.write_text(json.dumps(data), encoding="utf-8")
    assert report_status(tmp_path)["model"] == "expired"
