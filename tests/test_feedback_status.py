"""status：样本数 / 配对潜力 / 模型状态（顶层 numpy-free）。"""

from offipy.art import append as art_append
from offipy.art.features_registry import feature_schema_version
from offipy.art.profiles import RULE_TITLE_TOO_SMALL
from offipy.audit import Severity
from offipy.feedback.status import report_status


def _add(tmp_path, action, n):
    for _ in range(n):
        art_append(
            "balanced",
            RULE_TITLE_TOO_SMALL,
            action,
            Severity.MID,
            feedback_dir=tmp_path,
            features={"finding.confidence": 0.5},
            feature_schema_version=feature_schema_version(),
        )


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
