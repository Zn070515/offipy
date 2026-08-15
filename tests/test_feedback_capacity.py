"""A3 容量自适应（#121）：Lunt&Xu H≈√n 经验法则 + 软告警，永不拒绝模型。

纯函数测试针对 adaptive_hidden_dims / capacity_report；集成测试验证
capacity_warning 只记录不拒绝（soft），且 stats 落盘到 model.json。
"""

from offipy.art import append as art_append
from offipy.art.features_registry import feature_schema_version
from offipy.art.profiles import RULE_TITLE_TOO_SMALL
from offipy.audit import Severity
from offipy.feedback.mlp import adaptive_hidden_dims, capacity_report, params_count
from offipy.feedback.model import load_model, model_file
from offipy.feedback.train import run_training


def test_clamp_lower_bound_single_layer():
    assert adaptive_hidden_dims(1) == (4,)


def test_sqrt_rounds_to_3_clamped_to_4():
    assert adaptive_hidden_dims(10) == (4,)


def test_sqrt_exact_10_single_layer():
    assert adaptive_hidden_dims(100) == (10,)


def test_clamp_upper_bound_double_layer():
    assert adaptive_hidden_dims(4000) == (32, 16)


def test_layer_count_by_sample_count():
    # n<120 → 单层；n≥120 → 双层（含 n=120 边界本身）
    assert len(adaptive_hidden_dims(119)) == 1
    assert len(adaptive_hidden_dims(120)) == 2
    assert len(adaptive_hidden_dims(4)) == 1


def test_second_layer_is_half_of_first():
    assert adaptive_hidden_dims(120) == (11, 5)
    assert adaptive_hidden_dims(4000) == (32, 16)


def test_params_count_single_layer():
    # 单隐层形状 (input_dim, H, 1)：input_dim*H + H + H*1 + 1
    assert params_count(4, (4,)) == 25
    assert params_count(4, (4,)) == 4 * 4 + 4 + 4 + 1


def test_params_count_double_layer():
    # 双层：前一层 fan_out 层后一层 fan_in，逐层累加权重与偏置
    assert params_count(8, (32, 16)) == 8 * 32 + 32 + 32 * 16 + 16 + 16 * 1 + 1


def test_capacity_ok_when_samples_abundant():
    rep = capacity_report(1000, 4, (4,))  # params=25 → spp=40.0
    assert rep["params"] == 25
    assert rep["samples_per_param"] == 40.0
    assert rep["level"] == "ok"
    assert rep["hidden_dims"] == [4]
    assert rep["n"] == 1000 and rep["input_dim"] == 4


def test_capacity_warn_when_samples_scarce():
    rep = capacity_report(100, 4, (4,))  # spp=4.0 → warn
    assert rep["samples_per_param"] == 4.0
    assert rep["level"] == "warn"


def test_capacity_critical_when_samples_very_scarce():
    rep = capacity_report(10, 4, (4,))  # spp=0.4 → critical
    assert rep["samples_per_param"] == 0.4
    assert rep["level"] == "critical"


# --- 集成：软告警只记录，不拒绝写盘 ---


def _add(tmp_path, rule_id, action, n, *, features):
    for _ in range(n):
        art_append(
            "balanced",
            rule_id,
            action,
            Severity.MID,
            feedback_dir=tmp_path,
            features=features,
            feature_schema_version=feature_schema_version(),
        )


def test_capacity_warning_is_soft_and_never_rejects_model(tmp_path):
    # n=16 独立样本 + (4,) 隐层 → params≥13 → spp<5 → 至少 warn
    _add(
        tmp_path,
        RULE_TITLE_TOO_SMALL,
        "fixed",
        12,
        features={"finding.severity_ordinal": 3.0, "finding.confidence": 1.0},
    )
    _add(
        tmp_path,
        RULE_TITLE_TOO_SMALL,
        "accepted",
        4,
        features={"finding.severity_ordinal": 1.0, "finding.confidence": 0.2},
    )
    res = run_training(tmp_path, min_pairs=0)
    assert res["trained"] is True  # 软告警绝不拒绝
    assert res["samples"] == 16
    assert res["capacity_warning"] is True
    assert res["capacity"]["level"] in ("warn", "critical")
    assert res["capacity"]["samples_per_param"] < 5.0
    # n=16 → H=clamp(round(√16))=4 单层；两特征完全相关 → dedup 后 input_dim=1 → params=13
    assert res["capacity"]["hidden_dims"] == [4]
    assert res["capacity"]["params"] == params_count(res["capacity"]["input_dim"], (4,))
    assert res["capacity"]["params"] == 13
    # 模型确实落盘，且 stats 带 capacity 信息
    data = load_model(model_file(tmp_path))
    assert data is not None
    assert data["stats"]["capacity_warning"] is True
    assert data["stats"]["capacity"]["samples_per_param"] == res["capacity"]["samples_per_param"]
