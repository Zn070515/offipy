"""OUTPUTS head 纯推导函数：量化 / severity_shift / quality.score。"""

import math

from offipy.art.profiles import ALL_RULES
from offipy.audit import Severity
from offipy.feedback.heads import (
    apply_severity_shift,
    quality_score_from_worth,
    quantize_delta,
    severity_shift_from_worth,
)
from offipy.feedback.registry import OUTPUT_SCHEMA_VERSION, OUTPUTS


def test_quantize_delta_threshold():
    assert quantize_delta(0.0) == 0
    assert quantize_delta(0.49) == 0
    assert quantize_delta(0.5) == 1
    assert quantize_delta(1.0) == 1
    assert quantize_delta(-0.6) == -1


def test_severity_shift_clamped_negative_one_to_one():
    assert severity_shift_from_worth(0.8) == 0.8
    assert severity_shift_from_worth(5.0) == 1.0
    assert severity_shift_from_worth(-7.0) == -1.0


def test_apply_severity_shift_rounds_and_clamps():
    assert apply_severity_shift(Severity.LOW, 0.0) is Severity.LOW
    assert apply_severity_shift(Severity.LOW, 0.8) is Severity.MID
    assert apply_severity_shift(Severity.MID, -0.7) is Severity.LOW
    assert apply_severity_shift(Severity.HIGH, 1.0) is Severity.HIGH
    assert apply_severity_shift(Severity.LOW, -5.0) is Severity.LOW


def test_quality_score_range_and_direction():
    bad = quality_score_from_worth(1.0)  # 高 worth = 更可能需要修 = 更差
    neutral = quality_score_from_worth(0.0)
    good = quality_score_from_worth(-1.0)
    assert 0.0 <= bad <= neutral <= good <= 100.0
    # quality_score_from_worth 按 docstring 保留 1 位小数，期望值需同样 round
    assert math.isclose(good, round(100.0 / (1.0 + math.exp(-2.0)), 1), rel_tol=1e-6)


def test_quality_score_worth_scale_normalizes_saturation():
    # scale=1 与 scale=2 在 worth=0 都是 50（中心不变）
    assert quality_score_from_worth(0.0) == 50.0
    assert quality_score_from_worth(0.0, worth_scale=2.0) == 50.0
    # 拉伸 scale 缓解饱和：同一高 worth，scale 越大 → 归一越小 → 分数越高（单调）
    assert quality_score_from_worth(2.0, worth_scale=10.0) > quality_score_from_worth(
        2.0, worth_scale=1.0
    )
    # 非法 scale 防御性回退 1.0
    assert quality_score_from_worth(1.0, worth_scale=0) == quality_score_from_worth(1.0)
    assert quality_score_from_worth(1.0, worth_scale=-3.0) == quality_score_from_worth(1.0)


def test_outputs_registry_has_all_rules_and_three_heads():
    assert OUTPUT_SCHEMA_VERSION == "1"
    for rule_id in ALL_RULES:
        assert f"rule.delta.{rule_id}" in OUTPUTS
    assert "finding.severity_shift" in OUTPUTS
    assert "quality.score" in OUTPUTS
    assert all(s.kind == "derived" for s in OUTPUTS.values())
