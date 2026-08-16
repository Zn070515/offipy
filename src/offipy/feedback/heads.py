"""输出 head 的纯推导函数（stdlib-only，无 numpy）。

三个 v0.17 head 都从可学习 worth latent 派生：
- rule.delta.<rule_id>：历史记录 worth 均值 → 量化 ±1/0（阈值 DELTA_THRESHOLD）
- finding.severity_shift：当前 finding worth → clamp [-1,1] 连续位移
- quality.score：deck 全部 finding worth 均值 → sigmoid → 0-100
worth 语义：越高 = 越可能需要修（fixed 样本 worth > accepted 样本 worth）。
"""

from __future__ import annotations

import math

from offipy.audit import Severity

DELTA_THRESHOLD = 0.5
SEVERITY_SHIFT_LOW = -1.0
SEVERITY_SHIFT_HIGH = 1.0


def quantize_delta(worth: float) -> int:
    """rule.delta 量化：非有限值或 |worth| < DELTA_THRESHOLD → 0；否则四舍五入到 ±1。"""
    if not math.isfinite(worth) or abs(worth) < DELTA_THRESHOLD:
        return 0
    return 1 if worth > 0 else -1


def severity_shift_from_worth(worth: float) -> float:
    """finding.severity_shift：worth → clamp [-1, 1]（NaN 视为 0，不参与 shift）。"""
    if not math.isfinite(worth):
        return 0.0
    return max(SEVERITY_SHIFT_LOW, min(SEVERITY_SHIFT_HIGH, worth))


def quality_score_from_worth(worth: float, worth_scale: float = 1.0) -> float:
    """worth → plausibility score（0-100，保留 1 位小数）。NaN/inf → 0.0。

    worth_scale 归一避免饱和：worth 幅值被训练分布的 scale 拉伸后再进 sigmoid，
    高幅值/离群 worth 不再直接把分数压到 0 或 100。缺省 1.0 兼容旧测试。
    必须保留 round(..., 1)——head 与 analyze 测试按 1 位小数断言。
    """
    if not worth_scale or worth_scale <= 0:
        worth_scale = 1.0
    if not math.isfinite(worth):
        return 0.0
    return round(100.0 / (1.0 + math.exp(2.0 * (worth / worth_scale))), 1)


def apply_severity_shift(sev: Severity, shift: float) -> Severity:
    """把连续 shift 应用到 discrete Severity：|shift|≥0.5 才 ±1 级（夹回 LOW..HIGH）。

    #148：round() 银行家舍入在 x.5 处就近取偶，+0.5 与 -0.5 的移动方向不对称、
    可能非单调。改为与 quantize_delta 同阈值（DELTA_THRESHOLD=0.5）先量化出整数
    步（±1/0）再叠加——单调、方向明确，且与 rule.delta 的量化语义一致。
    """
    if not math.isfinite(shift):
        return sev
    step = 0
    if shift >= DELTA_THRESHOLD:
        step = 1
    elif shift <= -DELTA_THRESHOLD:
        step = -1
    return Severity(max(int(Severity.LOW), min(int(Severity.HIGH), int(sev) + step)))
