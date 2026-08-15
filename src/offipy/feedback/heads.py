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
    """rule.delta 量化：|worth| < DELTA_THRESHOLD → 0；否则四舍五入到 ±1。"""
    if abs(worth) < DELTA_THRESHOLD:
        return 0
    return 1 if worth > 0 else -1


def severity_shift_from_worth(worth: float) -> float:
    """finding.severity_shift：worth → clamp [-1, 1]。"""
    return max(SEVERITY_SHIFT_LOW, min(SEVERITY_SHIFT_HIGH, worth))


def quality_score_from_worth(worth: float, worth_scale: float = 1.0) -> float:
    """worth → plausibility score（0-100，保留 1 位小数）。

    worth_scale 归一避免饱和：worth 幅值被训练分布的 scale 拉伸后再进 sigmoid，
    高幅值/离群 worth 不再直接把分数压到 0 或 100。缺省 1.0 兼容旧测试。
    必须保留 round(..., 1)——head 与 analyze 测试按 1 位小数断言。
    """
    if not worth_scale or worth_scale <= 0:
        worth_scale = 1.0
    return round(100.0 / (1.0 + math.exp(2.0 * (worth / worth_scale))), 1)


def apply_severity_shift(sev: Severity, shift: float) -> Severity:
    """把连续 shift 应用到 discrete Severity（round 后夹回 LOW..HIGH）。"""
    moved = round(float(sev) + shift)
    return Severity(max(int(Severity.LOW), min(int(Severity.HIGH), moved)))
