"""反馈学习（P2 验证版）：记录审计后的人工修正，反哺审计权重。

对标 AeSlides/EvoPresent 的「可验证奖励」思路（docs/ppt_design_research.md
P2），轻量验证版：审计报告出来后，Claude/用户对每条 finding 处置
（修正 fixed / 接受 accepted / 忽略 ignored），写入 `~/.offipy/feedback.jsonl`。

dimension_weights() 把历史修正计数折算成审计权重——被修次数越多的维度，
下次审计扣分越狠，促使 Claude 更早重视。weights 可直接传给
aesthetic.audit_measurement(..., weights=weights)。这是纯本地的 RL 验证版：
数据不外传，逐步逼近人眼的判断。
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from .aesthetic import ALL_DIMENSIONS, CONSISTENCY, CONTRAST, PALETTE, TYPE_SCALE, WHITESPACE

__all__ = [
    "ALL_DIMENSIONS",
    "CONSISTENCY",
    "CONTRAST",
    "DEFAULT_DIR",
    "FEEDBACK_FILE",
    "PALETTE",
    "TYPE_SCALE",
    "VALID_ACTIONS",
    "WHITESPACE",
    "FeedbackRecord",
    "append",
    "dimension_weights",
    "load_records",
    "record_file",
]

# 记录文件默认位置：~/.offipy/feedback.jsonl
DEFAULT_DIR = Path.home() / ".offipy"
FEEDBACK_FILE = "feedback.jsonl"

VALID_ACTIONS = ("fixed", "accepted", "ignored")

# 权重折算：fixed 每条 +0.5，accepted 每条 -0.25（可抵消一次误报）
_FIXED_STEP = 0.5
_ACCEPTED_STEP = -0.25
_WEIGHT_MAX = 3.0
_WEIGHT_MIN = 0.5


@dataclass(frozen=True)
class FeedbackRecord:
    ts: str
    dimension: str
    severity: str
    page: int
    message: str
    action: str
    source: str = ""

    def to_dict(self) -> dict:
        return {
            "ts": self.ts,
            "dimension": self.dimension,
            "severity": self.severity,
            "page": self.page,
            "message": self.message,
            "action": self.action,
            **({"source": self.source} if self.source else {}),
        }

    @classmethod
    def from_dict(cls, data: dict) -> FeedbackRecord:
        return cls(
            ts=data["ts"],
            dimension=data["dimension"],
            severity=data["severity"],
            page=int(data.get("page", 0)),
            message=data.get("message", ""),
            action=data["action"],
            source=data.get("source", ""),
        )


def record_file(feedback_dir: str | Path | None = None) -> Path:
    """反馈记录文件路径（默认 ~/.offipy/feedback.jsonl）。"""
    d = Path(feedback_dir) if feedback_dir else DEFAULT_DIR
    return d / FEEDBACK_FILE


def _validate(dimension: str, action: str) -> None:
    if dimension not in ALL_DIMENSIONS:
        raise ValueError(f"未知维度 {dimension!r}（可选: {', '.join(ALL_DIMENSIONS)}）")
    if action not in VALID_ACTIONS:
        raise ValueError(f"未知处置 {action!r}（可选: {', '.join(VALID_ACTIONS)}）")


def append(
    dimension: str,
    action: str,
    severity: str = "MID",
    page: int = 0,
    message: str = "",
    source: str = "",
    feedback_dir: str | Path | None = None,
    ts: str | None = None,
) -> Path:
    """追加一条反馈记录，返回记录文件路径。目录不存在则创建。"""
    _validate(dimension, action)
    rec = FeedbackRecord(
        ts=ts or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        dimension=dimension,
        severity=severity,
        page=page,
        message=message,
        action=action,
        source=source,
    )
    f = record_file(feedback_dir)
    f.parent.mkdir(parents=True, exist_ok=True)
    with f.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(rec.to_dict(), ensure_ascii=False) + "\n")
    return f


def load_records(feedback_dir: str | Path | None = None) -> list[FeedbackRecord]:
    """读取全部反馈记录（文件不存在返回空列表，坏行跳过）。"""
    f = record_file(feedback_dir)
    if not f.exists():
        return []
    records = []
    for line in f.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            rec = FeedbackRecord.from_dict(json.loads(line))
            _validate(rec.dimension, rec.action)
            records.append(rec)
        except (json.JSONDecodeError, KeyError, TypeError, ValueError):
            continue  # 跳过坏行（含非数字 page / 未知维度 / 未知处置），不因一条脏数据破坏整体
    return records


def dimension_weights(
    feedback_dir: str | Path | None = None,
    base: dict[str, float] | None = None,
) -> dict[str, float]:
    """按历史反馈折算审计权重 → {dimension: weight}。

    base 给默认权重（缺省全 1.0）；fixed 每条 +0.5、accepted 每条 -0.25，
    封顶 [_WEIGHT_MIN, _WEIGHT_MAX]。无记录的维度保持 base 值。
    """
    base = (
        {d: float(base.get(d, 1.0)) for d in ALL_DIMENSIONS}
        if base
        else dict.fromkeys(ALL_DIMENSIONS, 1.0)
    )
    for rec in load_records(feedback_dir):
        if rec.action == "fixed":
            base[rec.dimension] = min(_WEIGHT_MAX, base[rec.dimension] + _FIXED_STEP)
        elif rec.action == "accepted":
            base[rec.dimension] = max(_WEIGHT_MIN, base[rec.dimension] + _ACCEPTED_STEP)
        # ignored：不影响权重
    return base
