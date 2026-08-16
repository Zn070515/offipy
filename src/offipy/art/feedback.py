"""规则级反馈学习（v2）：记录用户对每条艺术 finding 的处置，聚合为有界的严重度调整。

v1 的 `~/.offipy/feedback.jsonl`（offipy.feedback）是维度级权重；v2 下沉到规则级：
每条记录携带 (profile, rule_id, action)，`recommend_adjustments` 按 profile 过滤、
按 rule_id 分组，产出有界的 −1/+1 严重度调整；`apply_feedback` 把调整落到
`ArtProfile.feedback_severity_adjustments`（frozen dataclass 经 dataclasses.replace
新建实例，绝不改共享的内置 profile）。数据纯本地，不外传。
"""

from __future__ import annotations

import dataclasses
import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal, cast

from offipy.art.profiles import RULE_DIMENSIONS, ArtProfile, get_profile
from offipy.audit import Severity
from offipy.exceptions import InvalidArgumentError

__all__ = [
    "ART_FEEDBACK_FILE",
    "DEFAULT_DIR",
    "VALID_ACTIONS",
    "ArtFeedbackRecord",
    "append",
    "apply_feedback",
    "load_records",
    "recommend_adjustments",
    "record_file",
    "reschema_records",
]

# 记录文件默认位置：~/.offipy/art_feedback.jsonl（与 v1 的 feedback.jsonl 分离）
DEFAULT_DIR = Path.home() / ".offipy"
ART_FEEDBACK_FILE = "art_feedback.jsonl"

VALID_ACTIONS = ("fixed", "accepted", "ignored")
FeedbackAction = Literal["fixed", "accepted", "ignored"]

# 聚合阈值：fixed+accepted 不足 3 条不下结论；单一处置占比 >= 0.6 才给建议
_MIN_SAMPLES = 3
_MAJORITY = 0.6


@dataclass(frozen=True)
class ArtFeedbackRecord:
    ts: str
    profile: str
    rule_id: str
    dimension: str
    severity: Severity
    action: FeedbackAction
    slide_index: int | None = None
    message: str = ""
    source: str = ""
    features: dict[str, float] | None = None  # {feature_id: 标量} 扁平快照（encode_features 产出）
    feature_schema_version: str | None = None  # 训练样本的 FEATURES schema 版本

    def to_dict(self) -> dict[str, Any]:
        return {
            "ts": self.ts,
            "profile": self.profile,
            "rule_id": self.rule_id,
            "dimension": self.dimension,
            "severity": self.severity.name,
            "action": self.action,
            "slide_index": self.slide_index,
            "message": self.message,
            "source": self.source,
            **({"features": self.features} if self.features is not None else {}),
            **(
                {"feature_schema_version": self.feature_schema_version}
                if self.feature_schema_version is not None
                else {}
            ),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ArtFeedbackRecord:
        rule_id = data["rule_id"]
        if rule_id not in RULE_DIMENSIONS:
            raise ValueError(f"未知 rule_id {rule_id!r}")
        action = data["action"]
        if action not in VALID_ACTIONS:
            raise ValueError(f"未知 action {action!r}")
        slide_index = data.get("slide_index")
        if slide_index is not None and (
            not isinstance(slide_index, int) or isinstance(slide_index, bool) or slide_index < 1
        ):
            raise ValueError(f"非法 slide_index {slide_index!r}")
        return cls(
            ts=data["ts"],
            profile=data["profile"],
            rule_id=rule_id,
            dimension=data["dimension"],
            severity=_coerce_severity(data["severity"]),
            action=cast("FeedbackAction", action),
            slide_index=slide_index,
            message=data.get("message", ""),
            source=data.get("source", ""),
            features=data.get("features"),
            feature_schema_version=data.get("feature_schema_version"),
        )


def _coerce_severity(raw: str | int) -> Severity:
    """severity 兼容两种表示：名字字符串（"HIGH"，主）与 int 值（3，容错）。"""
    if isinstance(raw, str):
        return Severity[raw]
    return Severity(raw)


def record_file(feedback_dir: str | Path | None = None) -> Path:
    """反馈记录文件路径（默认 ~/.offipy/art_feedback.jsonl）。"""
    d = Path(feedback_dir) if feedback_dir else DEFAULT_DIR
    return d / ART_FEEDBACK_FILE


def _validate(
    profile: str, rule_id: str, action: str, severity: Severity, slide_index: int | None
) -> None:
    if not isinstance(profile, str) or not profile:
        raise InvalidArgumentError("profile 必须是非空字符串")
    if rule_id not in RULE_DIMENSIONS:
        raise InvalidArgumentError(
            f"未知 rule_id {rule_id!r}（可选: {', '.join(sorted(RULE_DIMENSIONS))}）"
        )
    if action not in VALID_ACTIONS:
        raise InvalidArgumentError(f"未知 action {action!r}（可选: {', '.join(VALID_ACTIONS)}）")
    if not isinstance(severity, Severity):
        raise InvalidArgumentError(f"severity 必须是 Severity 成员，得到 {severity!r}")
    if slide_index is not None and (
        not isinstance(slide_index, int) or isinstance(slide_index, bool) or slide_index < 1
    ):
        raise InvalidArgumentError(f"slide_index 必须是 None 或 >= 1 的整数，得到 {slide_index!r}")


def append(
    profile: str,
    rule_id: str,
    action: str,
    severity: Severity,
    *,
    slide_index: int | None = None,
    message: str = "",
    source: str = "",
    feedback_dir: str | Path | None = None,
    ts: str | None = None,
    features: dict[str, float] | None = None,
    feature_schema_version: str | None = None,
) -> Path:
    """追加一条规则级反馈记录，返回记录文件路径。目录不存在则创建。

    features 须为扁平标量 dict（{feature_id: float}），需可 JSON 序列化——
    仅由 encode_features 产出，不要传入 numpy 标量等非原生类型（会炸 json.dumps）。
    feature_schema_version 记录该快照对应的 FEATURES schema 版本。
    """
    _validate(profile, rule_id, action, severity, slide_index)
    rec = ArtFeedbackRecord(
        ts=ts or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        profile=profile,
        rule_id=rule_id,
        dimension=RULE_DIMENSIONS[rule_id],
        severity=severity,
        action=cast("FeedbackAction", action),
        slide_index=slide_index,
        message=message,
        source=source,
        features=features,
        feature_schema_version=feature_schema_version,
    )
    f = record_file(feedback_dir)
    f.parent.mkdir(parents=True, exist_ok=True)
    with f.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(rec.to_dict(), ensure_ascii=False) + "\n")
    return f


def load_records(feedback_dir: str | Path | None = None) -> list[ArtFeedbackRecord]:
    """读取全部规则级反馈记录（文件不存在返回空列表，坏行跳过）。"""
    f = record_file(feedback_dir)
    if not f.exists():
        return []
    records = []
    for line in f.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            records.append(ArtFeedbackRecord.from_dict(json.loads(line)))
        except (json.JSONDecodeError, KeyError, TypeError, ValueError):
            continue  # 跳过坏行，不因一条脏数据破坏整体
    return records


def recommend_adjustments(
    profile: str, *, feedback_dir: str | Path | None = None
) -> dict[str, int]:
    """按 (profile, rule_id) 聚合历史处置，产出有界的严重度调整 {rule_id: -1|+1}。

    ignored 不参与分母；fixed+accepted 不足 3 条不推荐；
    某类处置占比 >= 0.6 时给出对应调整（fixed → +1，accepted → −1）。
    """
    counts: dict[str, list[int]] = {}
    for rec in load_records(feedback_dir):
        if rec.profile != profile:
            continue
        bucket = counts.setdefault(rec.rule_id, [0, 0])
        if rec.action == "fixed":
            bucket[0] += 1
        elif rec.action == "accepted":
            bucket[1] += 1
        # ignored：不参与聚合
    result: dict[str, int] = {}
    for rule_id, (fixed, accepted) in counts.items():
        den = fixed + accepted
        if den < _MIN_SAMPLES:
            continue
        if fixed / den >= _MAJORITY:
            result[rule_id] = 1
        elif accepted / den >= _MAJORITY:
            result[rule_id] = -1
    return result


def apply_feedback(
    profile: str | ArtProfile,
    *,
    feedback_dir: str | Path | None = None,
) -> ArtProfile:
    """把反馈调整应用到 profile：返回携带 feedback_severity_adjustments 的新 ArtProfile。

    str 走 get_profile 解析；ArtProfile 直接用（保留全部用户字段）。
    经 dataclasses.replace 新建实例，不改动输入 profile 或 _BUILTIN 共享对象。
    """
    target = get_profile(profile) if isinstance(profile, str) else profile
    adjustments = recommend_adjustments(target.name, feedback_dir=feedback_dir)
    return dataclasses.replace(target, feedback_severity_adjustments=adjustments)


def reschema_records(feedback_dir: str | Path | None = None) -> dict[str, int]:
    """把 schema 过期但有 features 的记录重写为当前 feature_schema_version。

    #144：feature_schema_version bump 后旧记录被 valid_records 过滤。本函数把
    仍有特征快照的过期记录原地重写（features dict 保留，缺失 key 由 encode_vector
    补 0），返回 {rewritten, skipped_no_features, already_current}。无 features 的
    记录无法重编码，跳过并计数。
    """
    from offipy.art.features_registry import feature_schema_version

    current = feature_schema_version()
    f = record_file(feedback_dir)
    if not f.exists():
        return {"rewritten": 0, "skipped_no_features": 0, "already_current": 0}
    lines = f.read_text(encoding="utf-8").splitlines()
    rewritten = skipped = already = 0
    new_lines: list[str] = []
    for line in lines:
        if not line.strip():
            continue
        try:
            rec = ArtFeedbackRecord.from_dict(json.loads(line))
        except (json.JSONDecodeError, KeyError, TypeError, ValueError):
            new_lines.append(line)  # 坏行保留，不破坏文件
            continue
        if rec.feature_schema_version == current:
            already += 1
            new_lines.append(line)
        elif rec.features is not None:
            rec = dataclasses.replace(rec, feature_schema_version=current)
            rewritten += 1
            new_lines.append(json.dumps(rec.to_dict(), ensure_ascii=False))
        else:
            skipped += 1
            new_lines.append(line)
    if rewritten:
        f.write_text("\n".join(new_lines) + ("\n" if new_lines else ""), encoding="utf-8")
    return {"rewritten": rewritten, "skipped_no_features": skipped, "already_current": already}
