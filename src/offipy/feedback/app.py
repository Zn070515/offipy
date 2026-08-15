"""FeedbackApp：feedback 学习系统（schema app，train/status 两 op）。

纯本地纯 CPU，无 COM（has_com_root=False → server._alive 恒 True）。顶层只
标准库 + numpy-free 红线：numpy 只在 train()/status() 内部惰性 import，所以
base install（无 feedback extra）起 server 不崩。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from offipy.exceptions import InvalidArgumentError


class FeedbackApp:
    has_com_root = False

    def train(self, feedback_dir: str | None = None, *, seed: int = 42) -> dict[str, Any]:
        try:
            from .train import run_training
        except ImportError as exc:
            raise RuntimeError('反馈学习需要 numpy：pip install "offipy[feedback]"') from exc
        if feedback_dir is not None and not Path(feedback_dir).is_dir():
            raise InvalidArgumentError(f"feedback_dir 必须是目录: {feedback_dir}")
        return run_training(feedback_dir, seed=seed)

    def status(self, feedback_dir: str | None = None) -> dict[str, Any]:
        from .status import report_status

        return report_status(feedback_dir)

    def append(
        self,
        profile: str,
        rule_id: str,
        action: str,
        severity: str,
        *,
        slide_index: int | None = None,
        message: str = "",
        source: str = "",
        feedback_dir: str | None = None,
        ts: str | None = None,
        features: dict[str, float] | None = None,
        feature_schema_version: str | None = None,
    ) -> dict[str, Any]:
        """追加一条反馈标签记录（fixed/accepted/ignored），返回记录文件路径。

        severity 接受 Severity 或名字字符串（"HIGH"）；features 接受扁平 dict
        （Python API）或 JSON 字符串（CLI --features）。
        """
        import json

        from offipy.art.feedback import append as art_append
        from offipy.audit import Severity

        sev = severity if isinstance(severity, Severity) else Severity[str(severity)]
        feats = features
        if isinstance(feats, str):
            try:
                feats = json.loads(feats)
            except (ValueError, TypeError) as exc:
                raise InvalidArgumentError(
                    "features 必须是 JSON 对象（扁平特征 dict，如 "
                    '{"finding.confidence": 0.5}），得到: '
                    f"{feats!r}"
                ) from exc
        path = art_append(
            profile,
            rule_id,
            action,
            sev,
            slide_index=slide_index,
            message=message,
            source=source,
            feedback_dir=feedback_dir,
            ts=ts,
            features=feats,
            feature_schema_version=feature_schema_version,
        )
        return {"record": str(path)}
