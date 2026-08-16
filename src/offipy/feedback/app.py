"""FeedbackApp：feedback 学习系统（schema app，train/status/append 三 op）。

纯本地纯 CPU，无 COM（has_com_root=False → server._alive 恒 True）。顶层只
标准库 + numpy-free 红线：numpy 只在 train()/status() 内部惰性 import（append
不需要 numpy），所以 base install（无 feedback extra）起 server 不崩。
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

    def recommend(
        self,
        pptx: str,
        feedback_dir: str | None = None,
        *,
        profile: str = "balanced",
        json: bool = False,
    ) -> dict[str, Any]:
        """只读建议：对 .pptx 跑分析 + 学习推理，投影确定性建议与调整 finding。

        无有效模型 → InvalidArgumentError（不回退 v2 静默推荐）。json 参数仅为
        CLI 表面兼容（generic dispatch 恒输出 JSON）；服务器侧忽略。
        """
        _ = json  # --json 仅兼容 CLI 表面，输出恒 JSON（generic dispatch）
        from offipy.art.analyze import analyze_deck
        from offipy.art.suggest import project_adjusted_findings, project_suggestions
        from offipy.feedback import infer

        if not feedback_dir:
            # 无参时 bundle.load(None) 读默认 ~/.offipy，但 analyze_deck 的 feedback=True
            # 强制显式 feedback_dir——两条路径不一致。先在这边显式报错，别等 analyze_scene
            # 抛一个语义模糊的「feedback=True 必须显式提供 feedback_dir」。
            raise InvalidArgumentError(
                "feedback recommend 需要显式指定 feedback_dir（只读建议必须绑定反馈库，"
                "不回退到默认目录）"
            )
        bundle = infer.ModelBundle.load(feedback_dir)
        if bundle is None:
            raise InvalidArgumentError(
                "feedback recommend 需要有效模型：请先 offipy feedback train"
                "（无模型/过期/损坏时回退 v2 不推荐）"
            )
        report = analyze_deck(
            pptx=pptx,
            profile=profile,
            include_experimental_score=True,
            feedback=True,
            feedback_dir=feedback_dir,
        )
        art = report.art
        return {
            "source": pptx,
            "profile": profile,
            "experimental_score": art.experimental_score if art else None,
            "experimental_score_mode": art.experimental_score_mode if art else None,
            "quality_score_coverage": art.quality_score_coverage if art else None,
            "feedback_adjustments": dict(art.feedback_adjustments) if art else {},
            "adjusted_findings": project_adjusted_findings(report),
            "suggestions": project_suggestions(report, source=pptx),
            "warnings": [{"code": w.code, "message": w.message} for w in report.warnings],
        }

    def apply(self, profile: str, feedback_dir: str | None = None) -> dict[str, Any]:
        """把学习到的 rule.delta 持久化到 profile 存储（默认 ~/.offipy/art_profiles.json）。

        之后 `deck audit --profile <name>`（不带 --feedback-dir）也会吃到调整。
        无有效模型 → InvalidArgumentError。
        """
        from offipy.art.profiles import (
            PROFILE_STORE_DIR,
            PROFILE_STORE_FILE,
            get_profile,
            load_persisted_adjustments,
            save_persisted_adjustments,
        )
        from offipy.feedback.infer import learned_adjustments

        get_profile(profile)  # 未知 profile → InvalidArgumentError（含友好消息）
        adjustments = learned_adjustments(profile, feedback_dir=feedback_dir)
        if adjustments is None:
            raise InvalidArgumentError(
                "feedback apply 需要有效模型：无模型/过期/损坏时无 rule.delta 可持久化"
            )
        merged = load_persisted_adjustments()
        current = dict(merged.get(profile, {}))
        current.update(adjustments)
        merged[profile] = current
        store = save_persisted_adjustments(merged)
        return {
            "profile": profile,
            "adjustments": adjustments,
            "store": str(store),
            "store_dir": str(PROFILE_STORE_DIR),
            "store_file": PROFILE_STORE_FILE,
        }

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

        try:
            sev = severity if isinstance(severity, Severity) else Severity[str(severity)]
        except KeyError:
            raise InvalidArgumentError(
                f"severity 必须是 LOW/MID/HIGH 之一，得到 {severity!r}"
            ) from None
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
        # #143：有特征快照但缺 schema 版本 → 自动补当前版本，避免 valid_records 静默过滤。
        # 插在 feats 解析之后、art_append 调用之前；原有调用不变。
        if feats is not None and feature_schema_version is None:
            from offipy.art.features_registry import feature_schema_version as _current

            feature_schema_version = _current()
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
