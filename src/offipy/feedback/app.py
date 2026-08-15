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
