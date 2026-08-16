"""PPTX 动画注入引擎（入场效果 + 页面过渡）。"""

from .apply import apply_animations, apply_transitions
from .spec import EFFECT_CATALOG, AnimationSpec, TransitionSpec

__all__ = [
    "EFFECT_CATALOG",
    "AnimationSpec",
    "TransitionSpec",
    "apply_animations",
    "apply_transitions",
]
