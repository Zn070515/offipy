"""PPTX 动画注入引擎（入场效果 + 页面过渡）。"""

from .spec import EFFECT_CATALOG, AnimationSpec, TransitionSpec

# apply.py 由 Task 5/6 落地。此前包仍须可导入（Task 2-4 直接 import spec 子模块，
# 而 Python 导入子模块必先执行包 __init__）——临时守卫：等 apply.py 存在后此
# except 永不触发，恢复成 plan 的目标状态即可（可顺手删掉这个 try/except）。
try:
    from .apply import apply_animations, apply_transitions
except ModuleNotFoundError as exc:  # pragma: no cover - apply.py 尚未创建
    if exc.name != "offipy.animations.apply":
        raise
    apply_animations = None  # type: ignore[assignment]
    apply_transitions = None  # type: ignore[assignment]

__all__ = [
    "EFFECT_CATALOG",
    "AnimationSpec",
    "TransitionSpec",
    "apply_animations",
    "apply_transitions",
]
