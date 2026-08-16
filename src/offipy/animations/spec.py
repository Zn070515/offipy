"""动画声明归一化：显式 data-ppt-* 优先，约定回退其次，未知/畸形值跳过 + 告警。

vendor measure 只收集原始声明（anim_decl dict），所有映射/校验/告警语义都在
这里（offipy 侧）。bbox 替换型 record（asset/mermaid/drawio）不产动画。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from offipy.exceptions import InvalidArgumentError

# curated 效果目录（spec §效果与触发模型）
EFFECT_CATALOG = {"fade", "float_up", "fly_in", "wipe", "zoom_in", "grow"}

# direction 只对 fly_in/wipe 有意义，取值 bottom|left|right|top
_DIRECTIONAL_EFFECTS = {"fly_in", "wipe"}
_DIRECTIONS = {"bottom", "left", "right", "top"}
_TRIGGERS = {"click", "after"}

# 约定回退表：data-anim/data-aos 值 → (effect, direction)
# 未列出的值（flip、slide-up、fade-up-right…）不映射 → 跳过
_FALLBACK_MAP = {
    "fade": ("fade", None),
    "fade-up": ("float_up", None),
    "fade-left": ("fly_in", "left"),
    "fade-right": ("fly_in", "right"),
    "fade-down": ("fly_in", "top"),
    "zoom-in": ("zoom_in", None),
}

# bbox 替换型 record 标记：占位形状会被后处理整槽替换，锚点丢失 → 不产动画
_BBOX_REPLACEMENT_KINDS = {"asset"}
_BBOX_REPLACEMENT_CLASS_TOKENS = ("mermaid", "drawio")

_DEFAULT_DURATION = 0.5
_DEFAULT_DIRECTION = "bottom"


@dataclass
class ParsedAnimation:
    """parse_declaration 的归一化结果：不含 slide/target（由调用方组装 AnimationSpec）。

    独立于 AnimationSpec 存在：AnimationSpec 的 __post_init__ 强校验 slide≥1、
    target 非空（独立 API 规格纪律），而声明解析阶段还没有这两者——若复用
    AnimationSpec 会误抛 InvalidArgumentError。
    """

    effect: str
    direction: str
    trigger: str
    duration: float
    delay: float


@dataclass
class AnimationSpec:
    """独立 API 的动画声明（slide 为 1-based；target 为形状名）。"""

    slide: int
    target: str
    effect: str
    direction: str = _DEFAULT_DIRECTION
    trigger: str = "click"
    duration: float = _DEFAULT_DURATION
    delay: float = 0.0
    warnings: list[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        problems: list[str] = []
        if not isinstance(self.slide, int) or self.slide < 1:
            problems.append(f"slide 必须是 ≥1 的整数（实际 {self.slide!r}）")
        if not isinstance(self.target, str) or not self.target.strip():
            problems.append("target 必须是形状名（非空字符串）")
        if self.effect not in EFFECT_CATALOG:
            problems.append(f"effect 不在目录 {sorted(EFFECT_CATALOG)}（实际 {self.effect!r}）")
        if self.trigger not in _TRIGGERS:
            problems.append(f"trigger 必须是 {'/'.join(_TRIGGERS)}（实际 {self.trigger!r}）")
        if self.effect in _DIRECTIONAL_EFFECTS and self.direction not in _DIRECTIONS:
            allowed = "/".join(_DIRECTIONS)
            problems.append(
                f"direction 对 {self.effect} 必须是 {allowed}（实际 {self.direction!r}）"
            )
        if not (isinstance(self.duration, (int, float)) and self.duration > 0):
            problems.append(f"duration 必须是正数秒（实际 {self.duration!r}）")
        if not (isinstance(self.delay, (int, float)) and self.delay >= 0):
            problems.append(f"delay 必须是非负数秒（实际 {self.delay!r}）")
        if problems:
            raise InvalidArgumentError("动画声明非法：" + "；".join(problems))


@dataclass
class TransitionSpec:
    """页面过渡声明（slide 1-based）。kind 只用 classic OOXML 过渡。"""

    slide: int
    kind: str
    speed: str = "medium"

    def __post_init__(self) -> None:
        problems: list[str] = []
        if not isinstance(self.slide, int) or self.slide < 1:
            problems.append(f"slide 必须是 ≥1 的整数（实际 {self.slide!r}）")
        if self.kind not in {"fade", "wipe", "push", "cover"}:
            problems.append(f"kind 必须是 fade/wipe/push/cover（实际 {self.kind!r}）")
        if self.speed not in {"slow", "medium", "fast"}:
            problems.append(f"speed 必须是 slow/medium/fast（实际 {self.speed!r}）")
        if problems:
            raise InvalidArgumentError("过渡声明非法：" + "；".join(problems))


def _is_bbox_replacement(record: dict) -> bool:
    kind = record.get("kind")
    if kind in _BBOX_REPLACEMENT_KINDS:
        return True
    cls = record.get("className") or ""
    return any(tok in cls for tok in _BBOX_REPLACEMENT_CLASS_TOKENS)


def parse_declaration(raw: dict | None, record: dict | None) -> ParsedAnimation | None:
    """把 measure 收集的原始 anim_decl 归一化为 ParsedAnimation（无 slide/target）。

    返回 None 表示「不产动画」（未命中/未知值/bbox 替换型），告警写进 raw['_warnings']。
    record 是产生该声明的 record（供 bbox 型过滤）。slide/target 由调用方
    （postprocess 按页与 OFFIPY_ELEM 锚）组装成 AnimationSpec。
    """
    if not raw:
        return None
    record = record or {}
    warnings: list[str] = []
    if _is_bbox_replacement(record):
        warnings.append("bbox 替换型元素（asset/mermaid/drawio）不支持动画，已跳过")
        return _none_with(raw, warnings)

    effect: str | None = None
    direction: str | None = None
    trigger: str = "click"
    duration: float = _DEFAULT_DURATION
    delay: float = 0.0

    if raw.get("anim"):  # 显式 data-ppt-anim 优先
        effect = str(raw["anim"])
        if raw.get("dir"):
            direction = str(raw["dir"])
        if raw.get("trigger"):
            trigger = str(raw["trigger"])
        dur = raw.get("dur")
        if dur is not None:
            try:
                duration = float(dur)
            except (TypeError, ValueError):
                warnings.append(f"data-ppt-dur={dur!r} 非法，忽略整条动画声明")
                return _none_with(raw, warnings)
        dly = raw.get("delay")
        if dly is not None:
            try:
                delay = float(dly)
            except (TypeError, ValueError):
                warnings.append(f"data-ppt-delay={dly!r} 非法，忽略整条动画声明")
                return _none_with(raw, warnings)
    else:  # 约定回退
        legacy = raw.get("dataAnim") or raw.get("dataAos")
        if legacy:
            mapped = _FALLBACK_MAP.get(str(legacy))
            if mapped is None:
                warnings.append(f"未列入动画约定的值 {legacy!r}，跳过")
                return _none_with(raw, warnings)
            effect, direction = mapped
        elif raw.get("fadeIn") or raw.get("animateIn"):
            effect, direction = "fade", None
        else:
            return None

    if effect not in EFFECT_CATALOG:
        warnings.append(f"未知动画效果 {effect!r}，跳过")
        return _none_with(raw, warnings)
    if effect in _DIRECTIONAL_EFFECTS:
        if direction is not None and direction not in _DIRECTIONS:
            warnings.append(f"direction {direction!r} 对 {effect} 非法，跳过")
            return _none_with(raw, warnings)
    else:
        direction = _DEFAULT_DIRECTION
    if trigger not in _TRIGGERS:
        warnings.append(f"trigger {trigger!r} 非法，跳过")
        return _none_with(raw, warnings)
    if duration <= 0:
        warnings.append(f"duration {duration} 必须为正，跳过")
        return _none_with(raw, warnings)
    if delay < 0:
        warnings.append(f"delay {delay} 必须非负，跳过")
        return _none_with(raw, warnings)

    return ParsedAnimation(
        effect=effect,
        direction=direction or _DEFAULT_DIRECTION,
        trigger=trigger,
        duration=duration,
        delay=delay,
    )


def _none_with(raw: dict, warnings: list[str]) -> None:
    raw["_warnings"] = (raw.get("_warnings") or []) + warnings
    return
