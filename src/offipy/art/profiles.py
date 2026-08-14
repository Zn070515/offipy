"""规则常量 + 内置 profile（v2：阈值 + 开关 + 覆盖）。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from offipy.exceptions import InvalidArgumentError

if TYPE_CHECKING:
    from collections.abc import Mapping

    from offipy.audit import Severity

# ---- 规则 ID 常量（canonical，勿改名）----
RULE_NO_FOCUS = "art.hierarchy.no_focus"
RULE_TITLE_TOO_SMALL = "art.hierarchy.title_too_small"
RULE_OFF_BALANCE = "art.composition.off_balance"
RULE_CORNER_CLUSTER = "art.composition.corner_cluster"
RULE_SPACING_DRIFT = "art.composition.spacing_drift"
RULE_MANY_FAMILIES = "art.typography.many_families"
RULE_TINY_TEXT = "art.typography.tiny_text"
RULE_FLAT_SCALE = "art.typography.flat_scale"
RULE_LOW_CONTRAST = "art.color.low_contrast"
RULE_ACCENT_FLOOD = "art.color.accent_flood"
RULE_NO_ACCENT = "art.color.no_accent"
RULE_DISTORTED_IMAGE = "art.media.distorted_image"
RULE_TINY_IMAGE = "art.media.tiny_image"
RULE_MIXED_IMAGE_SIZES = "art.media.mixed_image_sizes"
RULE_TITLE_DRIFT = "art.consistency.title_drift"
RULE_MARGIN_DRIFT = "art.consistency.margin_drift"
RULE_BACKGROUND_LIKE_AREA = "art.composition.background_like_area"

ALL_RULES = frozenset(
    {
        RULE_NO_FOCUS,
        RULE_TITLE_TOO_SMALL,
        RULE_OFF_BALANCE,
        RULE_CORNER_CLUSTER,
        RULE_SPACING_DRIFT,
        RULE_MANY_FAMILIES,
        RULE_TINY_TEXT,
        RULE_FLAT_SCALE,
        RULE_LOW_CONTRAST,
        RULE_ACCENT_FLOOD,
        RULE_NO_ACCENT,
        RULE_DISTORTED_IMAGE,
        RULE_TINY_IMAGE,
        RULE_MIXED_IMAGE_SIZES,
        RULE_TITLE_DRIFT,
        RULE_MARGIN_DRIFT,
        RULE_BACKGROUND_LIKE_AREA,
    }
)

# 规则 → 维度 规范化注册表：feedback 只依赖此处，不 import 规则实现
RULE_DIMENSIONS: Mapping[str, str] = {
    RULE_NO_FOCUS: "hierarchy",
    RULE_TITLE_TOO_SMALL: "hierarchy",
    RULE_OFF_BALANCE: "composition",
    RULE_CORNER_CLUSTER: "composition",
    RULE_SPACING_DRIFT: "composition",
    RULE_BACKGROUND_LIKE_AREA: "composition",
    RULE_MANY_FAMILIES: "typography",
    RULE_TINY_TEXT: "typography",
    RULE_FLAT_SCALE: "typography",
    RULE_LOW_CONTRAST: "color",
    RULE_ACCENT_FLOOD: "color",
    RULE_NO_ACCENT: "color",
    RULE_DISTORTED_IMAGE: "media",
    RULE_TINY_IMAGE: "media",
    RULE_MIXED_IMAGE_SIZES: "media",
    RULE_TITLE_DRIFT: "consistency",
    RULE_MARGIN_DRIFT: "consistency",
}

_EXPERIMENTAL = frozenset(
    {
        RULE_OFF_BALANCE,
        RULE_CORNER_CLUSTER,
        RULE_NO_FOCUS,
        RULE_ACCENT_FLOOD,
        RULE_NO_ACCENT,
        RULE_BACKGROUND_LIKE_AREA,
    }
)


@dataclass(frozen=True)
class ArtProfile:
    name: str
    min_text_size_norm: float = 0.015
    min_contrast: float = 3.0
    max_font_families: int = 3
    flat_scale_ratio_min: float = 1.3
    max_accent_ratio: float = 0.35
    balance_tol: float = 0.25
    corner_cluster_ratio: float = 0.45
    spacing_drift_tol: float = 0.10
    title_size_min_norm: float = 0.03
    max_image_aspect_drift: float = 0.15
    min_image_area: float = 0.0025
    max_image_size_spread: float = 0.5
    title_drift_tol: float = 0.15
    margin_drift_tol: float = 0.10
    max_background_like_ratio: float = 0.75
    enabled_rules: frozenset[str] = ALL_RULES
    disabled_rules: frozenset[str] = frozenset()
    severity_overrides: Mapping[str, Severity] = field(default_factory=dict)
    feedback_severity_adjustments: Mapping[str, int] = field(default_factory=dict)
    confidence_overrides: dict[str, float] = field(default_factory=dict)
    experimental_rules: frozenset[str] = _EXPERIMENTAL

    def __post_init__(self) -> None:
        for rule_id, delta in self.feedback_severity_adjustments.items():
            if delta not in (-1, 1):
                raise InvalidArgumentError(
                    f"feedback_severity_adjustments value for {rule_id!r} "
                    f"must be -1 or +1, got {delta}"
                )


_BUILTIN = {
    "balanced": ArtProfile(
        name="balanced",
        min_text_size_norm=0.015,
        min_contrast=3.0,
        max_font_families=3,
        flat_scale_ratio_min=1.3,
        max_accent_ratio=0.35,
        balance_tol=0.25,
        corner_cluster_ratio=0.45,
        spacing_drift_tol=0.10,
        title_size_min_norm=0.03,
        max_image_aspect_drift=0.15,
        min_image_area=0.0025,
        max_image_size_spread=0.5,
        title_drift_tol=0.15,
        margin_drift_tol=0.10,
    ),
    "consulting": ArtProfile(
        name="consulting",
        min_text_size_norm=0.016,
        min_contrast=4.5,
        max_font_families=2,
        flat_scale_ratio_min=1.4,
        max_accent_ratio=0.25,
        balance_tol=0.20,
        corner_cluster_ratio=0.40,
        spacing_drift_tol=0.20,
        title_size_min_norm=0.032,
        max_image_aspect_drift=0.10,
        min_image_area=0.003,
        max_image_size_spread=0.4,
        title_drift_tol=0.10,
        margin_drift_tol=0.08,
        max_background_like_ratio=0.85,
    ),
    "academic": ArtProfile(
        name="academic",
        min_text_size_norm=0.013,
        min_contrast=4.5,
        max_font_families=4,
        flat_scale_ratio_min=1.2,
        max_accent_ratio=0.2,
        balance_tol=0.35,
        corner_cluster_ratio=0.55,
        spacing_drift_tol=0.15,
        title_size_min_norm=0.028,
        max_image_aspect_drift=0.2,
        min_image_area=0.002,
        max_image_size_spread=0.6,
        title_drift_tol=0.12,
        margin_drift_tol=0.12,
        disabled_rules=frozenset({RULE_OFF_BALANCE}),
    ),
    "technology": ArtProfile(
        name="technology",
        min_text_size_norm=0.014,
        min_contrast=3.5,
        max_font_families=3,
        flat_scale_ratio_min=1.35,
        max_accent_ratio=0.5,
        balance_tol=0.3,
        corner_cluster_ratio=0.5,
        spacing_drift_tol=0.12,
        title_size_min_norm=0.03,
        max_image_aspect_drift=0.15,
        min_image_area=0.0025,
        max_image_size_spread=0.5,
        title_drift_tol=0.15,
        margin_drift_tol=0.1,
    ),
    "event": ArtProfile(
        name="event",
        min_text_size_norm=0.02,
        min_contrast=2.5,
        max_font_families=4,
        flat_scale_ratio_min=1.5,
        max_accent_ratio=0.7,
        balance_tol=0.4,
        corner_cluster_ratio=0.7,
        spacing_drift_tol=0.2,
        title_size_min_norm=0.04,
        max_image_aspect_drift=0.25,
        min_image_area=0.003,
        max_image_size_spread=0.7,
        title_drift_tol=0.25,
        margin_drift_tol=0.15,
        max_background_like_ratio=0.90,
        disabled_rules=frozenset({RULE_OFF_BALANCE, RULE_CORNER_CLUSTER}),
    ),
}


def get_profile(name: str) -> ArtProfile:
    if name not in _BUILTIN:
        raise InvalidArgumentError(f"unknown art profile: {name}")
    return _BUILTIN[name]


def profile_names() -> list[str]:
    return list(_BUILTIN)
