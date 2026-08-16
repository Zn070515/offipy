"""FEATURES 特征注册表 + encode_features（offipy.art，避免 art→feedback→art 环）。

FEATURES 归 art：encode_features 消费 ArtFinding / slide / deck（都在 art），且被
art/feedback.py 的 append 直接调用——放 art 内是同包引用。学习核心单向 import art，
禁止反向。schema 版本：FEATURES 内容变更 → feature_schema_version() bump。
"""

from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from typing import TYPE_CHECKING, Any

from .features import _KNOWN_SLIDE_ROLES, compute_features
from .profiles import ALL_RULES, RULE_DIMENSIONS, profile_names

if TYPE_CHECKING:
    from collections.abc import Callable

    from .models import ArtFinding, ArtScene, ArtSlide

_SCHEMA_VERSION = "4"

# 已知 slide role 集合：直接共享 features.py 的单一事实来源（_KNOWN_SLIDE_ROLES），
# 不手抄副本——features.py 增删 role 时此处自动跟随，避免未知 role 静默坍缩成错误桶。
# 顺序按排序后的 index 编码；len(roles) 作为「其他」的显式 catch-all 桶。
_SLIDE_ROLE_ORDER = tuple(sorted(_KNOWN_SLIDE_ROLES))
_SLIDE_ROLE_OTHER = float(len(_SLIDE_ROLE_ORDER))


def feature_schema_version() -> str:
    return _SCHEMA_VERSION


@dataclass(frozen=True)
class FeatureSpec:
    id: str
    group: str  # "finding" | "slide" | "deck"
    kind: str  # "float" | "onehot_cat"
    applies_to: str  # rule_id 或 "all"
    missing_default: float = 0.0
    categories: tuple[str, ...] = ()
    extract: Callable[[dict[str, Any]], float | str | None] | None = None


# --- extract 帮助函数 ---


def _finding_attr(ctx: dict[str, Any], name: str) -> Any:
    return getattr(ctx["finding"], name, None)


def _detail_float(ctx: dict[str, Any], rule_id: str, key: str) -> float | None:
    finding: ArtFinding = ctx["finding"]
    if finding.rule_id != rule_id:
        return None
    v = finding.details.get(key)
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _detail_nested(ctx: dict[str, Any], rule_id: str, top: str, sub: str) -> float | None:
    finding: ArtFinding = ctx["finding"]
    if finding.rule_id != rule_id:
        return None
    v = finding.details.get(top)
    if not isinstance(v, dict):
        return None
    s = v.get(sub)
    return float(s) if isinstance(s, (int, float)) and not isinstance(s, bool) else None


def _detail_dict_max(ctx: dict[str, Any], rule_id: str, key: str) -> float | None:
    finding: ArtFinding = ctx["finding"]
    if finding.rule_id != rule_id:
        return None
    v = finding.details.get(key)
    if not isinstance(v, dict):
        return None
    vals = [x for x in v.values() if isinstance(x, (int, float)) and not isinstance(x, bool)]
    return float(max(vals)) if vals else None


def _detail_list_len(ctx: dict[str, Any], rule_id: str, key: str) -> float | None:
    finding: ArtFinding = ctx["finding"]
    if finding.rule_id != rule_id:
        return None
    v = finding.details.get(key)
    return float(len(v)) if isinstance(v, (list, tuple)) else None


def _measurement(
    rule_id: str, key: str, *, nested: str | None = None, agg: str = "raw"
) -> FeatureSpec:
    if agg == "len":
        extract = lambda ctx: _detail_list_len(ctx, rule_id, key)  # noqa: E731
    elif agg == "dict_max":
        extract = lambda ctx: _detail_dict_max(ctx, rule_id, key)  # noqa: E731
    elif nested is not None:
        extract = lambda ctx: _detail_nested(ctx, rule_id, key, nested)  # noqa: E731
    else:
        extract = lambda ctx: _detail_float(ctx, rule_id, key)  # noqa: E731
    return FeatureSpec(
        id=f"measure.{rule_id}.{key}",
        group="finding",
        kind="float",
        applies_to=rule_id,
        missing_default=0.0,
        extract=extract,
    )


_MEASUREMENT_SPECS: list[FeatureSpec] = [
    _measurement("art.typography.many_families", "families", agg="len"),
    _measurement("art.typography.flat_scale", "ratio"),
    _measurement("art.media.distorted_image", "natural_ratio"),
    _measurement("art.media.distorted_image", "physical_ratio"),
    _measurement("art.media.mixed_image_sizes", "spread"),
    _measurement("art.color.low_contrast", "ratio"),
    _measurement("art.color.low_contrast", "foreground_match_ratio"),
    _measurement("art.color.accent_flood", "accent_ratio"),
    _measurement("art.color.no_accent", "accent_ratio"),
    _measurement("art.consistency.title_drift", "x_drift"),
    _measurement("art.consistency.title_drift", "size_drift"),
    _measurement("art.consistency.margin_drift", "median"),
    _measurement("art.composition.off_balance", "balance_dist"),
    _measurement("art.composition.off_balance", "ink"),
    _measurement("art.composition.corner_cluster", "quadrants", agg="dict_max"),
    _measurement("art.composition.spacing_drift", "horizontal", nested="max_drift_ratio"),
    _measurement("art.composition.spacing_drift", "vertical", nested="max_drift_ratio"),
    _measurement("art.composition.background_like_area", "background_like_ratio"),
    _measurement("art.media.tiny_image", "area_ratio"),
    _measurement("art.typography.tiny_text", "font_size_norm"),
    _measurement("art.typography.tiny_text", "ratio_vs_min"),
    _measurement("art.hierarchy.title_too_small", "font_size_norm"),
    _measurement("art.hierarchy.title_too_small", "ratio_vs_min"),
    _measurement("art.hierarchy.no_focus", "focus_ratio"),
]


def _slide_scalar(ctx: dict[str, Any], top: str, sub: str | None = None) -> float | None:
    feats = ctx.get("slide_features", {})
    v = feats.get(top)
    if sub is not None and isinstance(v, dict):
        v = v.get(sub)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    return None


def _slide_spacing_drift(ctx: dict[str, Any]) -> float | None:
    feats = ctx.get("slide_features", {})
    h = feats.get("spacing", {}).get("horizontal", {}).get("max_drift_ratio")
    v = feats.get("spacing", {}).get("vertical", {}).get("max_drift_ratio")
    vals = [x for x in (h, v) if isinstance(x, (int, float)) and not isinstance(x, bool)]
    return float(max(vals)) if vals else None


def _slide_alignment_lines(ctx: dict[str, Any]) -> float | None:
    lines = ctx.get("slide_features", {}).get("alignment", {}).get("lines")
    return float(len(lines)) if isinstance(lines, list) else None


def _slide_page_signature(ctx: dict[str, Any]) -> float | None:
    role = ctx.get("slide_features", {}).get("page_signature", {}).get("role")
    if role in _SLIDE_ROLE_ORDER:
        return float(_SLIDE_ROLE_ORDER.index(role))
    if role is None:
        return None
    return _SLIDE_ROLE_OTHER


_RULE_ORDER = tuple(sorted(ALL_RULES))
_DIM_ORDER = tuple(sorted(set(RULE_DIMENSIONS.values())))
_PROFILE_ORDER = tuple(sorted(profile_names()))


def _onehot_spec(
    spec_id: str, categories: tuple[str, ...], extract: Callable[[dict[str, Any]], str | None]
) -> FeatureSpec:
    return FeatureSpec(
        id=spec_id,
        group="finding" if spec_id.startswith("finding.") else "deck",
        kind="onehot_cat",
        applies_to="all",
        missing_default=0.0,
        categories=categories,
        extract=extract,
    )


def _to_float_optional(v: Any) -> float | None:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    return None


def _page_ratio(ctx: dict[str, Any]) -> float | None:
    total = ctx.get("total_slides", 0)
    idx = ctx["finding"].slide_index
    if not total or idx is None:
        return None
    return float(idx) / float(total)


FEATURES: dict[str, FeatureSpec] = {
    s.id: s
    for s in [
        _onehot_spec("finding.rule_id", _RULE_ORDER, lambda ctx: str(ctx["finding"].rule_id)),
        _onehot_spec("finding.dimension", _DIM_ORDER, lambda ctx: str(ctx["finding"].dimension)),
        FeatureSpec(
            id="finding.severity_ordinal",
            group="finding",
            kind="float",
            applies_to="all",
            missing_default=2.0,
            extract=lambda ctx: _to_float_optional(_finding_attr(ctx, "severity")),
        ),
        FeatureSpec(
            id="finding.confidence",
            group="finding",
            kind="float",
            applies_to="all",
            missing_default=0.5,
            extract=lambda ctx: _to_float_optional(_finding_attr(ctx, "confidence")),
        ),
        FeatureSpec(
            id="finding.evidence_reliability",
            group="finding",
            kind="float",
            applies_to="all",
            missing_default=0.0,
            extract=lambda ctx: _to_float_optional(_finding_attr(ctx, "evidence_reliability")),
        ),
        FeatureSpec(
            id="finding.page_ratio",
            group="finding",
            kind="float",
            applies_to="all",
            missing_default=0.0,
            extract=_page_ratio,
        ),
        *_MEASUREMENT_SPECS,
        FeatureSpec(
            id="slide.alignment",
            group="slide",
            kind="float",
            applies_to="all",
            missing_default=0.0,
            extract=_slide_alignment_lines,
        ),
        FeatureSpec(
            id="slide.spacing",
            group="slide",
            kind="float",
            applies_to="all",
            missing_default=0.0,
            extract=_slide_spacing_drift,
        ),
        FeatureSpec(
            id="slide.mass",
            group="slide",
            kind="float",
            applies_to="all",
            missing_default=0.0,
            extract=lambda ctx: _slide_scalar(ctx, "mass", "ink"),
        ),
        FeatureSpec(
            id="slide.density",
            group="slide",
            kind="float",
            applies_to="all",
            missing_default=0.0,
            extract=lambda ctx: _slide_scalar(ctx, "density", "sum_area_ratio"),
        ),
        FeatureSpec(
            id="slide.font_hierarchy",
            group="slide",
            kind="float",
            applies_to="all",
            missing_default=1.0,
            extract=lambda ctx: _slide_scalar(ctx, "font_hierarchy", "ratio"),
        ),
        FeatureSpec(
            id="slide.palette",
            group="slide",
            kind="float",
            applies_to="all",
            missing_default=0.0,
            extract=lambda ctx: _slide_scalar(ctx, "palette", "accent_ratio"),
        ),
        FeatureSpec(
            id="slide.focus",
            group="slide",
            kind="float",
            applies_to="all",
            missing_default=0.0,
            extract=lambda ctx: _slide_scalar(ctx, "focus", "ratio"),
        ),
        FeatureSpec(
            id="slide.page_signature",
            group="slide",
            kind="float",
            applies_to="all",
            # 注意：missing_default=_SLIDE_ROLE_OTHER 会 conflate 两种状态——
            # 「无 slide 数据」（slide=None）与「真实 slide 上的未知 role」都落到
            # catch-all 桶 6.0。这是刻意的（比旧值 2.0 撞 cover 好），但 6.0 不恒等
            # 于「真实未知 role」，Task 6/7 消费时勿按后者解读。已知 role 取 sorted 序号 0..5。
            missing_default=_SLIDE_ROLE_OTHER,
            extract=_slide_page_signature,
        ),
        FeatureSpec(
            id="deck.total_slides",
            group="deck",
            kind="float",
            applies_to="all",
            missing_default=0.0,
            extract=lambda ctx: (
                float(ctx.get("total_slides", 0)) if ctx.get("total_slides") else None
            ),
        ),
        _onehot_spec(
            "deck.profile", _PROFILE_ORDER, lambda ctx: str(ctx.get("profile", "balanced"))
        ),
    ]
}


@lru_cache(maxsize=1)
def feature_keys() -> tuple[str, ...]:
    """确定性输入维度布局：全部 canonical feature key。"""
    keys: list[str] = []
    for spec in sorted(FEATURES.values(), key=lambda s: s.id):
        if spec.kind == "onehot_cat":
            keys.extend(f"{spec.id}.{c}" for c in spec.categories)
        else:
            keys.append(spec.id)
    return tuple(keys)


_KEY_TO_ENTRY: dict[str, tuple[FeatureSpec, str | None]] = {}
for _spec in FEATURES.values():
    if _spec.kind == "onehot_cat":
        for _c in _spec.categories:
            _KEY_TO_ENTRY[f"{_spec.id}.{_c}"] = (_spec, _c)
    else:
        _KEY_TO_ENTRY[_spec.id] = (_spec, None)


def _emit(full_key: str, ctx: dict[str, Any]) -> float:
    spec, category = _KEY_TO_ENTRY[full_key]
    if spec.kind == "onehot_cat":
        assert category is not None
        value = spec.extract(ctx) if spec.extract else None
        return 1.0 if value == category else 0.0
    value = spec.extract(ctx) if spec.extract else None
    return spec.missing_default if value is None else float(value)


def encode_features(
    finding: ArtFinding,
    slide: ArtSlide | None = None,
    deck: ArtScene | None = None,
    profile: str = "balanced",
) -> dict[str, float]:
    """ArtFinding/slide/deck → 扁平标量快照（全部 feature key，缺失回退 missing_default）。"""
    slide_feats = compute_features(slide) if slide is not None else {}
    ctx: dict[str, Any] = {
        "finding": finding,
        "slide": slide,
        "slide_features": slide_feats,
        "total_slides": len(deck.slides) if deck is not None else 0,
        "profile": profile,
    }
    return {key: _emit(key, ctx) for key in feature_keys()}
