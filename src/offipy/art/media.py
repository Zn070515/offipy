"""图片媒体规则：失真 / 过小 / 尺寸混杂。"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from offipy.audit import Severity

from .features import physical_aspect_ratio
from .profiles import (
    RULE_DISTORTED_IMAGE,
    RULE_MIXED_IMAGE_SIZES,
    RULE_TINY_IMAGE,
)
from .rules import RuleContext, RuleEvaluation, RuleSpec, make_finding

if TYPE_CHECKING:
    from .models import ArtElement, ArtSlide


def _images(slide: ArtSlide) -> list[ArtElement]:
    return [e for e in slide.elements if e.kind == "image"]


def distorted_image_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    imgs = _images(slide)
    eligible = imgs  # 所有图片都在评估范围
    covered = [
        e
        for e in imgs
        if e.natural_width is not None and e.natural_height is not None and e.natural_height > 0
    ]
    out = []
    for e in covered:
        # covered 过滤保证 natural_width/natural_height 非 None
        natural = cast("float", e.natural_width) / cast("float", e.natural_height)
        physical = physical_aspect_ratio(e, slide.width, slide.height)
        if natural == 0:
            continue
        drift = abs(physical - natural) / natural
        if drift > ctx.profile.max_image_aspect_drift:
            out.append(
                make_finding(
                    RULE_DISTORTED_IMAGE,
                    "media",
                    Severity.MID,
                    f"图片被拉伸失真（宽高比漂移 {drift:.2f}）。",
                    0.6,
                    slide.index,
                    primary=e,
                    details={
                        "natural_ratio": round(natural, 3),
                        "physical_ratio": round(physical, 3),
                    },
                )
            )
    return RuleEvaluation(findings=out, covered_count=len(covered), eligible_count=len(eligible))


def tiny_image_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    imgs = _images(slide)
    out = [
        make_finding(
            RULE_TINY_IMAGE,
            "media",
            Severity.LOW,
            f"图片过小（面积占比 {e.area:.4f}）。",
            0.5,
            slide.index,
            primary=e,
            details={"area_ratio": round(e.area, 3)},
        )
        for e in imgs
        if e.area < ctx.profile.min_image_area
    ]
    return RuleEvaluation(findings=out, covered_count=len(imgs), eligible_count=len(imgs))


def mixed_image_sizes_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    imgs = _images(slide)
    if len(imgs) < 3:
        return RuleEvaluation(covered_count=len(imgs), eligible_count=len(imgs))
    areas = sorted(e.area for e in imgs)
    if areas[0] == 0:
        return RuleEvaluation(covered_count=len(imgs), eligible_count=len(imgs))
    spread = (areas[-1] - areas[0]) / areas[-1]
    if spread <= ctx.profile.max_image_size_spread:
        return RuleEvaluation(covered_count=len(imgs), eligible_count=len(imgs))
    largest = max(imgs, key=lambda e: e.area)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_MIXED_IMAGE_SIZES,
                "media",
                Severity.LOW,
                f"图片尺寸混杂（面积跨度 {spread:.2f}）。",
                0.5,
                slide.index,
                primary=largest,
                details={"spread": round(spread, 3)},
            )
        ],
        covered_count=len(imgs),
        eligible_count=len(imgs),
    )


RULES = [
    RuleSpec(rule_id=RULE_DISTORTED_IMAGE, dimension="media", run=distorted_image_rule),
    RuleSpec(rule_id=RULE_TINY_IMAGE, dimension="media", run=tiny_image_rule),
    RuleSpec(rule_id=RULE_MIXED_IMAGE_SIZES, dimension="media", run=mixed_image_sizes_rule),
]
