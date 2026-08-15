"""art 测试构造器：默认像素单位、1-based slide index、px 字号。

rev2.1：helper 用 ArtElement 的前景色/背景色/边框色三分离字段；
font_size_norm 由 make_slide 按页高填入（模型 __post_init__ 不再回填）。
"""

from __future__ import annotations

from offipy.art.models import (
    ArtColor,
    ArtElement,
    ArtScene,
    ArtSlide,
    ArtTextRun,
    ElementPixelEvidence,
)


def _prepare_element(el: ArtElement, *, slide_index: int, slide_height: float) -> ArtElement:
    """重建元素：统一 slide_index，并给缺 norm 的 px/pt 元素填 norm = font_size/height。"""
    norm = el.font_size_norm
    if norm is None and el.font_size is not None and el.font_size_unit != "unknown":
        norm = el.font_size / slide_height
    return ArtElement(
        element_id=el.element_id,
        kind=el.kind,
        role=el.role,
        x=el.x,
        y=el.y,
        width=el.width,
        height=el.height,
        slide_index=slide_index,
        foreground=el.foreground,
        background=el.background,
        border=el.border,
        is_background=el.is_background,
        text=el.text,
        font_size=el.font_size,
        font_size_unit=el.font_size_unit,
        font_size_norm=norm,
        runs=el.runs,
        natural_width=el.natural_width,
        natural_height=el.natural_height,
        source=el.source,
        evidence=el.evidence,
        container=el.container,
        decoration=el.decoration,
        pixel_evidence=el.pixel_evidence,
        opacity=el.opacity,
        decoded_width=el.decoded_width,
        decoded_height=el.decoded_height,
        fill_kind=el.fill_kind,
    )


def make_element(
    element_id: str,
    kind: str = "shape",
    role: str = "body",
    x: float = 0.1,
    y: float = 0.1,
    w: float = 0.2,
    h: float = 0.2,
    slide_index: int = 1,
    foreground: ArtColor | None = None,
    background: ArtColor | None = None,
    border: ArtColor | None = None,
    is_background: bool = False,
    text: str = "",
    font_size: float | None = None,
    font_size_unit: str = "px",
    runs: list[ArtTextRun] | None = None,
    natural_width: float | None = None,
    natural_height: float | None = None,
    container: bool = False,
    decoration: bool = False,
    pixel_evidence: ElementPixelEvidence | None = None,
    opacity: float | None = None,
    decoded_width: float | None = None,
    decoded_height: float | None = None,
    fill_kind: str | None = None,
) -> ArtElement:
    return ArtElement(
        element_id=element_id,
        kind=kind,
        role=role,
        x=x,
        y=y,
        width=w,
        height=h,
        slide_index=slide_index,
        foreground=foreground,
        background=background,
        border=border,
        is_background=is_background,
        text=text,
        font_size=font_size,
        font_size_unit=font_size_unit,
        runs=runs or [],
        natural_width=natural_width,
        natural_height=natural_height,
        container=container,
        decoration=decoration,
        pixel_evidence=pixel_evidence,
        opacity=opacity,
        decoded_width=decoded_width,
        decoded_height=decoded_height,
        fill_kind=fill_kind,
    )


def make_text_element(
    element_id: str,
    text: str,
    x: float = 0.1,
    y: float = 0.1,
    w: float = 0.2,
    h: float = 0.08,
    font_size: float = 24.0,
    font_size_unit: str = "px",
    role: str = "body",
    slide_index: int = 1,
    foreground: ArtColor | None = None,
    background: ArtColor | None = None,
    border: ArtColor | None = None,
    is_background: bool = False,
    runs: list[ArtTextRun] | None = None,
) -> ArtElement:
    if runs is None:
        runs = [
            ArtTextRun(
                text=text, font_size=font_size, font_size_unit=font_size_unit, color=foreground
            )
        ]
    return make_element(
        element_id,
        kind="text",
        role=role,
        x=x,
        y=y,
        w=w,
        h=h,
        slide_index=slide_index,
        foreground=foreground,
        text=text,
        background=background,
        border=border,
        is_background=is_background,
        font_size=font_size,
        font_size_unit=font_size_unit,
        runs=runs,
    )


def make_image_element(
    element_id: str,
    x: float = 0.1,
    y: float = 0.1,
    w: float = 0.4,
    h: float = 0.3,
    natural_width: float = 800.0,
    natural_height: float = 600.0,
    decoded_width: float | None = None,
    decoded_height: float | None = None,
    slide_index: int = 1,
) -> ArtElement:
    return make_element(
        element_id,
        kind="image",
        role="image",
        x=x,
        y=y,
        w=w,
        h=h,
        slide_index=slide_index,
        natural_width=natural_width,
        natural_height=natural_height,
        decoded_width=decoded_width,
        decoded_height=decoded_height,
    )


def make_slide(
    index: int = 1,
    width: float = 1920.0,
    height: float = 1080.0,
    elements: list[ArtElement] | None = None,
    background_color: ArtColor | None = None,
) -> ArtSlide:
    els = [_prepare_element(e, slide_index=index, slide_height=height) for e in (elements or [])]
    return ArtSlide(
        index=index,
        width=width,
        height=height,
        elements=els,
        background_color=background_color,
    )


def make_scene(slides: list[ArtSlide], width_unit: str = "px") -> ArtScene:
    return ArtScene(slides=slides, width_unit=width_unit)
