"""场景适配器：测量数据（Measurement）与几何审计（PptxAudit）→ ArtScene。

rev2.1：两个适配器严格适配 0.11.6 真实 schema；测试/手写场景用 ArtScene.from_dict，
两者输入协议分离。build_scene(slides_dir=...) 一律拒绝。
"""

from __future__ import annotations

import json
from pathlib import Path

from offipy.exceptions import InvalidArgumentError

from .merge import merge_scenes
from .models import ArtColor, ArtElement, ArtScene, ArtSlide, ArtTextRun, ArtWarning

_KIND_MAP = {
    "text": "text",
    "txt": "text",
    "canvas": "image",  # 真实 schema：带 naturalSize 的 div/canvas → 图片
    "img": "image",
    "image": "image",
    "picture": "image",
    "shape": "shape",
    "rect": "shape",
    "line": "shape",
    "group": "container",
    "container": "container",
}

_ROLE_KEYWORDS = {
    "title": ["title", "heading", "slide-title", "标题", "大标题"],
    "subtitle": ["subtitle", "sub-heading", "副标题"],
    "caption": ["caption", "note", "注释", "说明"],
    "page_number": ["page-number", "pagenum", "页码"],
    "footer": ["footer", "页脚"],
    "background": ["background", "bg", "背景"],
    "decoration": ["deco", "decorative", "ornament", "装饰"],
    "image": ["img", "image", "figure", "插图"],
}


def _infer_element_role(className: str | None, tag: str | None, kind: str) -> str:
    text = " ".join(x for x in (className or "", tag or "") if x).lower()
    for role, kws in _ROLE_KEYWORDS.items():
        if any(k in text for k in kws):
            return role
    if kind == "image":
        return "image"
    if kind == "text":
        return "body"
    return "shape"


def _rgb_string_to_color(s: str | None) -> ArtColor | None:
    """解析 'rgb(r,g,b)' / 'rgba(r,g,b,a)'。透明（a=0）→ None。"""
    if not s:
        return None
    s = s.strip()
    if not (s.startswith("rgb") and "(" in s and s.endswith(")")):
        return None
    inner = s[s.index("(") + 1 : s.rindex(")")]
    parts = [p.strip() for p in inner.split(",")]
    try:
        r, g, b = (int(float(p)) for p in parts[:3])
        a = float(parts[3]) if len(parts) > 3 else 1.0
    except (ValueError, IndexError):
        return None
    if a <= 0.0:
        return None  # 透明背景 → 无背景证据
    return ArtColor(r, g, b, a)


def _measurement_color(raw: str | dict | None) -> ArtColor | None:
    if isinstance(raw, dict):
        return ArtColor.from_dict(raw)
    return _rgb_string_to_color(raw)


def _to_measurement_element(
    rec: dict, slide_index: int, slide_width: float, slide_height: float
) -> ArtElement:
    """真实 measurements record → ArtElement。rect 为像素，按页宽高归一化。"""
    rect = rec.get("rect") or {}
    x = float(rect.get("x", 0.0)) / slide_width
    y = float(rect.get("y", 0.0)) / slide_height
    w = float(rect.get("w", 0.0)) / slide_width
    h = float(rect.get("h", 0.0)) / slide_height
    raw_kind = rec.get("kind", "shape")
    kind = _KIND_MAP.get(raw_kind, "shape")
    style = rec.get("style") or {}
    deco = rec.get("deco") or {}
    natural = rec.get("naturalSize") or {}
    runs = [
        ArtTextRun(
            text=rt.get("text", ""),
            font_size=float(rt["fontSize"]) if rt.get("fontSize") else None,
            font_size_unit="px",
            font_family=rt.get("fontFamily"),
            color=_measurement_color(rt.get("color")),
        )
        for rt in rec.get("runs", [])
    ]
    fs = style.get("fontSize")
    if isinstance(fs, str) and fs.endswith("px"):
        font_size = float(fs[:-2])
        font_size_unit = "px"
    elif isinstance(fs, (int, float)):
        font_size = float(fs)
        font_size_unit = "px"
    else:
        font_size = None
        font_size_unit = "unknown"
    has_bg = bool(deco.get("hasBg", False))
    background = _measurement_color(deco.get("bg")) if has_bg else None
    natural_w = natural.get("w")
    natural_h = natural.get("h")
    return ArtElement(
        element_id=f"m{slide_index}-{rec.get('id')}",
        kind=kind,
        role=_infer_element_role(rec.get("className"), rec.get("tag"), kind),
        x=x,
        y=y,
        width=w,
        height=h,
        slide_index=slide_index,
        foreground=_measurement_color(style.get("color")),
        background=background,
        text=rec.get("text", ""),
        font_size=font_size,
        font_size_unit=font_size_unit,
        font_size_norm=(font_size / slide_height) if (font_size and slide_height) else None,
        runs=runs,
        natural_width=float(natural_w) if natural_w else None,
        natural_height=float(natural_h) if natural_h else None,
        source="measurement",
    )


class MeasurementAdapter:
    """真实 0.11.6 measurements.json → ArtScene。输入 dict 或 JSON 字符串/路径。"""

    def __init__(self, data: dict | str) -> None:
        self._data = data if isinstance(data, dict) else json.loads(data)

    def build(self) -> ArtScene:
        slides: list[ArtSlide] = []
        warnings: list[ArtWarning] = []
        for i, item in enumerate(self._data.get("slides", [])):
            s = item.get("slide") or item  # 兼容 {slide:{...}} 与裸 dict
            index = i + 1  # 位置索引 → 1-based 公开索引
            width = float(s.get("width", 1920.0))
            height = float(s.get("height", 1080.0))
            elements = [
                _to_measurement_element(rec, index, width, height)
                for rec in item.get("records", [])
            ]
            bg = _measurement_color(s.get("background"))
            slides.append(
                ArtSlide(
                    index=index, width=width, height=height, elements=elements, background_color=bg
                )
            )
        return ArtScene(slides=slides, width_unit="px", warnings=warnings, sources={"measurement"})


class PptxAuditAdapter:
    """0.11 几何审计报告（PptxAuditReport）→ ArtScene。单位 pt（英寸×72）。

    rev2.1：只读 report.slide_size / slide_count / shapes（SlideShapeSnapshot）。
    SlideShapeSnapshot.slide_index 已 1-based，不做 +1。无字号/颜色证据。
    """

    def __init__(self, report) -> None:
        self._report = report

    def build(self) -> ArtScene:
        w_in, h_in = self._report.slide_size
        width = w_in * 72.0
        height = h_in * 72.0
        # 预建全部 1..slide_count 页：空白页也保留，索引保持连续
        slides: dict[int, ArtSlide] = {
            i: ArtSlide(index=i, width=width, height=height)
            for i in range(1, self._report.slide_count + 1)
        }
        warnings: list[ArtWarning] = []
        for snap in self._report.shapes:
            index = snap.slide_index  # 已 1-based，不做 +1
            if index < 1 or index > self._report.slide_count:
                warnings.append(
                    ArtWarning(
                        code="art.adapter.index_out_of_range",
                        message=f"shape slide_index {index} 超出 slide_count，跳过",
                    )
                )
                continue
            if snap.geometry_unknown or None in (snap.left, snap.top, snap.width, snap.height):
                warnings.append(
                    ArtWarning(
                        code="art.adapter.geometry_unknown",
                        message=f"shape {snap.shape_id} 无几何信息，跳过",
                    )
                )
                continue
            kind = _KIND_MAP.get(snap.shape_type.lower(), "shape")
            if snap.shape_type.lower() in ("picture", "photo"):
                kind = "image"
            el = ArtElement(
                element_id=f"pptx-{index}-{snap.shape_id}",
                kind=kind,
                role=snap.role or "shape",
                x=(snap.left * 72.0) / width,
                y=(snap.top * 72.0) / height,
                width=(snap.width * 72.0) / width,
                height=(snap.height * 72.0) / height,
                slide_index=index,
                text=snap.text or "",
                source="pptx",
            )
            slides[index].elements.append(el)
        return ArtScene(
            slides=[slides[k] for k in sorted(slides)],
            width_unit="pt",
            warnings=warnings,
            sources={"pptx"},
        )


def build_scene(
    *,
    measurements: str | dict | None = None,
    pptx: str | None = None,
    pptx_report: object | None = None,
    slides_dir: str | None = None,
) -> ArtScene:
    """建场景：测量数据与/或几何审计（可复用已审计 report）。slides_dir 不支持。"""
    if slides_dir:
        raise InvalidArgumentError("Rendered slide analysis is not implemented in v0.12")
    if pptx is not None and pptx_report is not None:
        raise InvalidArgumentError("build_scene: pass either pptx (path) or pptx_report, not both")
    scenes: list[ArtScene] = []
    if measurements is not None:
        if isinstance(measurements, (str, Path)):
            p = Path(measurements)
            raw: dict | str = p.read_text(encoding="utf-8") if p.is_file() else str(measurements)
        else:
            raw = measurements
        scenes.append(MeasurementAdapter(raw).build())
    if pptx_report is not None:
        scenes.append(PptxAuditAdapter(pptx_report).build())
    elif pptx is not None:
        from offipy.audit import audit_pptx  # 惰性：避免顶层依赖 python-pptx

        scenes.append(PptxAuditAdapter(audit_pptx(pptx)).build())
    if not scenes:
        raise InvalidArgumentError("build_scene requires measurements or pptx")
    if len(scenes) == 1:
        return scenes[0]
    merged, _warnings = merge_scenes(primary=scenes[0], secondary=scenes[1])
    return merged
