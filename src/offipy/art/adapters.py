"""场景适配器：测量数据（Measurement）与几何审计（PptxAudit）→ ArtScene。

rev2.1：两个适配器严格适配 0.11.6 真实 schema；测试/手写场景用 ArtScene.from_dict，
两者输入协议分离。slides_dir 由 PixelEnricher 增强。
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

from offipy.exceptions import InvalidArgumentError

from .merge import merge_scenes
from .models import ArtColor, ArtElement, ArtScene, ArtSlide, ArtTextRun, ArtWarning

if TYPE_CHECKING:
    from offipy.audit.models import PptxAuditReport

# per-slide 元素上限：恶意/病态输入页元素数超限即截断，约束下游 O(n²)
# （features._union_area、merge_scenes）的复杂度上限。
_MAX_SLIDE_ELEMENTS = 3000

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


def _measurement_color(raw: str | dict[str, Any] | None) -> ArtColor | None:
    if isinstance(raw, dict):
        return ArtColor.from_dict(raw)
    return _rgb_string_to_color(raw)


def _num(value: Any, default: float) -> float:
    """测量数值字段：缺失/损坏回退默认，不抛 ValueError。"""
    if value is None:
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def _opt_num(value: Any, default: float | None = None) -> float | None:
    """测量可选数值字段：缺失 → None，损坏 → default（默认 None）。"""
    if value is None:
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def _to_measurement_element(
    rec: dict[str, Any], slide_index: int, slide_width: float, slide_height: float
) -> ArtElement:
    """真实 measurements record → ArtElement。rect 为像素，按页宽高归一化。"""
    rect = rec.get("rect") or {}
    x = _num(rect.get("x"), 0.0) / slide_width
    y = _num(rect.get("y"), 0.0) / slide_height
    w = _num(rect.get("w"), 0.0) / slide_width
    h = _num(rect.get("h"), 0.0) / slide_height
    raw_kind = rec.get("kind", "shape")
    kind = _KIND_MAP.get(raw_kind, "shape")
    style = rec.get("style") or {}
    deco = rec.get("deco") or {}
    natural = rec.get("naturalSize") or {}
    runs = [
        ArtTextRun(
            text=rt.get("text", ""),
            font_size=_opt_num(rt.get("fontSize")),
            font_size_unit="px",
            font_family=rt.get("fontFamily"),
            color=_measurement_color(rt.get("color")),
        )
        for rt in rec.get("runs", [])
    ]
    fs = style.get("fontSize")
    if isinstance(fs, str) and fs.endswith("px"):
        font_size = _opt_num(fs[:-2])
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
        natural_width=_opt_num(natural_w),
        natural_height=_opt_num(natural_h),
        source="measurement",
    )


class MeasurementAdapter:
    """真实 0.11.6 measurements.json → ArtScene。输入 dict 或 JSON 字符串/路径。"""

    def __init__(self, data: dict[str, Any] | str) -> None:
        self._data: dict[str, Any] = data if isinstance(data, dict) else json.loads(data)

    def build(self) -> ArtScene:
        slides: list[ArtSlide] = []
        warnings: list[ArtWarning] = []
        for i, item in enumerate(self._data.get("slides", [])):
            s = item.get("slide") or item  # 兼容 {slide:{...}} 与裸 dict
            index = i + 1  # 位置索引 → 1-based 公开索引
            width = _num(s.get("width"), 1920.0)
            height = _num(s.get("height"), 1080.0)
            records = item.get("records", [])
            if len(records) > _MAX_SLIDE_ELEMENTS:
                warnings.append(
                    ArtWarning(
                        code="art.adapter.elements_truncated",
                        message=(
                            f"页 {index} 元素数 {len(records)} 超过上限 {_MAX_SLIDE_ELEMENTS}，截断"
                        ),
                    )
                )
                records = records[:_MAX_SLIDE_ELEMENTS]
            elements = [_to_measurement_element(rec, index, width, height) for rec in records]
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

    def __init__(self, report: PptxAuditReport) -> None:
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
        truncated: set[int] = set()
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
            if len(slides[index].elements) >= _MAX_SLIDE_ELEMENTS:
                if index not in truncated:
                    truncated.add(index)
                    warnings.append(
                        ArtWarning(
                            code="art.adapter.elements_truncated",
                            message=f"页 {index} 元素数超过上限 {_MAX_SLIDE_ELEMENTS}，截断",
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
                # 上面的守卫已排除 None（geometry_unknown / None in (...)），这里显式收窄
                x=(cast("float", snap.left) * 72.0) / width,
                y=(cast("float", snap.top) * 72.0) / height,
                width=(cast("float", snap.width) * 72.0) / width,
                height=(cast("float", snap.height) * 72.0) / height,
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
    measurements: str | dict[str, Any] | None = None,
    pptx: str | None = None,
    pptx_report: object | None = None,
    slides_dir: str | None = None,
) -> ArtScene:
    """建场景：测量数据 / 几何审计 / 像素 slides_dir 三源融合（v0.12.1）。"""
    if slides_dir is not None:
        from .pixels import PixelEnricher, empty_scene_from_slides

        d = Path(slides_dir)
        if not d.is_dir():
            raise InvalidArgumentError(f"slides_dir 不存在或不是目录: {slides_dir}")
    if pptx is not None and pptx_report is not None:
        raise InvalidArgumentError("build_scene: pass either pptx (path) or pptx_report, not both")
    scenes: list[ArtScene] = []
    if measurements is not None:
        if isinstance(measurements, (str, Path)):
            p = Path(measurements)
            try:
                is_file = p.is_file()
            except OSError:
                # 超长 JSON 字符串在 Linux 触发 ENAMETOOLONG，Windows 返回 False → 视为 JSON 字符串
                is_file = False
            raw: dict[str, Any] | str = (
                p.read_text(encoding="utf-8") if is_file else str(measurements)
            )
        else:
            raw = measurements
        scenes.append(MeasurementAdapter(raw).build())
    if pptx_report is not None:
        scenes.append(PptxAuditAdapter(cast("PptxAuditReport", pptx_report)).build())
    elif pptx is not None:
        from offipy.audit import audit_pptx  # 惰性：避免顶层依赖 python-pptx

        scenes.append(PptxAuditAdapter(audit_pptx(pptx)).build())
    if not scenes:
        if slides_dir is not None:
            scene = empty_scene_from_slides(d)
        else:
            raise InvalidArgumentError("build_scene requires measurements, pptx, or slides_dir")
    elif len(scenes) == 1:
        scene = scenes[0]
    else:
        scene, _warnings = merge_scenes(primary=scenes[0], secondary=scenes[1])
    if slides_dir is not None:
        expected_sha = _expected_sha256(pptx, pptx_report)
        run_id = _measurements_run_id(measurements)
        scene = PixelEnricher(d).enrich(scene, expected_sha256=expected_sha, run_id=run_id)
    return scene


def _sha256_file(path: str | Path) -> str:
    import hashlib

    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _expected_sha256(pptx: str | None, pptx_report: object | None) -> str | None:
    if pptx is not None:
        return _sha256_file(pptx)
    return getattr(pptx_report, "source_sha256", None)


def _measurements_run_id(measurements: str | dict[str, Any] | None) -> str | None:
    if isinstance(measurements, dict):
        return measurements.get("run_id")
    if isinstance(measurements, (str, Path)):
        p = Path(measurements)
        try:
            is_file = p.is_file()
        except OSError:
            is_file = False
        if is_file:
            try:
                return cast("str | None", json.loads(p.read_text(encoding="utf-8")).get("run_id"))
            except (OSError, ValueError):
                return None
        # 内联 JSON 字符串也尝试提取 run_id（路径字符串解析失败安全返回 None）
        try:
            return cast("str | None", json.loads(str(measurements)).get("run_id"))
        except (OSError, ValueError):
            return None
    return None
