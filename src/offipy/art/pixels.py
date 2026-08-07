"""像素证据增强：RenderedSlide（slides_dir PNG）→ ArtScene 第三证据源。

v0.12.1：纯 stdlib + 惰性 Pillow（缺依赖抛可操作 InvalidArgumentError）。
组合 PNG 无法把像素归因到元素——绝不判断溢出/真实遮挡；只做页面级
背景/留白证据 + 元素级声明颜色验证。

冻结契约（rev2.1）：
- 背景未知时透明像素用 sentinel 掩码排除；ratio 相对有效像素。
- 颜色用 /32 RGB 分桶（最多 512 种 ≤ 2^16，getcolors 永不返回 None）。
- Shape 只验证声明 fill，绝不盲猜中心色。
"""

from __future__ import annotations

import json
import re
from collections import Counter
from dataclasses import dataclass, replace
from pathlib import Path

from offipy.exceptions import InvalidArgumentError

from .models import (
    ArtColor,
    ArtElement,
    ArtScene,
    ArtSlide,
    ArtWarning,
    ElementPixelEvidence,
    PixelColorShare,
    SlidePixelEvidence,
)

_SLIDE_RE = re.compile(r"^slide_(\d+)\.png$")

_MAX_COLORS = 1 << 16
_PALETTE_BUCKET = 32
_PALETTE_TOP_N = 8
_COLOR_TOL_BUCKET = 1
_MAX_DIM_SAMPLE = 256
_FG_MATCH_MIN = 0.30
_FILL_MATCH_MIN = 0.60
_INSET = 0.15
_COMPLEX_BUCKET_MIN_SHARE = 0.05
_COMPLEX_THRESHOLD = 0.5
_PAGE_ASPECT_TOL = 0.01
_BG_SAMPLE_MARGIN = 0.04
_OPAQUE_SENTINEL = (252, 0, 252)
_ALPHA_MIN = 128
# 注意：/32 分桶后过滤的是整个桶 (7,0,7)（R∈[224,255], G∈[0,31], B∈[224,255]），
# 所以真实的热洋红元素（#FF00FF 及其邻近色）也会被一起排除在统计外。
# 这是冻结契约的取舍，仅记录，不改变行为。
_SENTINEL_BUCKET = (
    _OPAQUE_SENTINEL[0] // _PALETTE_BUCKET,
    _OPAQUE_SENTINEL[1] // _PALETTE_BUCKET,
    _OPAQUE_SENTINEL[2] // _PALETTE_BUCKET,
)


def _pil():
    """惰性 import Pillow：pixels.py 顶层不依赖 PIL，纯 `import offipy` 不加载。"""
    from PIL import Image, ImageChops

    return Image, ImageChops


def _bucket_rep(key: tuple[int, int, int]) -> ArtColor:
    return ArtColor(
        min(key[0] * _PALETTE_BUCKET + 16, 255),
        min(key[1] * _PALETTE_BUCKET + 16, 255),
        min(key[2] * _PALETTE_BUCKET + 16, 255),
    )


def _bucket_dist(a: tuple[int, int, int], b: tuple[int, int, int]) -> int:
    return max(abs(a[0] - b[0]), abs(a[1] - b[1]), abs(a[2] - b[2]))


def _composite(im, bg: ArtColor):
    """有 alpha 的图按已知背景合成；无 alpha 直接返回。"""
    Image, _ = _pil()
    if im.mode == "RGBA":
        bg_img = Image.new("RGBA", im.size, (bg.r, bg.g, bg.b, 255))
        return Image.alpha_composite(bg_img, im)
    return im


def _mask_opaque(im):
    """背景未知：把透明像素替换成 sentinel（后续统计时排除该桶）。"""
    Image, _ = _pil()
    if im.mode != "RGBA":
        return im
    alpha = im.getchannel("A")
    rgb = im.convert("RGB")
    transparent = alpha.point(lambda a: 255 if a < _ALPHA_MIN else 0)
    return Image.composite(Image.new("RGB", rgb.size, _OPAQUE_SENTINEL), rgb, transparent)


def _pixel_stats(im, bg: ArtColor | None) -> dict[tuple[int, int, int], int]:
    """区域颜色桶直方图：合成或掩码 → /32 分桶 → getcolors。

    Pillow 12 的 point() 只接受平坦 LUT（每通道 256 值，RGB 共 768）；
    合成到已知不透明背景后 RGBA→RGB（alpha 恒 255，转换无损）。
    """
    im = _composite(im, bg) if bg is not None else _mask_opaque(im)
    if im.mode != "RGB":
        im = im.convert("RGB")
    lut = [v // _PALETTE_BUCKET for v in range(256)]
    bucketed = im.point(lut * 3)
    counts: dict[tuple[int, int, int], int] = {}
    for count, color in bucketed.getcolors(maxcolors=_MAX_COLORS) or []:
        if color == _SENTINEL_BUCKET:
            continue
        counts[color] = counts.get(color, 0) + count
    return counts


def _palette(counts: dict) -> list[PixelColorShare]:
    total = sum(counts.values())
    if not total:
        return []
    top = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:_PALETTE_TOP_N]
    return [PixelColorShare(color=_bucket_rep(k), ratio=round(c / total, 4)) for k, c in top]


def _match_ratio(counts: dict, c: ArtColor) -> float:
    bucket = (c.r // _PALETTE_BUCKET, c.g // _PALETTE_BUCKET, c.b // _PALETTE_BUCKET)
    total = sum(counts.values())
    if not total:
        return 0.0
    matched = sum(
        cnt for key, cnt in counts.items() if _bucket_dist(key, bucket) <= _COLOR_TOL_BUCKET
    )
    return matched / total


def _complexity(counts: dict) -> float:
    total = sum(counts.values())
    if not total:
        return 0.0
    diverse = sum(1 for c in counts.values() if c / total >= _COMPLEX_BUCKET_MIN_SHARE)
    return min(1.0, diverse / 8.0)


def _downsample(im):
    w, h = im.size
    longest = max(w, h)
    if longest <= _MAX_DIM_SAMPLE:
        return im
    scale = _MAX_DIM_SAMPLE / longest
    return im.resize(
        (max(1, round(w * scale)), max(1, round(h * scale))),
        _pil()[0].Resampling.BILINEAR,
    )


def _background_like_ratio(im, bg: ArtColor) -> float:
    """背景相似像素占比：三通道差和 ≤ 3×tol 视为接近，透明像素排除出分母。"""
    Image, ImageChops = _pil()
    if im.mode == "RGBA":
        im = _mask_opaque(im)
    base = Image.new("RGB", im.size, (bg.r, bg.g, bg.b))
    dsum = ImageChops.add(
        ImageChops.add(
            ImageChops.difference(im, base).getchannel("R"),
            ImageChops.difference(im, base).getchannel("G"),
        ),
        ImageChops.difference(im, base).getchannel("B"),
    )
    close = dsum.point(lambda v: 255 if v <= _COLOR_TOL_BUCKET * 3 else 0)
    matched = close.histogram()[255]
    sentinel_img = Image.new("RGB", im.size, _OPAQUE_SENTINEL)
    ssum = ImageChops.add(
        ImageChops.add(
            ImageChops.difference(im, sentinel_img).getchannel("R"),
            ImageChops.difference(im, sentinel_img).getchannel("G"),
        ),
        ImageChops.difference(im, sentinel_img).getchannel("B"),
    )
    is_sentinel = ssum.point(lambda v: 1 if v == 0 else 0)
    sentinel_count = is_sentinel.histogram()[1]
    valid = max(1, im.width * im.height - sentinel_count)
    return matched / valid


@dataclass(frozen=True)
class _BgEstimate:
    background: ArtColor | None
    confidence: float
    uniformity: float
    like_ratio: float


def _estimate_background(im) -> _BgEstimate:
    """8 采样点（4 角 + 4 边中点）估背景：置信度 + 均匀度 + 背景相似占比。"""
    w, h = im.size
    if w <= 2 or h <= 2:
        return _BgEstimate(None, 0.0, 0.0, 0.0)
    mx = max(1, int(w * _BG_SAMPLE_MARGIN))
    my = max(1, int(h * _BG_SAMPLE_MARGIN))
    points = [
        (mx, my),
        (w - 1 - mx, my),
        (mx, h - 1 - my),
        (w - 1 - mx, h - 1 - my),
        (w // 2, my),
        (w // 2, h - 1 - my),
        (mx, h // 2),
        (w - 1 - mx, h // 2),
    ]
    samples: list[tuple[int, int, int]] = []
    for px, py in points:
        c = im.getpixel((px, py))
        if len(c) >= 4 and c[3] < _ALPHA_MIN:
            continue
        samples.append((c[0], c[1], c[2]))
    if not samples:
        return _BgEstimate(None, 0.0, 0.0, 0.0)
    bucket_counts = Counter(
        (s[0] // _PALETTE_BUCKET, s[1] // _PALETTE_BUCKET, s[2] // _PALETTE_BUCKET) for s in samples
    )
    top_bucket, top_n = bucket_counts.most_common(1)[0]
    confidence = top_n / len(samples)
    tops = [
        s
        for s in samples
        if (s[0] // _PALETTE_BUCKET, s[1] // _PALETTE_BUCKET, s[2] // _PALETTE_BUCKET) == top_bucket
    ]
    dists = [max(abs(a[0] - b[0]), abs(a[1] - b[1]), abs(a[2] - b[2])) for a in tops for b in tops]
    mean_d = (sum(dists) / len(dists)) if dists else 0.0
    uniformity = max(0.0, 1.0 - mean_d / 120.0)
    bg = _bucket_rep(top_bucket)
    like_ratio = _background_like_ratio(im, bg)
    return _BgEstimate(bg, round(confidence, 4), round(uniformity, 4), round(like_ratio, 4))


def _page_evidence(im) -> SlidePixelEvidence:
    est = _estimate_background(im)
    small = _downsample(im)
    counts = _pixel_stats(small, est.background)
    return SlidePixelEvidence(
        background=est.background,
        background_confidence=est.confidence if est.background is not None else None,
        background_uniformity=est.uniformity if est.background is not None else None,
        palette=_palette(counts),
        background_like_ratio=est.like_ratio if est.background is not None else None,
    )


def _declared_fg(el: ArtElement) -> ArtColor | None:
    for r in el.runs:
        if r.color is not None:
            return r.color
    return el.foreground


def _text_evidence(region, el: ArtElement, bg: ArtColor | None) -> ElementPixelEvidence:
    fg = _declared_fg(el)
    if fg is None:
        return ElementPixelEvidence(
            color_confidence=0.0, method="unsupported", unsupported_reason="no_declared_fg"
        )
    counts = _pixel_stats(region, bg)
    complexity = _complexity(counts)
    if complexity >= _COMPLEX_THRESHOLD:
        return ElementPixelEvidence(
            foreground=fg,
            background=None,
            background_complexity=round(complexity, 3),
            # complexity>=0.5 here, so 1.0-complexity would be <=0.5 — flat 0.5
            color_confidence=0.5,
            method="complex_background",
        )
    fg_ratio = _match_ratio(counts, fg)
    verified = fg_ratio >= _FG_MATCH_MIN
    return ElementPixelEvidence(
        foreground=fg,
        background=None,  # 组合 PNG 不把区域主色归因为文本背景（冻结契约，#38）
        foreground_match_ratio=round(fg_ratio, 3),
        background_complexity=round(complexity, 3),
        color_confidence=0.85 if verified else 0.3,
        method="declared_verified" if verified else "declared_not_found",
    )


def _shape_evidence(region, el: ArtElement) -> ElementPixelEvidence:
    fill = el.background
    if fill is None or el.has_text() or el.container:
        return ElementPixelEvidence(
            color_confidence=0.0, method="unsupported", unsupported_reason="shape_gate"
        )
    x0 = int(region.width * _INSET)
    y0 = int(region.height * _INSET)
    x1 = int(region.width * (1 - _INSET))
    y1 = int(region.height * (1 - _INSET))
    if x1 - x0 < 2 or y1 - y0 < 2:
        return ElementPixelEvidence(
            color_confidence=0.0, method="unsupported", unsupported_reason="inset_too_small"
        )
    inner = region.crop((x0, y0, x1, y1))
    counts = _pixel_stats(inner, None)
    fill_ratio = _match_ratio(counts, fill)
    complexity = _complexity(counts)
    if fill_ratio >= _FILL_MATCH_MIN and complexity < _COMPLEX_THRESHOLD:
        return ElementPixelEvidence(
            foreground=fill,
            background=None,
            background_match_ratio=round(fill_ratio, 3),
            background_complexity=round(complexity, 3),
            color_confidence=0.85,
            method="center_fill_verified",
        )
    return ElementPixelEvidence(
        foreground=fill,
        background=None,
        background_match_ratio=round(fill_ratio, 3),
        background_complexity=round(complexity, 3),
        color_confidence=0.3,
        method="unsupported",
        unsupported_reason="fill_not_dominant",
    )


def _element_evidence(im, el: ArtElement, bg: ArtColor | None) -> ElementPixelEvidence:
    x0 = int(el.x * im.width)
    y0 = int(el.y * im.height)
    x1 = int((el.x + el.width) * im.width)
    y1 = int((el.y + el.height) * im.height)
    if x1 - x0 < 2 or y1 - y0 < 2:
        return ElementPixelEvidence(
            color_confidence=0.0, method="unsupported", unsupported_reason="region_too_small"
        )
    region = im.crop((max(0, x0), max(0, y0), min(im.width, x1), min(im.height, y1)))
    if el.kind == "text":
        return _text_evidence(region, el, bg)
    if el.kind == "shape":
        return _shape_evidence(region, el)
    return ElementPixelEvidence(
        color_confidence=0.0, method="unsupported", unsupported_reason=f"kind_{el.kind}"
    )


def _replace_pixel_evidence(el: ArtElement, im, bg: ArtColor | None) -> ArtElement:
    return replace(el, pixel_evidence=_element_evidence(im, el, bg))


class PixelEnricher:
    """把 slides_dir 的逐页 PNG 作为像素证据增强 ArtScene（原地 mutate）。

    惰性依赖 Pillow：构造时不 import，首次读图才加载；缺依赖抛可操作错误。
    """

    def __init__(self, slides_dir: str | Path) -> None:
        self._slides_dir = Path(slides_dir)

    def _image_size(self, path: Path) -> tuple[int, int]:
        Image, _ = _pil()
        with Image.open(path) as im:
            return im.size

    def scan(self) -> tuple[dict[int, Path], list[ArtWarning]]:
        """扫描目录：返回 {整数页号: PNG 路径} 与冲突/无页 warning。"""
        if not self._slides_dir.is_dir():
            raise InvalidArgumentError(f"slides_dir 不存在或不是目录: {self._slides_dir}")
        warnings: list[ArtWarning] = []
        seen: dict[int, list[Path]] = {}
        for f in sorted(self._slides_dir.iterdir()):
            m = _SLIDE_RE.match(f.name)
            if not m:
                continue
            seen.setdefault(int(m.group(1)), []).append(f)
        pages: dict[int, Path] = {}
        for idx in sorted(seen):
            files = seen[idx]
            if len(files) > 1:
                warnings.append(
                    ArtWarning(
                        code="art.pixel.page_conflict",
                        message=(
                            f"页 {idx} 存在多个 PNG（{', '.join(f.name for f in files)}），"
                            "该页不进入分析"
                        ),
                    )
                )
                continue
            pages[idx] = files[0]
        if not pages:
            warnings.append(
                ArtWarning(code="art.pixel.no_pages", message="slides_dir 中没有 slide_<n>.png")
            )
        return pages, warnings

    def _read_deck_info(self) -> dict:
        info_path = self._slides_dir / "_deck_info.json"
        if not info_path.is_file():
            return {}
        try:
            return json.loads(info_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}

    def _verify_fingerprint(
        self,
        info: dict,
        expected_sha256: str | None,
        run_id: str | None,
        scene: ArtScene,
    ) -> None:
        """来源指纹校验；中止只通过抛 InvalidArgumentError 表达。"""
        if not info:
            scene.warnings.append(
                ArtWarning(
                    code="art.pixel.source_unverified",
                    message="缺少 _deck_info.json，无法验证像素来源",
                )
            )
            return
        sha = info.get("pptx_sha256")
        if sha is not None and expected_sha256 is not None:
            if sha != expected_sha256:
                raise InvalidArgumentError(
                    f"slides_dir 指纹与当前 PPTX 不一致（{sha[:8]} ≠ {expected_sha256[:8]}），"
                    "拒绝混合来源分析"
                )
            return
        if sha is not None:
            info_run = info.get("run_id")
            if info_run is not None and run_id is not None:
                if info_run != run_id:
                    raise InvalidArgumentError(
                        f"slides_dir 的 run_id 与当前 measurements 不一致（{info_run} ≠ {run_id}）"
                    )
                return
            scene.warnings.append(
                ArtWarning(
                    code="art.pixel.source_unverified",
                    message="slides_dir 指纹缺少可验证 run_id，来源未验证",
                )
            )
            return
        scene.warnings.append(
            ArtWarning(
                code="art.pixel.source_unverified",
                message="slides_dir 指纹缺少 pptx_sha256，来源未验证",
            )
        )
        return

    def enrich(
        self,
        scene: ArtScene,
        *,
        expected_sha256: str | None = None,
        run_id: str | None = None,
    ) -> ArtScene:
        """读 slides_dir 像素证据增强 scene，返回同一 scene（原地 mutate）。"""
        pages, scan_warnings = self.scan()
        for w in scan_warnings:
            if w not in scene.warnings:
                scene.warnings.append(w)
        info = self._read_deck_info()
        self._verify_fingerprint(info, expected_sha256, run_id, scene)
        Image, _ = _pil()
        covered = 0
        for slide in scene.slides:
            path = pages.get(slide.index)
            if path is None:
                scene.warnings.append(
                    ArtWarning(
                        code="art.pixel.page_missing",
                        message=f"页 {slide.index} 无对应 PNG",
                    )
                )
                continue
            try:
                with Image.open(path) as im:
                    im.load()
            except Exception as exc:
                scene.warnings.append(
                    ArtWarning(
                        code="art.pixel.decode_failed",
                        message=f"页 {slide.index} PNG 解码失败: {exc}",
                    )
                )
                continue
            if abs(im.width / im.height - slide.width / slide.height) > _PAGE_ASPECT_TOL:
                scene.warnings.append(
                    ArtWarning(
                        code="art.pixel.aspect_mismatch",
                        message=f"页 {slide.index} 纵横比与场景不符，跳过",
                    )
                )
                continue
            try:
                slide.pixel_evidence = _page_evidence(im)
                slide.elements = [
                    _replace_pixel_evidence(el, im, slide.pixel_evidence.background)
                    for el in slide.elements
                ]
                covered += 1
            except Exception as exc:
                scene.warnings.append(
                    ArtWarning(
                        code="art.pixel.analysis_failed",
                        message=f"页 {slide.index} 像素分析失败: {exc}",
                    )
                )
                continue
        scene_indexes = {s.index for s in scene.slides}
        for idx in sorted(set(pages) - scene_indexes):
            scene.warnings.append(
                ArtWarning(code="art.pixel.page_extra", message=f"PNG 页 {idx} 在场景中不存在")
            )
        if covered > 0:
            scene.sources.add("pixel")
        scene.metadata["pixel_pages_covered"] = covered
        scene.metadata["pixel_pages_total"] = len(scene.slides)
        return scene


def empty_scene_from_slides(slides_dir: str | Path) -> ArtScene:
    """slides_dir-only：空 ArtScene，index/宽高取 PNG 像素尺寸，保留真实页号。"""
    enricher = PixelEnricher(slides_dir)
    pages, warnings = enricher.scan()
    slides: list[ArtSlide] = []
    for idx, p in sorted(pages.items()):
        try:
            width, height = enricher._image_size(p)
        except Exception as exc:
            warnings.append(
                ArtWarning(code="art.pixel.decode_failed", message=f"页 {idx} PNG 解码失败: {exc}")
            )
            continue
        slides.append(ArtSlide(index=idx, width=float(width), height=float(height)))
    indexes = sorted(s.index for s in slides)
    if indexes:
        missing = [i for i in range(indexes[0], indexes[-1] + 1) if i not in set(indexes)]
        if missing:
            warnings.append(
                ArtWarning(
                    code="art.pixel.page_gap",
                    message=f"页号不连续，缺页: {', '.join(map(str, missing))}",
                )
            )
    return ArtScene(slides=slides, width_unit="px", warnings=warnings, sources=set())
