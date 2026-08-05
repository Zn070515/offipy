"""PPTX Shape Role 分类（零第三方依赖）。

内部角色：background / header / footer / page_number / title / content /
decoration / unknown。Role 不作为绝对事实，只用于调整规则与抑制误报。

判定依据：placeholder 类型、位置、尺寸、z-order、文本是否纯数字、是否覆盖整页、
顶部/底部区域、跨页重复指纹（shape_type + normalized_text + 位置桶 + 尺寸桶）。

分类分两遍：
1. 单页可判定：placeholder（title/page_number/header/footer/content）、全页背景
   （无文本 + 覆盖≥90% + 近中心 + z-order 低）、几何页码（纯数字 + 底部 15% +
   小尺寸）。
2. 跨页重复（需全部页上下文）：非 placeholder 页眉页脚（顶部/底部区域 + 文本
   相同 + 指纹重复 >60% 页）、重复装饰（指纹重复 >60% 页）。

不能全局忽略所有纯数字短文本——只有底部小尺寸的纯数字才算页码。
"""

from __future__ import annotations

from collections import defaultdict

from .extract import _ShapeRecord

_TOP_BOTTOM_FRACTION = 0.15
_COVERAGE_THRESHOLD = 0.9
_CENTER_TOLERANCE_FRACTION = 0.15
_MAX_BACKGROUND_Z_ORDER = 2
_PAGE_NUMBER_MAX_WIDTH = 2.0
_PAGE_NUMBER_MAX_HEIGHT = 0.6
_FINGERPRINT_BUCKET = 0.5  # 英寸

_PLACEHOLDER_ROLES = {
    "SLIDE_NUMBER": "page_number",
    "HEADER": "header",
    "FOOTER": "footer",
    "DATE": "footer",
    "TITLE": "title",
    "CENTER_TITLE": "title",
    "VERTICAL_TITLE": "title",
}

_REPEAT_EXEMPT_ROLES = ("title", "page_number", "background", "header", "footer")


def _normalize_text(text: str) -> str:
    return " ".join(text.split()).lower()


def _placeholder_role(placeholder_type: str) -> str:
    return _PLACEHOLDER_ROLES.get(placeholder_type, "content")


def classify_presentation(records: list[_ShapeRecord], slide_size: tuple[float, float]) -> None:
    """就地给每个 record.role 赋值（需要全部页上下文）。"""
    slide_w, slide_h = slide_size
    for rec in records:
        rec.role = _classify_single(rec, slide_w, slide_h)
    slide_count = len({r.slide_index for r in records})
    _assign_repeated(records, slide_w, slide_h, slide_count)


def _classify_single(rec: _ShapeRecord, slide_w: float, slide_h: float) -> str:
    if rec.placeholder_type:
        r = _placeholder_role(rec.placeholder_type)
        if r in ("title", "page_number", "header", "footer"):
            return r
        return "content"
    if _is_full_bleed(rec, slide_w, slide_h):
        return "background"
    if _is_page_number(rec, slide_h):
        return "page_number"
    if rec.has_text_frame and rec.text.strip():
        return "content"
    return "unknown"


def _is_full_bleed(rec: _ShapeRecord, slide_w: float, slide_h: float) -> bool:
    if rec.is_group or rec.is_connector or rec.parent_shape_id is not None:
        return False
    if rec.left is None or rec.top is None or rec.width is None or rec.height is None:
        return False
    if rec.text.strip():
        return False  # 有实质文本 → 是内容容器而非背景
    if rec.width * rec.height / (slide_w * slide_h) < _COVERAGE_THRESHOLD:
        return False
    cx = rec.left + rec.width / 2.0
    cy = rec.top + rec.height / 2.0
    if abs(cx - slide_w / 2.0) > _CENTER_TOLERANCE_FRACTION * slide_w:
        return False
    if abs(cy - slide_h / 2.0) > _CENTER_TOLERANCE_FRACTION * slide_h:
        return False
    return rec.z_order <= _MAX_BACKGROUND_Z_ORDER


def _is_page_number(rec: _ShapeRecord, slide_h: float) -> bool:
    if rec.left is None or rec.top is None or rec.width is None or rec.height is None:
        return False
    if not rec.text.strip().isdigit():
        return False
    if rec.width > _PAGE_NUMBER_MAX_WIDTH or rec.height > _PAGE_NUMBER_MAX_HEIGHT:
        return False
    cy = rec.top + rec.height / 2.0
    return cy > (1 - _TOP_BOTTOM_FRACTION) * slide_h


def _assign_repeated(
    records: list[_ShapeRecord], slide_w: float, slide_h: float, slide_count: int
) -> None:
    if slide_count <= 1:
        return
    fp_slides: dict[str, set[int]] = defaultdict(set)
    for rec in records:
        if rec.role in _REPEAT_EXEMPT_ROLES:
            continue
        fp = _fingerprint(rec)
        if fp is not None:
            fp_slides[fp].add(rec.slide_index)
    for rec in records:
        if rec.role not in ("unknown", "content"):
            continue
        fp = _fingerprint(rec)
        if fp is None or fp not in fp_slides:
            continue
        if len(fp_slides[fp]) * 5 <= slide_count * 3:  # 需严格 >60% 页
            continue
        cy = rec.top + rec.height / 2.0 if rec.top is not None and rec.height is not None else None
        if cy is not None and rec.text.strip():
            if cy < _TOP_BOTTOM_FRACTION * slide_h:
                rec.role = "header"
            elif cy > (1 - _TOP_BOTTOM_FRACTION) * slide_h:
                rec.role = "footer"
            else:
                rec.role = "decoration"
        else:
            rec.role = "decoration"


def _fingerprint(rec: _ShapeRecord) -> str | None:
    if (
        rec.is_group
        or rec.left is None
        or rec.top is None
        or rec.width is None
        or rec.height is None
    ):
        return None
    x_c = round((rec.left + rec.width / 2.0) / _FINGERPRINT_BUCKET)
    y_c = round((rec.top + rec.height / 2.0) / _FINGERPRINT_BUCKET)
    w_b = round(rec.width / _FINGERPRINT_BUCKET)
    h_b = round(rec.height / _FINGERPRINT_BUCKET)
    return f"{rec.shape_type}|{_normalize_text(rec.text)}|{x_c},{y_c}|{w_b}x{h_b}"
