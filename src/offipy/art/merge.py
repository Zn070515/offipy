"""场景融合：测量为主、几何审计为副，一对一匹配 + 匹配置信度 + 证据合并。

rev2.1：不依赖 DOM ID 与 PPTX ID 相同；用 matched 集合保证一对一；
未匹配的 secondary 元素与 secondary-only 页面都保留并附 warning。
rev2.2：双源 role 词表不一致（measurement title/body/subtitle ↔ audit
content/footer）导致 100% 未匹配。匹配改为文本强佐证 + 几何兜底，role 作
软信号：归一 role 相同且文本相同 → 身份 0.8；文本相同（跨词表 role）→ 0.7；
双方都有文本但不同 → 不合并；至少一边无文本 → 几何 + role 加分。
"""

from __future__ import annotations

from typing import Any, cast

from .models import ArtElement, ArtScene, ArtSlide, ArtWarning

_ROLE_ALIASES = {
    # audit 词表 → art 词表（audit 非占位符文本统一 content，未知形状 unknown）
    "content": "body",
    "unknown": "shape",
}


def _norm_role(role: str) -> str:
    return _ROLE_ALIASES.get(role, role)


def _norm_text(text: str) -> str:
    """文本归一：软换行 \\x0b ↔ \\n 统一，折叠空白，小写。"""
    return " ".join(text.replace("\x0b", "\n").split()).lower()


def _match_confidence(primary_el: ArtElement, secondary_el: ArtElement) -> float | None:
    """匹配置信度：shape_id 精确=1.0；身份=0.8；文本强佐证=0.7；几何=0.5×(1-d/0.2)。

    文本身份分支（0.8/0.7）不设距离门——文本是强信号，跨源几何可能因坐标参考
    系差异而偏移；纯几何分支（至少一边无文本）受距离门约束以防误并。
    """
    if primary_el.element_id == secondary_el.element_id:
        return 1.0
    p_text = _norm_text(primary_el.text)
    s_text = _norm_text(secondary_el.text)
    same_text = bool(p_text) and p_text == s_text
    same_role = _norm_role(primary_el.role) == _norm_role(secondary_el.role)
    if same_role and same_text:
        return 0.8
    if same_text:
        return 0.7
    pcx, pcy = primary_el.x + primary_el.width / 2, primary_el.y + primary_el.height / 2
    ecx, ecy = secondary_el.x + secondary_el.width / 2, secondary_el.y + secondary_el.height / 2
    d = cast("float", ((pcx - ecx) ** 2 + (pcy - ecy) ** 2) ** 0.5)
    # 双方都有文本但不同 → 非同一元素（即使几何接近也不合并）
    if p_text and s_text:
        return None
    if d > 0.2:
        return None
    # 无文本佐证：几何底分先过门槛，role 加分不能救回相距过远的对
    base = 0.5 * (1.0 - d / 0.2)
    if base <= 0.05:
        return None
    if same_role:
        base += 0.1
    return base


def _merge_element(
    primary_el: ArtElement, secondary_el: ArtElement | None, confidence: float | None
) -> ArtElement:
    """重建新元素（ArtElement frozen）：source/evidence 写入新实例。"""
    evidence: dict[str, Any] = {"measurement": _snapshot(primary_el)}
    source = "measurement"
    if secondary_el is not None:
        evidence["pptx"] = _snapshot(secondary_el)
        evidence["match_confidence"] = confidence
        source = "merged"
    return ArtElement(
        element_id=primary_el.element_id,
        kind=primary_el.kind,
        role=primary_el.role,
        x=primary_el.x,
        y=primary_el.y,
        width=primary_el.width,
        height=primary_el.height,
        slide_index=primary_el.slide_index,
        foreground=primary_el.foreground,
        background=primary_el.background,
        border=primary_el.border,
        is_background=primary_el.is_background,
        text=primary_el.text,
        font_size=primary_el.font_size,
        font_size_unit=primary_el.font_size_unit,
        font_size_norm=primary_el.font_size_norm,
        runs=primary_el.runs,
        natural_width=primary_el.natural_width,
        natural_height=primary_el.natural_height,
        source=source,
        evidence=evidence,
        container=primary_el.container,
        decoration=primary_el.decoration,
    )


def merge_scenes(primary: ArtScene, secondary: ArtScene) -> tuple[ArtScene, list[ArtWarning]]:
    """以 primary（测量）为主场景，把 secondary（pptx 审计）一对一合并。

    对 primary 每个元素，在未匹配的 secondary 元素里选匹配置信度最高者；
    未匹配的 secondary 元素追加（warning art.merge.unmatched）；
    secondary 独有页面保留（warning art.merge.slide_secondary_only）。
    主场景单位（px）为准；pptx 证据只作对照，规则不混用单位比较。
    """
    warnings: list[ArtWarning] = []
    out_slides: list[ArtSlide] = []
    for ps in primary.slides:
        ss = secondary.by_slide(ps.index)
        # 每页独立匹配集合：不跨页复用 secondary 元素
        matched_secondary: set[str] = set()
        elements: list[ArtElement] = []
        if ss is None:
            warnings.append(
                ArtWarning(
                    code="art.merge.slide_missing",
                    message=f"secondary 缺 slide {ps.index}",
                )
            )
            elements.extend(_merge_element(el, None, None) for el in ps.elements)
            out_slides.append(
                ArtSlide(
                    index=ps.index,
                    width=ps.width,
                    height=ps.height,
                    elements=elements,
                    background_color=ps.background_color,
                )
            )
            continue
        for el in ps.elements:
            best = None
            best_c = 0.0
            for se in ss.elements:
                if se.element_id in matched_secondary:
                    continue
                c = _match_confidence(el, se)
                if c is not None and c > best_c:
                    best_c = c
                    best = se
            if best is not None:
                matched_secondary.add(best.element_id)
            elements.append(_merge_element(el, best, best_c if best else None))
        # secondary 未匹配元素 → 追加 + warning
        for se in ss.elements:
            if se.element_id not in matched_secondary:
                elements.append(se)
                warnings.append(
                    ArtWarning(
                        code="art.merge.unmatched",
                        message=f"secondary 元素 {se.element_id} 未匹配，已追加",
                    )
                )
                matched_secondary.add(se.element_id)
        out_slides.append(
            ArtSlide(
                index=ps.index,
                width=ps.width,
                height=ps.height,
                elements=elements,
                background_color=ps.background_color,
            )
        )
    # secondary 独有页面 → 保留 + warning
    for ss in secondary.slides:
        if primary.by_slide(ss.index) is None:
            out_slides.append(
                ArtSlide(
                    index=ss.index,
                    width=ss.width,
                    height=ss.height,
                    elements=list(ss.elements),
                    background_color=ss.background_color,
                )
            )
            warnings.append(
                ArtWarning(
                    code="art.merge.slide_secondary_only",
                    message=f"secondary 独有 slide {ss.index} 已保留",
                )
            )
    out_slides.sort(key=lambda s: s.index)
    scene = ArtScene(
        slides=out_slides,
        width_unit=primary.width_unit,
        warnings=list(primary.warnings) + list(secondary.warnings) + warnings,
        sources=set(primary.sources) | set(secondary.sources),
    )
    return scene, warnings


def _snapshot(el: ArtElement) -> dict[str, Any]:
    return {
        "font_size": el.font_size,
        "font_size_unit": el.font_size_unit,
        "font_size_norm": el.font_size_norm,
        "x": el.x,
        "y": el.y,
        "width": el.width,
        "height": el.height,
        "foreground": el.foreground.to_dict() if el.foreground else None,
        "background": el.background.to_dict() if el.background else None,
    }
