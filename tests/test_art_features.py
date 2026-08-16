import pytest

from art_helpers import make_element, make_image_element, make_slide, make_text_element
from offipy.art.features import (
    AlignmentLine,
    _drifted_count,
    _gaps,
    _row_clusters,
    accent_elements,
    alignment_features,
    compute_features,
    density_features,
    effective_background,
    element_color_weights,
    focus_features,
    font_hierarchy,
    infer_slide_role,
    is_accent,
    palette_features,
    physical_aspect_ratio,
    spacing_features,
    visual_mass,
)
from offipy.art.models import ArtColor, ArtElement, ArtTextRun


def _el(element_id, x, y, w, h, role="body", kind="text"):
    return ArtElement(
        element_id=element_id, kind=kind, role=role, x=x, y=y, width=w, height=h, slide_index=1
    )


def test_alignment_line():
    line = AlignmentLine(position=0.1, members=["a", "b", "c"], axis="vertical")
    assert line.position == 0.1
    assert len(line.members) == 3
    assert line.axis == "vertical"


def test_alignment_detects_shared_left_edge():
    els = [
        _el("a", 0.1, 0.1, 0.2, 0.1),
        _el("b", 0.1, 0.3, 0.2, 0.1),
        _el("c", 0.1, 0.5, 0.2, 0.1),
    ]
    feats = alignment_features(els)
    lines = feats["lines"]
    assert any(ln.axis == "vertical" and abs(ln.position - 0.1) < 0.01 for ln in lines)


def test_alignment_needs_three_members():
    els = [_el("a", 0.1, 0.1, 0.2, 0.1), _el("b", 0.1, 0.3, 0.2, 0.1)]
    assert alignment_features(els)["lines"] == []


def test_gaps_uses_ordered_elements():
    els = [_el("a", 0.0, 0.0, 0.2, 0.1), _el("b", 0.4, 0.0, 0.2, 0.1), _el("c", 0.8, 0.0, 0.2, 0.1)]
    gaps = _gaps(els, axis="x")
    # gap = 前一个的 right 到后一个的 left：0.4-(0.0+0.2)=0.2，0.8-(0.4+0.2)=0.2
    assert gaps == [0.2, 0.2]


def test_gaps_uses_centers_for_negative_span():
    els = [_el("a", 0.0, 0.0, 0.5, 0.1), _el("b", 0.3, 0.0, 0.5, 0.1)]
    gaps = _gaps(els, axis="x")
    assert gaps == [0.0]  # span 重叠时退回中心距并 clamp 到 0


def test_row_clusters_group_by_center_y():
    els = [
        _el("a", 0.0, 0.0, 0.2, 0.1),
        _el("b", 0.4, 0.01, 0.2, 0.1),  # 与 a 同行
        _el("c", 0.0, 0.4, 0.2, 0.1),  # 另一行
    ]
    rows = _row_clusters(els)
    assert len(rows) == 2
    assert len(rows[0]) == 2  # a,b 同行


def test_spacing_drift_detected_within_row():
    els = [
        _el("a", 0.0, 0.0, 0.2, 0.1),
        _el("b", 0.4, 0.0, 0.2, 0.1),
        _el("c", 0.8, 0.0, 0.2, 0.1),
        _el("d", 1.5, 0.0, 0.2, 0.1),  # 与 c 间距 0.5，偏离中位 0.2
    ]
    feats = spacing_features(els)
    assert feats["horizontal"]["drift_count"] >= 1


def test_spacing_regular_no_drift():
    els = [
        _el("a", 0.0, 0.0, 0.2, 0.1),
        _el("b", 0.4, 0.0, 0.2, 0.1),
        _el("c", 0.8, 0.0, 0.2, 0.1),
    ]
    feats = spacing_features(els)
    assert feats["horizontal"]["drift_count"] == 0


def test_spacing_cross_row_no_false_drift():
    # 不同行的元素不产生水平相邻（避免跨行误报）
    els = [
        _el("a", 0.0, 0.0, 0.2, 0.1),
        _el("b", 0.8, 0.0, 0.2, 0.1),
        _el("c", 0.1, 0.4, 0.2, 0.1),  # 第二行，x 与 a 接近但不同行
    ]
    feats = spacing_features(els)
    assert feats["horizontal"]["drift_count"] == 0


def test_drifted_count_median_tolerance():
    values = [0.2, 0.2, 0.5]
    assert _drifted_count(values) == 1


def test_visual_mass_ink_text():
    els = [
        ArtElement(
            element_id="t",
            kind="text",
            role="title",
            x=0.0,
            y=0.0,
            width=0.5,
            height=0.1,
            slide_index=1,
            text="Hello",
            font_size_norm=0.05,
        ),
    ]
    # ink = max(len(text),1) * font_size_norm = 5 * 0.05 = 0.25
    assert abs(visual_mass(els)["ink"] - 0.25) < 1e-9
    assert visual_mass(els)["dominant_id"] == "t"


def test_visual_mass_filters_background():
    els = [
        ArtElement(
            element_id="bg",
            kind="shape",
            role="background",
            x=0.0,
            y=0.0,
            width=1.0,
            height=1.0,
            slide_index=1,
            text="x",
            font_size_norm=0.05,
            is_background=True,
        ),
        ArtElement(
            element_id="t",
            kind="text",
            role="body",
            x=0.0,
            y=0.0,
            width=0.5,
            height=0.1,
            slide_index=1,
            text="Hi",
            font_size_norm=0.05,
        ),
    ]
    # bg 被排除，只剩 text：max(2,1)*0.05 = 0.1
    assert abs(visual_mass(els)["ink"] - 0.1) < 1e-9


def test_visual_mass_dominant_is_max_contributor():
    # 图片面积大但 ink 贡献小（area×2.0 高于文字），dominant 取贡献最大者
    els = [
        ArtElement(
            element_id="small",
            kind="text",
            role="title",
            x=0.0,
            y=0.0,
            width=0.5,
            height=0.1,
            slide_index=1,
            text="Hello",
            font_size_norm=0.05,
        ),  # 贡献 0.25
        ArtElement(
            element_id="big",
            kind="shape",
            role="body",
            x=0.0,
            y=0.2,
            width=0.3,
            height=0.1,
            slide_index=1,
        ),  # 贡献 0.015
    ]
    assert visual_mass(els)["dominant_id"] == "small"


def test_density_union_area():
    els = [
        ArtElement(
            element_id="a",
            kind="shape",
            role="body",
            x=0.0,
            y=0.0,
            width=0.5,
            height=0.5,
            slide_index=1,
        ),
        ArtElement(
            element_id="b",
            kind="shape",
            role="body",
            x=0.25,
            y=0.0,
            width=0.5,
            height=0.5,
            slide_index=1,
        ),
    ]
    feats = density_features(els)
    assert abs(feats["sum_area_ratio"] - (0.25 + 0.25)) < 1e-9
    assert abs(feats["union_area_ratio"] - 0.375) < 1e-9  # 重叠 0.125，并 0.375


def test_density_union_area_with_y_gap():
    # 垂直方向有空白 → 并集不能把空白算进去
    els = [
        ArtElement(
            element_id="a",
            kind="shape",
            role="body",
            x=0.0,
            y=0.0,
            width=0.5,
            height=0.2,
            slide_index=1,
        ),
        ArtElement(
            element_id="b",
            kind="shape",
            role="body",
            x=0.0,
            y=0.5,
            width=0.5,
            height=0.2,
            slide_index=1,
        ),
    ]
    feats = density_features(els)
    assert abs(feats["union_area_ratio"] - 0.2) < 1e-9  # 0.5×0.2 + 0.5×0.2


def test_physical_aspect_ratio_from_norm():
    # width=0.5, height=0.3，页 1920×1080 → 物理 960×324 → ratio 2.963
    el = ArtElement(
        element_id="i",
        kind="image",
        role="image",
        x=0.0,
        y=0.0,
        width=0.5,
        height=0.3,
        slide_index=1,
    )
    r = physical_aspect_ratio(el, page_width=1920.0, page_height=1080.0)
    assert abs(r - (960.0 / 324.0)) < 1e-6


def test_infer_slide_role_cover_title_subtitle():
    # 标题 + 副标题，正文极少（≤1）→ cover（不再用「元素数 ≤3」作硬条件）
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=52.0, role="title"),
            make_text_element("s", "Sub", y=0.3, font_size=26.0, role="subtitle"),
        ],
    )
    assert infer_slide_role(slide) == "cover"


def test_infer_slide_role_content():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", y=0.05, font_size=40.0),
            make_text_element("b1", "B1", y=0.2, font_size=24.0),
            make_text_element("b2", "B2", y=0.3, font_size=24.0),
            make_image_element("i", y=0.4, w=0.4, h=0.3),
        ],
    )
    assert infer_slide_role(slide) == "content"


def test_infer_slide_role_explicit_marker_wins():
    # 显式 role 标记优先于启发式
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=48.0, role="section"),
            make_text_element("b", "Body", y=0.2, font_size=24.0),
        ],
    )
    assert infer_slide_role(slide) == "section"


def test_infer_slide_role_gallery_when_three_images():
    slide = make_slide(
        1,
        elements=[
            make_image_element("i1", y=0.1),
            make_image_element("i2", y=0.4),
            make_image_element("i3", y=0.7),
        ],
    )
    assert infer_slide_role(slide) == "gallery"


def test_effective_background_element_first():
    el = make_element("t", kind="text", role="body", text="Hi", background=ArtColor(240, 240, 240))
    slide = make_slide(1, elements=[el], background_color=ArtColor(255, 255, 255))
    bg = effective_background(el, slide)
    assert (bg.r, bg.g, bg.b) == (240, 240, 240)


def test_effective_background_falls_back_to_slide():
    el = make_text_element("t", "Hi", font_size=20.0)  # 无元素背景
    slide = make_slide(1, elements=[el], background_color=ArtColor(255, 255, 255))
    bg = effective_background(el, slide)
    assert (bg.r, bg.g, bg.b) == (255, 255, 255)


def test_effective_background_unknown_is_none():
    # 元素无背景 + 页面无背景 → None（不默认白，规则降 coverage）
    el = make_text_element("t", "Hi", font_size=20.0)
    slide = make_slide(1, elements=[el], background_color=None)
    assert effective_background(el, slide) is None


def test_font_hierarchy_ratio():
    slide = make_slide(
        1,
        elements=[
            make_text_element("title", "Title", font_size=48.0, role="title"),
            make_text_element("body", "Body", y=0.2, font_size=24.0, role="body"),
        ],
    )
    fh = font_hierarchy(slide)
    assert abs(fh["ratio"] - 2.0) < 1e-9
    assert fh["title_id"] == "title"


def test_palette_accent_ratio_area_weighted():
    slide = make_slide(
        1,
        elements=[
            make_text_element("a", "A", font_size=20.0, foreground=ArtColor(230, 0, 0)),
            make_text_element("b", "B", font_size=20.0, foreground=ArtColor(230, 0, 0), w=0.4),
            make_text_element("c", "C", font_size=20.0, foreground=ArtColor(30, 30, 30)),
        ],
    )
    pal = palette_features(slide)
    # accent 两元素面积 (0.2*0.08 + 0.4*0.08)，非 accent 0.2*0.08
    assert pal["accent_ratio"] > 0.5
    assert pal["dominant"][0] > 0


def test_palette_uses_foreground_not_background():
    slide = make_slide(
        1,
        elements=[
            make_text_element(
                "a",
                "A",
                font_size=20.0,
                foreground=ArtColor(30, 30, 30),
                background=ArtColor(230, 0, 0),
            ),
        ],
    )
    pal = palette_features(slide)
    assert pal["accent_ratio"] == 0.0  # 红色背景不算强调色


def test_palette_rgb_bucket():
    c = ArtColor(20, 21, 22)
    assert (c.r // 32) * 32 == 0
    assert (ArtColor(250, 250, 250).r // 32) * 32 == 224


def test_focus_requires_three_and_ratio():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=48.0, role="title"),
            make_text_element("b", "Body", y=0.2, font_size=48.0, role="body"),
            make_text_element("c", "Cap", y=0.3, font_size=48.0, role="caption"),
        ],
    )
    assert focus_features(slide)["has_focus"] is False  # ratio < 1.5


def test_focus_true_with_dominant():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=72.0, role="title"),
            make_text_element("b1", "B1", y=0.2, font_size=20.0, role="body"),
            make_text_element("b2", "B2", y=0.3, font_size=20.0, role="body"),
        ],
    )
    assert focus_features(slide)["has_focus"] is True


def test_compute_features_no_width_unit():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=48.0, role="title"),
            make_text_element("b", "Body", y=0.2, font_size=24.0, role="body"),
        ],
    )
    feats = compute_features(slide)
    sig = feats["page_signature"]
    assert "role" in sig and "dominant_id" in sig
    assert "width_unit" not in sig


def test_make_slide_sets_element_slide_index():
    # 模型不变量：make_slide 后 element.slide_index 必须等于 slide.index
    slide = make_slide(
        3,
        elements=[
            make_text_element("t", "Title", font_size=48.0),
            make_text_element("b", "Body", y=0.2, font_size=24.0),
        ],
    )
    assert slide.index == 3
    assert all(e.slide_index == slide.index for e in slide.elements)


def test_element_color_weights_run_text_length():
    el = make_text_element(
        "t",
        "Hello",
        font_size=20.0,
        foreground=ArtColor(30, 30, 30),
        runs=[
            ArtTextRun(text="He", font_size=20.0, font_size_unit="px", color=ArtColor(230, 0, 0)),
            ArtTextRun(text="llo", font_size=20.0, font_size_unit="px", color=ArtColor(30, 30, 30)),
        ],
    )
    weights = element_color_weights(el)
    assert weights == [(ArtColor(230, 0, 0), 2.0), (ArtColor(30, 30, 30), 3.0)]


def test_element_color_weights_fallback_foreground():
    el = ArtElement(
        element_id="s",
        kind="shape",
        role="body",
        x=0.1,
        y=0.1,
        width=0.2,
        height=0.2,
        slide_index=1,
        foreground=ArtColor(230, 0, 0),
    )
    assert element_color_weights(el) == [(ArtColor(230, 0, 0), 1.0)]


def test_element_color_weights_empty():
    el = ArtElement(
        element_id="s",
        kind="shape",
        role="body",
        x=0.1,
        y=0.1,
        width=0.2,
        height=0.2,
        slide_index=1,
    )
    assert element_color_weights(el) == []


def test_is_accent_public():
    assert is_accent(ArtColor(230, 0, 0)) is True
    assert is_accent(ArtColor(30, 30, 30)) is False


def test_accent_elements_skips_skip_roles_and_zero_area():
    slide = make_slide(
        1,
        elements=[
            make_text_element("body", "B", font_size=20.0, foreground=ArtColor(30, 30, 30)),
            make_text_element(
                "pn", "1", font_size=20.0, role="page_number", foreground=ArtColor(30, 30, 30)
            ),
            make_text_element(
                "ft", "f", font_size=20.0, role="footer", foreground=ArtColor(30, 30, 30)
            ),
        ],
    )
    ids = [e.element_id for e in accent_elements(slide)]
    assert ids == ["body"]


def test_palette_accent_ratio_run_weighted():
    # run 级加权核心路径：Hello(accent, 5 字) + World(中性, 5 字) → 面积均分 0.5
    slide = make_slide(
        1,
        elements=[
            make_text_element(
                "t",
                "HelloWorld",
                font_size=20.0,
                foreground=ArtColor(30, 30, 30),
                runs=[
                    ArtTextRun(
                        text="Hello", font_size=20.0, font_size_unit="px", color=ArtColor(230, 0, 0)
                    ),
                    ArtTextRun(
                        text="World",
                        font_size=20.0,
                        font_size_unit="px",
                        color=ArtColor(30, 30, 30),
                    ),
                ],
            ),
        ],
    )
    pal = palette_features(slide)
    assert pal["accent_ratio"] == pytest.approx(0.5, rel=1e-3)
