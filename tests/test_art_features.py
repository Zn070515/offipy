from offipy.art.features import (
    AlignmentLine,
    _drifted_count,
    _gaps,
    _row_clusters,
    alignment_features,
    density_features,
    physical_aspect_ratio,
    spacing_features,
    visual_mass,
)
from offipy.art.models import ArtElement


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
