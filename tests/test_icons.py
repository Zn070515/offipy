# tests/test_icons.py
"""图标声明解析 / SVG path 解析 / 注入测试（纯 Python，不依赖 Office）。"""

import math

import pytest

from offipy.icons import (
    _geom_points,
    _parse_path,
    _parse_points_list,
    load_icon_svg,
)

# 注意：IconDecl / parse_icon_declarations 在 Task 3 才实现，届时追加进顶部 import


# ---------- path parser ----------


def test_parse_path_m_l_z():
    sp = _parse_path("M10 10 L20 10 L20 20 Z")
    assert len(sp) == 1
    assert sp[0].close is True
    assert sp[0].points == [(10.0, 10.0), (20.0, 10.0), (20.0, 20.0)]


def test_parse_path_relative():
    sp = _parse_path("M10 10 l5 0 l0 5 z")
    assert len(sp) == 1
    assert sp[0].close is True
    assert sp[0].points == [(10.0, 10.0), (15.0, 10.0), (15.0, 15.0)]


def test_parse_path_hv_commands():
    sp = _parse_path("M0 0 H10 V10 h-5 z")
    assert sp[0].points == [(0.0, 0.0), (10.0, 0.0), (10.0, 10.0), (5.0, 10.0)]


def test_parse_path_cubic_flattens():
    sp = _parse_path("M0 0 C0 10 10 10 10 0 Z")
    assert sp[0].close is True
    assert len(sp[0].points) >= 5  # 压平出多段
    assert sp[0].points[-1] == (10.0, 0.0)
    assert sp[0].points[0] == (0.0, 0.0)


def test_parse_path_quad_flattens():
    sp = _parse_path("M0 0 Q5 10 10 0 Z")
    assert len(sp[0].points) >= 5
    assert sp[0].points[-1] == (10.0, 0.0)


def test_parse_path_arc_flattens():
    sp = _parse_path("M0 0 A5 5 0 0 1 10 0 Z")
    assert len(sp[0].points) >= 5
    assert sp[0].points[-1] == (10.0, 0.0)


def test_parse_path_multi_subpath():
    sp = _parse_path("M0 0 L5 5 Z M10 10 L15 15")
    assert len(sp) == 2
    assert sp[0].close is True
    assert sp[1].close is False
    assert sp[1].points == [(10.0, 10.0), (15.0, 15.0)]


def test_parse_path_smooth_commands():
    # S 反射上一个 C 控制点；T 反射上一个 Q 控制点（不抛错即可）
    sp = _parse_path("M0 0 C0 5 5 5 5 0 S10 -5 10 0 Z")
    assert len(sp) == 1
    sp2 = _parse_path("M0 0 Q2 5 5 0 T10 0 Z")
    assert len(sp2) == 1


def test_parse_path_empty_raises():
    with pytest.raises(ValueError):
        _parse_path("")


def test_parse_path_bad_first_token_raises():
    with pytest.raises(ValueError):
        _parse_path("10 10 L20 20")


# ---------- geometry 元素 ----------


def test_geom_line():
    pts, close = _geom_points("line", {"x1": "0", "y1": "0", "x2": "10", "y2": "5"})
    assert pts == [(0.0, 0.0), (10.0, 5.0)]
    assert close is False


def test_geom_polyline_and_polygon():
    pts, close = _geom_points("polyline", {"points": "6 9 12 15 18 9"})
    assert pts == [(6.0, 9.0), (12.0, 15.0), (18.0, 9.0)]
    assert close is False
    _, close = _geom_points("polygon", {"points": "0 0 4 0 2 3"})
    assert close is True


def test_geom_rect():
    pts, close = _geom_points("rect", {"x": "1", "y": "2", "width": "4", "height": "3"})
    assert close is True
    assert pts == [(1.0, 2.0), (5.0, 2.0), (5.0, 5.0), (1.0, 5.0)]


def test_geom_circle():
    from offipy.icons import _CIRCLE_SAMPLES

    pts, close = _geom_points("circle", {"cx": "5", "cy": "5", "r": "5"})
    assert close is True
    assert len(pts) == _CIRCLE_SAMPLES
    assert abs(sum(x for x, _ in pts) / len(pts) - 5.0) < 1e-6  # 中心≈cx
    assert abs(sum(y for _, y in pts) / len(pts) - 5.0) < 1e-6


def test_parse_points_list():
    assert _parse_points_list("6 9 12 15 18 9") == [(6.0, 9.0), (12.0, 15.0), (18.0, 9.0)]


# ---------- 资产访问 ----------

PH_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" '
    'viewBox="0 0 256 256" fill="currentColor">'
    '<path d="M213.66 82.34a8 8 0 0 0-11.32 0L128 156.69l-74.34-74.35'
    'a8 8 0 0 0-11.32 11.32l80 80a8 8 0 0 0 11.32 0l80-80a8 8 0 0 0 0-11.32Z"/>'
    "</svg>"
)

LU_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" '
    'viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round">'
    '<line x1="12" x2="12" y1="2" y2="22"/>'
    '<polyline points="6 9 12 15 18 9"/>'
    "</svg>"
)


@pytest.fixture
def fake_assets(tmp_path, monkeypatch):
    from offipy import icons

    assets = tmp_path / "assets" / "icons"
    (assets / "phosphor").mkdir(parents=True)
    (assets / "lucide").mkdir(parents=True)
    (assets / "phosphor" / "check.svg").write_text(PH_SVG, encoding="utf-8")
    (assets / "lucide" / "zap.svg").write_text(LU_SVG, encoding="utf-8")
    monkeypatch.setattr(icons, "ASSETS_DIR", assets)
    return assets


def test_load_icon_svg(fake_assets):
    text = load_icon_svg("ph:check")
    assert 'fill="currentColor"' in text
    text = load_icon_svg("lu:zap")
    assert 'stroke="currentColor"' in text


def test_load_icon_svg_missing_raises(fake_assets):
    with pytest.raises(ValueError, match="ph:missing"):
        load_icon_svg("ph:missing")


def test_load_icon_svg_bad_set(fake_assets):
    with pytest.raises(ValueError, match="xx"):
        load_icon_svg("xx:check")


def test_svg_to_subpaths_phosphor(fake_assets):
    from offipy.icons import _svg_to_subpaths

    subpaths, sw = _svg_to_subpaths(load_icon_svg("ph:check"))
    assert len(subpaths) == 1
    assert subpaths[0].close is True
    assert len(subpaths[0].points) >= 3


def test_svg_to_subpaths_lucide(fake_assets):
    from offipy.icons import _svg_to_subpaths

    subpaths, sw = _svg_to_subpaths(load_icon_svg("lu:zap"))
    assert len(subpaths) == 2  # line + polyline
    assert subpaths[0].close is False
    assert sw == 2.0


# ---------- 健壮性（M3-T2 审查补测） ----------


def test_parse_path_truncated_raises():
    with pytest.raises(ValueError):
        _parse_path("M10 10 C0 5 5")


def test_parse_path_degenerate_arc_lineto():
    # 零半径弧按 SVG F.6.2 应 lineto 到终点，不丢终点
    sp = _parse_path("M0 0 A0 0 0 0 1 10 0")
    assert sp[0].points[-1] == (10.0, 0.0)


def test_parse_path_first_not_m_raises():
    with pytest.raises(ValueError):
        _parse_path("L10 10")


def test_geom_ellipse():
    from offipy.icons import _CIRCLE_SAMPLES

    pts, close = _geom_points("ellipse", {"cx": "0", "cy": "0", "rx": "4", "ry": "2"})
    assert close is True
    assert len(pts) == _CIRCLE_SAMPLES
    # 长半轴 4，所有点到中心距离 ∈ [2, 4]
    assert all(2.0 - 1e-9 <= math.hypot(x, y) <= 4.0 + 1e-9 for x, y in pts)


def test_load_icon_svg_invalid_name_raises(fake_assets):
    with pytest.raises(ValueError):
        load_icon_svg("ph:../x")
