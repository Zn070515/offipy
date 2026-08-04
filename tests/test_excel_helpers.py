"""Excel 纯逻辑单测：常量映射 + 解析辅助（不依赖 Office/COM）。"""
import pytest

from offipy.excel import (
    _BORDER_INDEX,
    _BORDER_WEIGHT,
    _COND_OPERATOR,
    _LINE_STYLE,
    _ORIENTATION,
    _PAPER_SIZE,
    _resolve_sides,
    _resolve_style,
)


def test_constants_tables():
    assert _BORDER_INDEX["left"] == 7 and _BORDER_INDEX["inside-h"] == 12
    assert _LINE_STYLE["double"] == -4119 and _LINE_STYLE["none"] == -4142
    assert _BORDER_WEIGHT["hairline"] == 1 and _BORDER_WEIGHT["medium"] == -4138
    assert _COND_OPERATOR["greater"] == 5 and _COND_OPERATOR["between"] == 1
    assert _ORIENTATION["landscape"] == 2
    assert _PAPER_SIZE["a4"] == 9


def test_resolve_sides_single():
    assert _resolve_sides("top") == [8]
    assert _resolve_sides("inside-h") == [12]


def test_resolve_sides_comma():
    assert _resolve_sides("top,bottom") == [8, 9]


def test_resolve_sides_all():
    assert _resolve_sides("all") == [7, 8, 9, 10, 11, 12]
    assert _resolve_sides(None) == [7, 8, 9, 10, 11, 12]


def test_resolve_sides_outside_inside():
    assert _resolve_sides("outside") == [7, 8, 9, 10]
    assert _resolve_sides("inside") == [11, 12]


def test_resolve_sides_unknown_raises():
    with pytest.raises(ValueError):
        _resolve_sides("top-left")


def test_resolve_style_valid():
    assert _resolve_style("thick", _BORDER_WEIGHT, "线宽") == 4
    assert _resolve_style("continuous", _LINE_STYLE, "线型") == 1
    assert _resolve_style("greater", _COND_OPERATOR, "条件格式运算符") == 5
    assert _resolve_style("landscape", _ORIENTATION, "页面方向") == 2
    assert _resolve_style("a4", _PAPER_SIZE, "纸张") == 9


def test_resolve_style_case_insensitive():
    assert _resolve_style("Thin", _BORDER_WEIGHT, "线宽") == 2


def test_resolve_style_unknown_raises():
    with pytest.raises(ValueError):
        _resolve_style("super-thick", _BORDER_WEIGHT, "线宽")
