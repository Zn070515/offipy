"""Word 纯逻辑单测：常量映射 + 解析辅助（不依赖 Office/COM）。"""

import pytest

from offipy.word import (
    _ALIGN,
    _AUTOFIT,
    _HIGHLIGHT,
    _LINE_SPACING,
    _LINE_STYLE,
    _ORIENTATION,
    _PAGE_NUMBER_ALIGN,
    _PAPER,
    _REPLACE,
    _ROW_HEIGHT_RULE,
    _TABLE_LINE_WIDTH,
    _TABLE_SIDES,
    _UNDERLINE,
    _end_range,
    _resolve_style,
    _resolve_table_sides,
    _rgb,
)


def test_constants_tables():
    # gen_py 实探值：段落对齐 / 下划线 / 高亮 / 行距
    assert _ALIGN["left"] == 0 and _ALIGN["justify"] == 3
    assert _UNDERLINE["single"] == 1 and _UNDERLINE["wavy"] == 11
    assert _HIGHLIGHT["yellow"] == 7 and _HIGHLIGHT["green"] == 11
    assert _LINE_SPACING["single"] == 0 and _LINE_SPACING["double"] == 2
    # 页面 / 页码 / 替换
    assert _ORIENTATION["landscape"] == 1
    assert _PAPER["a4"] == 7 and _PAPER["letter"] == 2
    assert _PAGE_NUMBER_ALIGN["center"] == 1
    assert _REPLACE["all"] == 2
    # 表格
    assert _LINE_STYLE["single"] == 1 and _LINE_STYLE["none"] == 0
    assert _TABLE_SIDES["top"] == 1 and _TABLE_SIDES["left"] == 2 and _TABLE_SIDES["inside-v"] == 6
    assert _ROW_HEIGHT_RULE["at_least"] == 1
    assert _AUTOFIT["content"] == 1
    assert _TABLE_LINE_WIDTH["1pt"] == 8 and _TABLE_LINE_WIDTH["6pt"] == 48


def test_rgb_converts_to_com_long():
    # Word/Excel 共用 COLORREF 公式：R 在最低字节
    assert _rgb("#FF0000") == 255
    assert _rgb("#2251FF") == 0xFF5122  # R=0x22, G=0x51, B=0xFF


def test_resolve_style_valid_and_case_insensitive():
    assert _resolve_style("center", _ALIGN, "对齐") == 1
    assert _resolve_style("Double", _LINE_SPACING, "行距") == 2
    assert _resolve_style("a4", _PAPER, "纸张") == 7
    assert _resolve_style("right", _PAGE_NUMBER_ALIGN, "页码对齐") == 2


def test_resolve_style_unknown_raises():
    with pytest.raises(ValueError):
        _resolve_style("middle", _ALIGN, "对齐")


def test_resolve_table_sides_variants():
    assert _resolve_table_sides("all") == [1, 2, 3, 4, 5, 6]
    assert _resolve_table_sides(None) == [1, 2, 3, 4, 5, 6]
    assert _resolve_table_sides("outside") == [1, 2, 3, 4]
    assert _resolve_table_sides("inside") == [5, 6]
    assert _resolve_table_sides("left,top") == [2, 1]


def test_resolve_table_sides_unknown_raises():
    with pytest.raises(ValueError):
        _resolve_table_sides("diagonal")


def test_end_range_returns_collapsed_end():
    from unittest.mock import Mock

    doc = Mock()
    doc.Content.Collapse = Mock()
    rng = _end_range(doc)
    doc.Content.Collapse.assert_called_once_with(0)  # wdCollapseEnd
    assert rng is doc.Content
