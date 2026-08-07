"""Excel 纯逻辑单测：常量映射 + 解析辅助（不依赖 Office/COM）。"""

import pytest

from offipy.excel import (
    _BORDER_INDEX,
    _BORDER_WEIGHT,
    _COND_OPERATOR,
    _LINE_STYLE,
    _ORIENTATION,
    _PAPER_SIZE,
    ExcelApp,
    _parse_cell,
    _resolve_sides,
    _resolve_style,
)
from offipy.exceptions import InvalidArgumentError


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


# --- _parse_cell 收严（P0/§11：Excel 真实坐标，越界/畸形一律拒绝） ---


def test_parse_cell_valid():
    assert _parse_cell("A1") == (1, 1)
    assert _parse_cell("Z26") == (26, 26)
    assert _parse_cell("AA1") == (1, 27)
    assert _parse_cell("XFD1048576") == (1048576, 16384)  # 上限值合法


def test_parse_cell_col_out_of_bounds():
    with pytest.raises(InvalidArgumentError):
        _parse_cell("XFE1")  # XFE > XFD(16384)


def test_parse_cell_row_out_of_bounds():
    with pytest.raises(InvalidArgumentError):
        _parse_cell("A1048577")  # > 1048576
    with pytest.raises(InvalidArgumentError):
        _parse_cell("A0")  # 0 基行非法


def test_parse_cell_malformed_rejected():
    # 回归：'A1B2' 曾按字母/数字拆分被误读成 (12, 28)
    for bad in ("A1B2", "1A", "A", "123", "", "A1!"):
        with pytest.raises(InvalidArgumentError):
            _parse_cell(bad)


def test_parse_cell_case_insensitive():
    assert _parse_cell("a1") == (1, 1)
    assert _parse_cell("xfd1048576") == (1048576, 16384)


# --- #41：三个 range op 畸形地址先于触 COM 抛 InvalidArgumentError ---


def _raise_if_called(*args, **kwargs):
    raise AssertionError("畸形地址不应触达 COM 层 _ws")


@pytest.fixture
def excel_stub(monkeypatch):
    app = object.__new__(ExcelApp)
    monkeypatch.setattr(ExcelApp, "_ws", _raise_if_called)
    return app


@pytest.mark.parametrize(
    "op, kwargs",
    [
        ("set_number_format", {"range_addr": "???", "fmt": "#,##0"}),
        (
            "add_conditional_format",
            {"range_addr": "???:", "rule": "cell", "operator": "greater", "value": 1},
        ),
        ("autofit", {"range_addr": "1A1"}),
    ],
)
def test_malformed_range_rejected_before_com(excel_stub, op, kwargs):
    with pytest.raises(InvalidArgumentError, match="非法区域"):
        getattr(excel_stub, op)("sheet1", doc_id="d1", **kwargs)


@pytest.mark.parametrize(
    "op, kwargs",
    [
        ("set_number_format", {"range_addr": "A1:B2", "fmt": "#,##0"}),
        (
            "add_conditional_format",
            {"range_addr": "A1:B2", "rule": "cell", "operator": "greater", "value": 1},
        ),
        ("autofit", {"range_addr": "A1:B2"}),
    ],
)
def test_valid_range_reaches_com(excel_stub, op, kwargs):
    with pytest.raises(AssertionError, match="不应触达"):
        getattr(excel_stub, op)("sheet1", doc_id="d1", **kwargs)
