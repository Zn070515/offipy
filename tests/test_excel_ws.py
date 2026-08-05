"""ExcelApp._ws 工作表解析（不碰 COM）。

只捕 pywintypes.com_error（测试里替换为 _FakeComError）：名称不存在且非数字
→ ComOperationError（不裸 ValueError）；纯数字 → 按序号解析。
"""

from types import SimpleNamespace

import pytest

from offipy import excel
from offipy.exceptions import ComOperationError


class _FakeComError(Exception):
    pass


class _FakeBook:
    def __init__(self):
        self._sheet = SimpleNamespace()

    def Worksheets(self, sheet):
        if isinstance(sheet, int):
            if sheet == 1:
                return self._sheet
            raise _FakeComError()
        if sheet == "数据":
            return self._sheet
        raise _FakeComError()


def _app(book):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._require_book = lambda doc_id=None: book
    return app


def _ws(app, sheet):
    return app._ws(sheet)


def test_ws_by_name_hit(monkeypatch):
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    book = _FakeBook()
    assert _ws(_app(book), "数据") is book._sheet


def test_ws_by_index_int(monkeypatch):
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    book = _FakeBook()
    assert _ws(_app(book), 1) is book._sheet


def test_ws_missing_name_raises_com_error(monkeypatch):
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    with pytest.raises(ComOperationError, match="工作表不存在"):
        _ws(_app(_FakeBook()), "不存在的表")


def test_ws_numeric_string_falls_back_to_index(monkeypatch):
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    book = _FakeBook()
    assert _ws(_app(book), "1") is book._sheet


def test_ws_numeric_string_out_of_range(monkeypatch):
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    with pytest.raises(ComOperationError, match="工作表不存在"):
        _ws(_app(_FakeBook()), "2")


def test_ws_weird_digit_no_bare_valueerror(monkeypatch):
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    # "①" isdigit() 为 True 但 int() 抛 ValueError：不得裸 ValueError 外泄
    with pytest.raises(ComOperationError, match="工作表不存在"):
        _ws(_app(_FakeBook()), "①")
