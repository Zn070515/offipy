"""C5/C6: freeze_panes 整数校验 + add_sheet 幂等与名称校验（不碰 COM）。

用 `__new__` 构造实例跳过 __init__ 的 COM 初始化，专注参数校验与幂等逻辑。
"""

from types import SimpleNamespace

import pytest

from offipy import excel
from offipy.exceptions import InvalidArgumentError

# ---------------------------------------------------------------------------
# freeze_panes
# ---------------------------------------------------------------------------


def test_freeze_panes_rejects_non_int_rows_cols():
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    for bad in ("2", 2.5, True):
        with pytest.raises(InvalidArgumentError, match="整数"):
            app.freeze_panes("S", rows=bad, cols=0, doc_id="b1")
        with pytest.raises(InvalidArgumentError, match="整数"):
            app.freeze_panes("S", rows=0, cols=bad, doc_id="b1")


def test_freeze_panes_rejects_negative():
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    with pytest.raises(InvalidArgumentError, match="≥0"):
        app.freeze_panes("S", rows=-1, cols=0, doc_id="b1")
    with pytest.raises(InvalidArgumentError, match="≥0"):
        app.freeze_panes("S", rows=0, cols=-1, doc_id="b1")


def test_freeze_panes_valid_proceeds():
    # 合法 rows/cols 一路传到 Activate + FreezePanes（验证冻结 cell 坐标 (rows+1, cols+1)）
    calls = {}

    class _Cells:
        def __init__(self, row, col):
            calls["select"] = (row, col)

        def Select(self):
            pass

    class _Ws:
        def Activate(self):
            calls["activated"] = True

        def Cells(self, row, col):
            return _Cells(row, col)

    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._require_book = lambda doc_id=None: SimpleNamespace()
    app._ws = lambda sheet, doc_id=None: _Ws()
    app.app = SimpleNamespace(ActiveWindow=SimpleNamespace(FreezePanes=None))
    app.freeze_panes("S", rows=2, cols=3, doc_id="b1")
    assert app.app.ActiveWindow.FreezePanes is True
    assert calls["select"] == (3, 4)
    assert calls["activated"] is True


def test_freeze_panes_zero_clears():
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._require_book = lambda doc_id=None: SimpleNamespace()
    app._ws = lambda sheet, doc_id=None: SimpleNamespace(
        Activate=lambda: None,
        Cells=lambda r, c: SimpleNamespace(Select=lambda: None),
    )
    app.app = SimpleNamespace(ActiveWindow=SimpleNamespace(FreezePanes=True))
    app.freeze_panes("S", rows=0, cols=0, doc_id="b1")
    assert app.app.ActiveWindow.FreezePanes is False


# ---------------------------------------------------------------------------
# add_sheet
# ---------------------------------------------------------------------------


class _FakeComError(Exception):
    pass


class _Ws:
    def __init__(self, name):
        self.Name = name


class _Worksheets:
    """形如 COM Worksheets 集合：Worksheets(name) 索引、Worksheets.Add() 追加。"""

    def __init__(self, sheets):
        self._sheets = sheets

    def __call__(self, name):
        for ws in self._sheets:
            if ws.Name == name:
                return ws
        raise _FakeComError()

    def Add(self):
        ws = _Ws("Sheet1")
        self._sheets.insert(0, ws)
        return ws


class _AddBook:
    """已含 Sheet1；Worksheets(name) 命中返回，未命中抛错；Add 追加到列表。"""

    def __init__(self):
        self.sheets = [_Ws("Sheet1")]
        self.Worksheets = _Worksheets(self.sheets)


def test_add_sheet_idempotent_when_exists(monkeypatch):
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    book = _AddBook()
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._require_book = lambda doc_id=None: book
    result = app.add_sheet("Sheet1", doc_id="b1")
    assert result is book.sheets[0]
    assert len(book.sheets) == 1  # 未新增重复表


def test_add_sheet_creates_new(monkeypatch):
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    book = _AddBook()
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._require_book = lambda doc_id=None: book
    ws = app.add_sheet("数据", doc_id="b1")
    assert ws.Name == "数据"
    assert len(book.sheets) == 2


@pytest.mark.parametrize(
    "bad",
    ["", "a" * 32, "a/b", "a\\b", "a?b", "a*b", "a[b", "a]b", "a:b"],
)
def test_add_sheet_rejects_invalid_names(monkeypatch, bad):
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._require_book = lambda doc_id=None: _AddBook()
    with pytest.raises(InvalidArgumentError):
        app.add_sheet(bad, doc_id="b1")


def test_add_sheet_rejects_control_char(monkeypatch):
    # chr(0) 运行时构造 NUL，避免在 .py 源码里嵌入转义
    monkeypatch.setattr(excel, "_COM_ERROR", _FakeComError)
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._require_book = lambda doc_id=None: _AddBook()
    with pytest.raises(InvalidArgumentError):
        app.add_sheet("a" + chr(0) + "b", doc_id="b1")
