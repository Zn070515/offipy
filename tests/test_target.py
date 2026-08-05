"""目标身份原语 + 只读不创建（P0-7/P0-8）纯逻辑测试。

用 `__new__` 绕过 `__init__`（避免真连 COM），再 monkeypatch core.active_doc /
core.doc_alive 注入假文档：断言 active_* 无文档返回 None 且不调 Add()；
依赖文档的 op 无文档抛 TargetNotFoundError；get_target 返回身份 dict。
"""

import pytest

from offipy import excel, server, word
from offipy.exceptions import InvalidArgumentError, TargetNotFoundError


class _NoAdd:
    """记录 Add 是否被调用；真被调用即失败。"""

    def __init__(self):
        self.add_called = False

    def Add(self, *a, **k):
        self.add_called = True
        raise AssertionError("P0-8: 读操作不得隐式 Add()")


class _FakeApp:
    def __init__(self, active, no_add):
        self.ActiveWorkbook = active  # excel
        self.ActiveDocument = active  # word
        self.ActivePresentation = active  # ppt
        self.Workbooks = no_add
        self.Documents = no_add
        self.Presentations = no_add
        self.DisplayAlerts = True  # _alerts_scope 读写 DisplayAlerts


class _Book:
    def __init__(self, name, path):
        self.Name = name
        self.FullName = path


def _no_doc_env(monkeypatch):
    monkeypatch.setattr("offipy.core.active_doc", lambda *a: None)
    monkeypatch.setattr("offipy.core.doc_alive", lambda *a: False)


# --- active_* 只读不创建（P0-8） ---


def test_active_book_none_without_doc(monkeypatch):
    no_add = _NoAdd()
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, no_add)
    _no_doc_env(monkeypatch)
    assert app.active_book() is None
    assert no_add.add_called is False


def test_active_doc_none_without_doc(monkeypatch):
    no_add = _NoAdd()
    app = word.WordApp.__new__(word.WordApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, no_add)
    _no_doc_env(monkeypatch)
    assert app.active_doc() is None
    assert no_add.add_called is False


def test_active_pres_none_without_doc(monkeypatch):
    from offipy import ppt

    no_add = _NoAdd()
    app = ppt.PptApp.__new__(ppt.PptApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, no_add)
    _no_doc_env(monkeypatch)
    assert app.active_pres() is None
    assert no_add.add_called is False


def test_active_book_falls_back_to_real_active(monkeypatch):
    no_add = _NoAdd()
    book = _Book("Book1", r"C:\x\Book1.xlsx")
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(book, no_add)
    _no_doc_env(monkeypatch)
    assert app.active_book() is book  # 落到 self.app.ActiveWorkbook，仍不 Add
    assert no_add.add_called is False


# --- get_target 身份（P0-7） ---


def test_get_target_excel(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(_Book("Book1", r"C:\x\Book1.xlsx"), _NoAdd())
    _no_doc_env(monkeypatch)
    assert app.get_target() == {
        "app": "excel",
        "doc_id": "book1",  # 实时解析的 ActiveWorkbook 并入文档表后分配 doc_id
        "name": "Book1",
        "path": r"C:\x\Book1.xlsx",
    }


def test_get_target_explicit_doc_id(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {"b1": _Book("Book1", r"C:\x\Book1.xlsx")}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    # 显式 doc_id 路由要求句柄存活（与 _no_doc_env 的 doc_alive=False 相反）
    monkeypatch.setattr("offipy.core.active_doc", lambda *a: None)
    monkeypatch.setattr("offipy.core.doc_alive", lambda *a: True)
    assert app.get_target(doc_id="b1")["doc_id"] == "b1"
    assert app.get_target(doc_id="b1")["name"] == "Book1"


def test_get_target_unknown_doc_id_raises(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    with pytest.raises(TargetNotFoundError):
        app.get_target(doc_id="nope")


def test_get_target_none_without_doc(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    assert app.get_target() is None


# --- 依赖文档的 op 无文档 → TargetNotFoundError ---


def test_read_range_no_doc_raises(monkeypatch):
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    with pytest.raises(TargetNotFoundError):
        app.read_range(1, "A1")
    with pytest.raises(TargetNotFoundError):
        app.save()


def test_word_read_doc_text_no_doc_raises(monkeypatch):
    app = word.WordApp.__new__(word.WordApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    with pytest.raises(TargetNotFoundError):
        app.read_doc_text()


def test_ppt_read_slide_texts_no_doc_raises(monkeypatch):
    from offipy import ppt

    app = ppt.PptApp.__new__(ppt.PptApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    app.app = _FakeApp(None, _NoAdd())
    _no_doc_env(monkeypatch)
    with pytest.raises(TargetNotFoundError):
        app.read_slide_texts()


# --- expected_target 绑定（P0-7，server dispatch 层） ---


class _Visible:
    Visible = True


class _TargetApp:
    def __init__(self, target):
        self.app = _Visible()
        self._target = target
        self.close_calls = []

    def get_target(self, doc_id=None):
        return self._target

    def close_book(self, doc_id=None):
        self.close_calls.append(doc_id)
        return "closed"

    def read_range(self, sheet, range_addr):
        return [[1]]


def test_dispatch_expected_target_match():
    app = _TargetApp(
        {"app": "excel", "doc_id": "book1", "name": "Book1", "path": r"C:\x\Book1.xlsx"}
    )
    result = server.dispatch(app, "close_book", {"expected_target": {"name": "Book1"}}, "excel")
    assert result == "closed"
    assert app.close_calls == ["book1"]  # resolve-once：校验的 doc_id 注入方法调用


def test_dispatch_expected_target_mismatch():
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "Other", "path": None})
    with pytest.raises(TargetNotFoundError):
        server.dispatch(app, "close_book", {"expected_target": {"name": "Book1"}}, "excel")
    assert app.close_calls == []  # P0-5：校验失败不得执行（防「校验 A 执行 B」）


def test_dispatch_expected_target_no_target():
    app = _TargetApp(None)
    with pytest.raises(TargetNotFoundError):
        server.dispatch(app, "close_book", {"expected_target": {"name": "Book1"}}, "excel")
    assert app.close_calls == []


def test_dispatch_expected_target_empty_dict_rejected():
    # P0-4：旧 _target_matches({}) 恒真绕过——空对象必须拒绝
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "Book1", "path": None})
    with pytest.raises(InvalidArgumentError):
        server.dispatch(app, "close_book", {"expected_target": {}}, "excel")
    assert app.close_calls == []


def test_dispatch_expected_target_unknown_key_rejected():
    app = _TargetApp({"app": "excel", "doc_id": "book1", "name": "Book1", "path": None})
    with pytest.raises(InvalidArgumentError):
        server.dispatch(app, "close_book", {"expected_target": {"bogus": 1}}, "excel")
    assert app.close_calls == []


def test_dispatch_expected_target_doc_id_binding():
    # 显式 doc_id 绑定：校验用 doc_id 解析目标后直接注入方法参数
    app = _TargetApp({"app": "excel", "doc_id": "book2", "name": "Book2", "path": None})
    result = server.dispatch(app, "close_book", {"expected_target": {"doc_id": "book2"}}, "excel")
    assert result == "closed"
    assert app.close_calls == ["book2"]


def test_dispatch_expected_target_ignored_on_readonly():
    # 只读 op 不做绑定校验，expected_target 仅被消费（不传给方法）
    app = _TargetApp(None)
    result = server.dispatch(
        app,
        "read_range",
        {"sheet": 1, "range_addr": "A1", "expected_target": {"name": "x"}},
        "excel",
    )
    assert result == [[1]]
