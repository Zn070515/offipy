"""P2-2 多文档：doc_id 显式路由 / activate / list_docs / 缺省活动跟随。

用 `__new__` 绕过 __init__，注入 fake 应用与文档表，纯逻辑验证：
- new_book 连续创建后活动目标跟随最新；get_target 指向活动工作簿
- activate(doc_id) 切换后续缺省操作的目标
- 内容 op 显式 doc_id 路由到指定工作簿，不受活动目标影响
- list_docs 列出全部已登记文档；close_book(doc_id) 从表移除
- 未知 doc_id → TargetNotFoundError
- word/ppt 套同模式
"""

import os
from pathlib import Path

import pytest

from offipy import core, excel, ppt, word
from offipy.exceptions import FileConflictError, TargetNotFoundError


class _FakeCell:
    def __init__(self, book, row, col):
        self._book = book
        self._key = (row, col)

    @property
    def Value(self):
        return self._book._cells.get(self._key)

    @Value.setter
    def Value(self, v):
        self._book._cells[self._key] = v


class _FakeBook:
    def __init__(self, name):
        self.Name = name
        self.FullName = rf"C:\x\{name}.xlsx"
        self.Path = ""  # 空 = 从未保存（save()/close_book() 依赖）
        self.Saved = True
        self._cells = {}
        self.closed = False
        self.activated = 0
        self.save_calls = 0
        self.saveas_calls = []
        self.close_calls = []

    def Activate(self):
        self.activated += 1  # 真实 UI 同步桩：activate() 必须触达

    def Worksheets(self, *a):
        return self  # 工作表对象：Cells 代理把写入记到 self._cells

    def Cells(self, row, col):
        return _FakeCell(self, row, col)

    def Save(self):
        self.save_calls += 1
        self.Saved = True

    def SaveAs(self, path):
        self.saveas_calls.append(path)
        self.Path = os.path.dirname(path)
        self.FullName = path
        self.Saved = True

    def Close(self, **kw):
        self.close_calls.append(kw)
        self.closed = True


class _FakeExcelApp:
    def __init__(self):
        self.DisplayAlerts = True
        self._created = []
        self.active = None

    @property
    def Workbooks(self):
        return self

    def Add(self):
        book = _FakeBook(f"Book{len(self._created) + 1}")
        self._created.append(book)
        self.active = book
        return book

    @property
    def ActiveWorkbook(self):
        return self.active


def _no_probe(monkeypatch):
    monkeypatch.setattr(core, "active_doc", lambda *a: None)
    monkeypatch.setattr(core, "doc_alive", lambda *a: True)


def _new_excel(monkeypatch):
    _no_probe(monkeypatch)
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    app.app = _FakeExcelApp()
    app._docs = {}
    app._active_id = None
    app._seq = 0
    return app


def _new_word(monkeypatch):
    _no_probe(monkeypatch)
    app = word.WordApp.__new__(word.WordApp)
    app.app = _FakeWordApp()
    app._docs = {}
    app._active_id = None
    app._seq = 0
    return app


def _new_ppt(monkeypatch):
    _no_probe(monkeypatch)
    app = ppt.PptApp.__new__(ppt.PptApp)
    app.app = _FakePptApp()
    app._docs = {}
    app._active_id = None
    app._seq = 0
    return app


def test_excel_multi_doc_flow(monkeypatch):
    app = _new_excel(monkeypatch)
    a = app.new_book()  # book1
    b = app.new_book()  # book2
    assert a == "book1"
    assert b == "book2"
    # 活动目标跟随最新创建；get_target 指向活动工作簿
    assert app.get_target()["name"] == "Book2"
    # activate 切换后续缺省操作的目标 → 写 A
    app.activate(a)
    assert app._docs[a].activated == 1  # P0-6：activate 同步真实 UI（Workbook.Activate）
    app.set_cell(1, "A1", 100)
    assert app._docs[a]._cells[(1, 1)] == 100
    assert (1, 1) not in app._docs[b]._cells
    # 显式 doc_id 路由到指定工作簿，不受活动目标影响
    app.set_cell(1, "B2", 200, doc_id=b)
    assert app._docs[b]._cells[(2, 2)] == 200
    assert (2, 2) not in app._docs[a]._cells
    # get_cell 读回，缺省走活动、显式路由指定
    assert app.get_cell(1, "A1") == 100
    assert app.get_cell(1, "B2", doc_id=b) == 200
    # list_docs 列出全部登记文档
    docs = app.list_docs()
    assert set(docs) == {"book1", "book2"}
    assert docs["book1"]["name"] == "Book1"


def test_excel_close_doc_removes_from_table(monkeypatch):
    app = _new_excel(monkeypatch)
    a = app.new_book()
    b = app.new_book()
    book_b = app._docs[b]  # 关闭前抓住句柄
    app.activate(a)
    app.close_book(save=False, doc_id=b)
    assert book_b.closed
    assert b not in app._docs
    assert a in app._docs
    # 关闭非活动文档不影响活动目标
    assert app.get_target()["name"] == "Book1"


def test_excel_close_active_clears_until_new(monkeypatch):
    app = _new_excel(monkeypatch)
    app.new_book()
    app.close_book(save=False)  # 关闭活动工作簿
    app.app.active = None  # Excel 侧也无活动工作簿
    assert app.active_book() is None
    with pytest.raises(TargetNotFoundError):
        app.read_range(1, "A1")


def test_excel_unknown_doc_id_raises(monkeypatch):
    app = _new_excel(monkeypatch)
    app.new_book()
    with pytest.raises(TargetNotFoundError):
        app.read_range(1, "A1", doc_id="nope")
    with pytest.raises(TargetNotFoundError):
        app.activate("nope")
    with pytest.raises(TargetNotFoundError):
        app.save(doc_id="nope")


# --- close/save 防弹窗：save=False 不触发另存为；save=True 从未保存自动落盘 ---


def test_excel_close_discard_sets_saved_true_no_dialog(monkeypatch):
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    assert app.close_book(save=False) is None
    assert book.Saved is True  # 先标 Saved=True 才不弹另存为（Excel unsaved 特例）
    assert book.closed
    assert book.close_calls[-1]["SaveChanges"] == 2  # xlDoNotSaveChanges
    assert book.save_calls == 0
    assert book.saveas_calls == []  # 不 SaveAs、不弹对话框


def test_excel_close_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]  # Path="" → 从未保存
    path = app.close_book(save=True)
    assert book.saveas_calls  # 自动落盘，不弹另存为
    assert path == book.saveas_calls[-1]
    assert os.path.dirname(path) == str(tmp_path)  # 同层目录 = cwd
    assert path.endswith(".xlsx")
    assert book.close_calls[-1]["SaveChanges"] == 1  # xlSaveChanges
    assert book.closed


def test_excel_close_save_saved_returns_fullname(monkeypatch):
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    book.Path = "C:/x"  # 已有保存路径 → 原位，不再额外 Save
    path = app.close_book(save=True)
    assert path == book.FullName
    assert book.close_calls[-1]["SaveChanges"] == 1
    assert book.save_calls == 0
    assert book.saveas_calls == []


def test_excel_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    path = app.save()
    assert book.saveas_calls
    assert path == book.saveas_calls[-1]
    assert os.path.dirname(path) == str(tmp_path)
    assert path.endswith(".xlsx")


def test_excel_save_saved_uses_in_place(monkeypatch):
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    book.Path = "C:/x"
    path = app.save()
    assert book.save_calls == 1
    assert path == book.FullName
    assert book.saveas_calls == []


def test_excel_save_explicit_path_overwrite_protection(monkeypatch, tmp_path):
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    dest = os.path.join(str(tmp_path), "out.xlsx")
    assert app.save(dest) == dest
    assert book.saveas_calls == [dest]
    Path(dest).write_text("x")  # 真实落盘，ensure_writable 才判定已存在
    with pytest.raises(FileConflictError):
        app.save(dest)
    assert book.saveas_calls == [dest]  # 未覆盖 → 不重复 SaveAs
    assert app.save(dest, overwrite=True) == dest
    assert book.saveas_calls == [dest, dest]


# --- word / ppt 套同模式 ---


class _FakeContent:
    def __init__(self):
        self.text = ""

    def InsertAfter(self, t):
        self.text += t

    @property
    def Text(self):
        return self.text


class _FakeWordDoc:
    def __init__(self, name):
        self.Name = name
        self.FullName = rf"C:\x\{name}.docx"
        self.Path = ""
        self.Saved = True
        self.Content = _FakeContent()
        self.activated = 0
        self.save_calls = 0
        self.saveas_calls = []
        self.close_calls = []

    def Activate(self):
        self.activated += 1

    def Save(self):
        self.save_calls += 1
        self.Saved = True

    def SaveAs2(self, path):
        self.saveas_calls.append(path)
        self.Path = os.path.dirname(path)
        self.FullName = path
        self.Saved = True

    def Close(self, **kw):
        self.close_calls.append(kw)


class _FakeWordApp:
    def __init__(self):
        self.DisplayAlerts = 0
        self._created = []
        self.active = None

    @property
    def Documents(self):
        return self

    def Add(self):
        d = _FakeWordDoc(f"Doc{len(self._created) + 1}")
        self._created.append(d)
        self.active = d
        return d

    @property
    def ActiveDocument(self):
        return self.active


def test_word_multi_doc_route(monkeypatch):
    _no_probe(monkeypatch)
    app = word.WordApp.__new__(word.WordApp)
    app.app = _FakeWordApp()
    app._docs = {}
    app._active_id = None
    app._seq = 0
    d1 = app.new_doc()  # doc1
    d2 = app.new_doc()  # doc2
    app.activate(d1)
    app.write_line("hello")
    assert app._docs[d1].Content.text == "hello\r\n"
    assert app._docs[d2].Content.text == ""
    app.write_line("world", doc_id=d2)
    assert app._docs[d2].Content.text == "world\r\n"
    assert app.read_doc_text(doc_id=d2) == "world\r\n"
    assert app.get_target()["name"] == "Doc1"


def test_word_close_discard_sets_saved_true_no_dialog(monkeypatch):
    app = _new_word(monkeypatch)
    doc = app._docs[app.new_doc()]
    assert app.close_doc(save=False) is None
    assert doc.Saved is True
    assert doc.close_calls[-1]["SaveChanges"] == 0  # wdDoNotSaveChanges
    assert doc.save_calls == 0
    assert doc.saveas_calls == []


def test_word_close_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    app = _new_word(monkeypatch)
    doc = app._docs[app.new_doc()]
    path = app.close_doc(save=True)
    assert doc.saveas_calls
    assert path == doc.saveas_calls[-1]
    assert os.path.dirname(path) == str(tmp_path)
    assert path.endswith(".docx")
    assert doc.close_calls[-1]["SaveChanges"] == -1  # wdSaveChanges


def test_word_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    app = _new_word(monkeypatch)
    doc = app._docs[app.new_doc()]
    path = app.save()
    assert doc.saveas_calls
    assert path == doc.saveas_calls[-1]
    assert os.path.dirname(path) == str(tmp_path)
    assert path.endswith(".docx")


class _FakePres:
    def __init__(self, name):
        self.Name = name
        self.FullName = rf"C:\x\{name}.pptx"
        self.Path = ""
        self.Saved = True
        self.activated = 0
        self.save_calls = 0
        self.saveas_calls = []

    def Activate(self):
        self.activated += 1

    def Save(self):
        self.save_calls += 1
        self.Saved = True

    def SaveAs(self, path):
        self.saveas_calls.append(path)
        self.Path = os.path.dirname(path)
        self.FullName = path
        self.Saved = True


class _FakePptApp:
    def __init__(self):
        self.DisplayAlerts = 1
        self._created = []
        self.active = None

    @property
    def Presentations(self):
        return self

    def Add(self):
        p = _FakePres(f"Pres{len(self._created) + 1}")
        self._created.append(p)
        self.active = p
        return p

    @property
    def ActivePresentation(self):
        return self.active


def test_ppt_multi_doc_ids(monkeypatch):
    _no_probe(monkeypatch)
    app = ppt.PptApp.__new__(ppt.PptApp)
    app.app = _FakePptApp()
    app._docs = {}
    app._active_id = None
    app._seq = 0
    p1 = app.new_pres()  # pres1
    p2 = app.new_pres()  # pres2
    assert p1 == "pres1"
    assert p2 == "pres2"
    assert app.get_target()["name"] == "Pres2"
    app.activate(p1)
    assert app.get_target()["name"] == "Pres1"
    assert set(app.list_docs()) == {"pres1", "pres2"}


def test_ppt_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    app = _new_ppt(monkeypatch)
    pres = app._docs[app.new_pres()]
    path = app.save()
    assert pres.saveas_calls
    assert path == pres.saveas_calls[-1]
    assert os.path.dirname(path) == str(tmp_path)
    assert path.endswith(".pptx")


def test_ppt_save_saved_uses_in_place(monkeypatch):
    app = _new_ppt(monkeypatch)
    pres = app._docs[app.new_pres()]
    pres.Path = "C:/x"
    path = app.save()
    assert pres.save_calls == 1
    assert path == pres.FullName
    assert pres.saveas_calls == []
