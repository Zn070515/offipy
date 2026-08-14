"""P2-2 多文档：doc_id 显式路由 / activate / list_docs / 缺省活动跟随。

用 `__new__` 绕过 __init__，注入 fake 应用与文档表，纯逻辑验证：
- new_book 连续创建后活动目标跟随最新；get_target 指向活动工作簿
- activate(doc_id) 切换后续缺省操作的目标
- 内容 op 显式 doc_id 路由到指定工作簿，不受活动目标影响
- list_docs 列出全部已登记文档；close_book(doc_id) 从表移除
- 未知 doc_id → TargetNotFoundError
- word/ppt 套同模式
"""

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
        self.app.active = self  # P0-3：激活即成为真实 ActiveWorkbook

    def Cells(self, row, col):
        return _FakeCell(self, row, col)

    def Worksheets(self, sheet):
        return self  # 单工作表假体：sheet 名/序号忽略，单元格落同一 _cells

    def Save(self):
        self.save_calls += 1
        self.Saved = True

    def SaveAs(self, path):
        self.saveas_calls.append(path)
        self.Path = Path(path).parent
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
        book.app = self
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
    return app


def _new_word(monkeypatch):
    _no_probe(monkeypatch)
    app = word.WordApp.__new__(word.WordApp)
    app.app = _FakeWordApp()
    app._docs = {}
    app._active_id = None
    return app


def _new_ppt(monkeypatch):
    _no_probe(monkeypatch)
    app = ppt.PptApp.__new__(ppt.PptApp)
    app.app = _FakePptApp()
    app._docs = {}
    app._active_id = None
    return app


def test_excel_multi_doc_flow(monkeypatch):
    app = _new_excel(monkeypatch)
    a = app.new_book()
    b = app.new_book()
    assert a.startswith("book") and b.startswith("book")
    assert a != b  # H10：doc_id 高熵，不再顺序可猜
    # 活动目标跟随最新创建；get_target 指向活动工作簿
    assert app.get_target()["name"] == "Book2"
    # activate 切换后续缺省操作的目标 → 写 A
    app.activate(a)
    assert app._docs[a].activated == 1  # P0-6：activate 同步真实 UI（Workbook.Activate）
    app.set_cell(1, "A1", 100, follow_active=True)
    assert app._docs[a]._cells[1, 1] == 100
    assert (1, 1) not in app._docs[b]._cells
    # 显式 doc_id 路由到指定工作簿，不受活动目标影响
    app.set_cell(1, "B2", 200, doc_id=b)
    assert app._docs[b]._cells[2, 2] == 200
    assert (2, 2) not in app._docs[a]._cells
    # get_cell 读回，缺省走活动、显式路由指定
    assert app.get_cell(1, "A1") == 100
    assert app.get_cell(1, "B2", doc_id=b) == 200
    # list_docs 列出全部登记文档
    docs = app.list_docs()
    assert set(docs) == {a, b}
    assert docs[a]["name"] == "Book1"


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
    app.close_book(save=False, follow_active=True)  # 关闭活动工作簿
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


def test_unknown_doc_id_message_notes_session_boundary(monkeypatch):
    # 0.10.2：未知 doc_id 报错说明「当前会话」+ 跨会话 doc_id 不互通——此前
    # 写「用 list_docs 查看当前打开的」，但本地直连 Ppt()/Excel()/Word() 的
    # doc_id 与会话式 Remote*/CLI/HTTP 互不相通，list_docs 显示在也不代表
    # 本会话能查到，误导排查方向。
    app = _new_excel(monkeypatch)
    app.new_book()
    with pytest.raises(TargetNotFoundError, match="会话"):
        app.read_range(1, "A1", doc_id="nope")


# --- doc_id 稳定身份（P0-4）：同底层文档重开/重连复用同 doc_id ---


def test_excel_register_reuses_same_fullname(monkeypatch):
    app = _new_excel(monkeypatch)
    b1 = _FakeBook("report")
    b1.Path = "C:/x"  # 已保存 → 稳定身份 = FullName.lower()
    did = app._register(b1)
    assert did.startswith("book")  # H10：doc_id 高熵，带 book 前缀
    b2 = _FakeBook("report")  # 不同 wrapper，同 FullName
    b2.Path = "C:/x"
    assert app._register(b2) == did  # 复用同一 doc_id
    assert app._docs[did] is b2  # 句柄换成实时对象
    assert set(app._docs) == {did}


def test_excel_register_reuses_same_unsaved_name(monkeypatch):
    app = _new_excel(monkeypatch)
    a = app._register(_FakeBook("Book7"))  # 未保存 → 身份 = Name.lower()
    assert a.startswith("book")
    assert app._register(_FakeBook("Book7")) == a
    assert set(app._docs) == {a}


def test_word_register_reuses_same_fullname(monkeypatch):
    app = _new_word(monkeypatch)
    d1 = _FakeWordDoc("report")
    d1.Path = "C:/x"
    did = app._register(d1)
    assert did.startswith("doc")
    d2 = _FakeWordDoc("report")
    d2.Path = "C:/x"
    assert app._register(d2) == did
    assert set(app._docs) == {did}


def test_ppt_register_reuses_same_fullname(monkeypatch):
    app = _new_ppt(monkeypatch)
    p1 = _FakePres("report")
    p1.Path = "C:/x"
    did = app._register(p1)
    assert did.startswith("pres")
    p2 = _FakePres("report")
    p2.Path = "C:/x"
    assert app._register(p2) == did
    assert set(app._docs) == {did}


# --- close/save 防弹窗：save=False 不触发另存为；save=True 从未保存自动落盘 ---


def test_excel_close_discard_sets_saved_true_no_dialog(monkeypatch):
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    assert app.close_book(save=False, follow_active=True) is None
    assert book.Saved is True  # 先标 Saved=True 才不弹另存为（Excel unsaved 特例）
    assert book.closed
    assert book.close_calls[-1]["SaveChanges"] == 2  # xlDoNotSaveChanges
    assert book.save_calls == 0
    assert book.saveas_calls == []  # 不 SaveAs、不弹对话框


def test_excel_close_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]  # Path="" → 从未保存
    path = app.close_book(save=True, follow_active=True)
    assert book.saveas_calls  # 自动落盘，不弹另存为
    assert path == book.saveas_calls[-1]
    assert Path(path).parent == tmp_path / "documents"  # 用户数据目录
    assert path.endswith(".xlsx")
    assert book.close_calls[-1]["SaveChanges"] == 1  # xlSaveChanges
    assert book.closed


def test_excel_close_save_saved_returns_fullname(monkeypatch):
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    book.Path = "C:/x"  # 已有保存路径 → 原位，不再额外 Save
    path = app.close_book(save=True, follow_active=True)
    assert path == book.FullName
    assert book.close_calls[-1]["SaveChanges"] == 1
    assert book.save_calls == 0
    assert book.saveas_calls == []


def test_excel_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    path = app.save(follow_active=True)
    assert book.saveas_calls
    assert path == book.saveas_calls[-1]
    assert Path(path).parent == tmp_path / "documents"
    assert path.endswith(".xlsx")


def test_excel_save_saved_uses_in_place(monkeypatch):
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    book.Path = "C:/x"
    path = app.save(follow_active=True)
    assert book.save_calls == 1
    assert path == book.FullName
    assert book.saveas_calls == []


def test_excel_save_explicit_path_overwrite_protection(monkeypatch, tmp_path):
    app = _new_excel(monkeypatch)
    book = app._docs[app.new_book()]
    dest = str(tmp_path / "out.xlsx")
    assert app.save(dest, follow_active=True) == dest
    assert book.saveas_calls == [dest]
    Path(dest).write_text("x")  # 真实落盘，ensure_writable 才判定已存在
    with pytest.raises(FileConflictError):
        app.save(dest, follow_active=True)
    assert book.saveas_calls == [dest]  # 未覆盖 → 不重复 SaveAs
    assert app.save(dest, overwrite=True, follow_active=True) == dest
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
        self.app.active = self  # P0-3：激活即成为真实 ActiveDocument

    def Save(self):
        self.save_calls += 1
        self.Saved = True

    def SaveAs2(self, path):
        self.saveas_calls.append(path)
        self.Path = Path(path).parent
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
        d.app = self
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
    d1 = app.new_doc()  # doc1
    d2 = app.new_doc()  # doc2
    app.activate(d1)
    app.write_line("hello", follow_active=True)
    assert app._docs[d1].Content.text == "hello\r\n"
    assert app._docs[d2].Content.text == ""
    app.write_line("world", doc_id=d2)
    assert app._docs[d2].Content.text == "world\r\n"
    assert app.read_doc_text(doc_id=d2) == "world\n"  # #16：段落符归一化成 \n
    assert app.get_target()["name"] == "Doc1"


def test_word_close_discard_sets_saved_true_no_dialog(monkeypatch):
    app = _new_word(monkeypatch)
    doc = app._docs[app.new_doc()]
    assert app.close_doc(save=False, follow_active=True) is None
    assert doc.Saved is True
    assert doc.close_calls[-1]["SaveChanges"] == 0  # wdDoNotSaveChanges
    assert doc.save_calls == 0
    assert doc.saveas_calls == []


def test_word_close_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    app = _new_word(monkeypatch)
    doc = app._docs[app.new_doc()]
    path = app.close_doc(save=True, follow_active=True)
    assert doc.saveas_calls
    assert path == doc.saveas_calls[-1]
    assert Path(path).parent == tmp_path / "documents"
    assert path.endswith(".docx")
    assert doc.close_calls[-1]["SaveChanges"] == -1  # wdSaveChanges


def test_word_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    app = _new_word(monkeypatch)
    doc = app._docs[app.new_doc()]
    path = app.save(follow_active=True)
    assert doc.saveas_calls
    assert path == doc.saveas_calls[-1]
    assert Path(path).parent == tmp_path / "documents"
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
        self.app.active = self  # P0-3：激活即成为真实 ActivePresentation

    def Save(self):
        self.save_calls += 1
        self.Saved = True

    def SaveAs(self, path):
        self.saveas_calls.append(path)
        self.Path = Path(path).parent
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
        p.app = self
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
    p1 = app.new_pres()
    p2 = app.new_pres()
    assert p1.startswith("pres") and p2.startswith("pres")
    assert p1 != p2  # H10：doc_id 高熵
    assert app.get_target()["name"] == "Pres2"
    app.activate(p1)
    assert app.get_target()["name"] == "Pres1"
    assert set(app.list_docs()) == {p1, p2}


def test_ppt_save_unsaved_autosaves(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr("offipy.paths.user_data_dir", lambda: tmp_path)
    app = _new_ppt(monkeypatch)
    pres = app._docs[app.new_pres()]
    path = app.save(follow_active=True)
    assert pres.saveas_calls
    assert path == pres.saveas_calls[-1]
    assert Path(path).parent == tmp_path / "documents"
    assert path.endswith(".pptx")


def test_ppt_save_saved_uses_in_place(monkeypatch):
    app = _new_ppt(monkeypatch)
    pres = app._docs[app.new_pres()]
    pres.Path = "C:/x"
    path = app.save(follow_active=True)
    assert pres.save_calls == 1
    assert path == pres.FullName
    assert pres.saveas_calls == []


# --- P1-5：list_docs 并入真实活动焦点并刷新 _active_id，不报陈旧焦点 ---


def test_excel_list_docs_refreshes_real_active(monkeypatch):
    app = _new_excel(monkeypatch)
    stale = app.new_book()  # 陈旧 active
    external = _FakeBook("Book9")  # 用户在真实 Excel 切到的活动工作簿
    external.app = app.app
    monkeypatch.setattr(core, "active_doc", lambda app_name, attr: external)
    docs = app.list_docs()
    assert app._active_id != stale  # 焦点刷新到真实活动工作簿
    assert docs[stale]["active"] is False
    assert docs[app._active_id]["active"] is True
    assert docs[app._active_id]["name"] == "Book9"


def test_word_list_docs_refreshes_real_active(monkeypatch):
    app = _new_word(monkeypatch)
    stale = app.new_doc()  # 陈旧 active
    external = _FakeWordDoc("Doc9")
    external.app = app.app
    monkeypatch.setattr(core, "active_doc", lambda app_name, attr: external)
    docs = app.list_docs()
    assert app._active_id != stale
    assert docs[stale]["active"] is False
    assert docs[app._active_id]["active"] is True
    assert docs[app._active_id]["name"] == "Doc9"


def test_ppt_list_docs_refreshes_real_active(monkeypatch):
    app = _new_ppt(monkeypatch)
    stale = app.new_pres()  # 陈旧 active
    external = _FakePres("Pres9")
    external.app = app.app
    monkeypatch.setattr(core, "active_doc", lambda app_name, attr: external)
    docs = app.list_docs()
    assert app._active_id != stale
    assert docs[stale]["active"] is False
    assert docs[app._active_id]["active"] is True
    assert docs[app._active_id]["name"] == "Pres9"
