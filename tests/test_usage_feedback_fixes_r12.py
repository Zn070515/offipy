"""Round-12 真实使用项目 Issue 反馈的修复测试（#21-#32，无 COM 依赖，纯逻辑/桩替身）。

分组：
- A. Word 错误路径前置校验（#23/#28/#30）
- B. Excel 错误路径前置校验（#24/#31）
- C. Ppt 校验 + close_pres（#27/#32/#26）
- D. api.op() 只读 follow_active（#25）
- E. CLI 退出码统一（#29）
- F. deck naturalDisplay + open_live 锁（#21/#22）
"""

import pytest

from offipy._comguard import _COM_ERROR
from offipy.exceptions import ComOperationError, InvalidArgumentError, TargetNotFoundError

# ================================================================ A. word


def test_add_list_rejects_unknown_style():
    from offipy import word

    app = word.WordApp.__new__(word.WordApp)
    with pytest.raises(InvalidArgumentError, match="未知列表样式"):
        app.add_list(["甲"], style="checkbox", doc_id="d1")
    with pytest.raises(InvalidArgumentError):
        app.add_list(["甲"], style="star", doc_id="d1")


def test_add_list_accepts_numbered(monkeypatch):
    from offipy import word

    applied = []

    class _ListFormat:
        def ApplyBulletDefault(self):
            applied.append("bullet")

        def ApplyNumberDefault(self):
            applied.append("numbered")

    class _Rng:
        def __init__(self):
            self.ListFormat = _ListFormat()

    class _Para:
        def __init__(self, start):
            self.Range = type("R", (), {"Start": start, "End": start + 1})()

    class _Paras:
        Count = 3

        def __call__(self, idx):
            return _Para(idx)

    class _Doc:
        def __init__(self):
            self.Paragraphs = _Paras()
            self._rng = _Rng()

        def Range(self, start, end):
            return self._rng

    doc = _Doc()
    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: doc)
    monkeypatch.setattr(app, "write_line", lambda text, doc_id=None: None)
    app.add_list(["甲", "乙"], style="numbered", doc_id="d1")
    assert applied == ["numbered"]


def test_add_list_bullet_and_bulleted_alias(monkeypatch):
    from offipy import word

    applied = []

    class _ListFormat:
        def ApplyBulletDefault(self):
            applied.append("bullet")

        def ApplyNumberDefault(self):
            applied.append("numbered")

    class _Rng:
        def __init__(self):
            self.ListFormat = _ListFormat()

    class _Para:
        def __init__(self, start):
            self.Range = type("R", (), {"Start": start, "End": start + 1})()

    class _Paras:
        Count = 0

        def __call__(self, idx):
            return _Para(idx)

    class _Doc:
        def __init__(self):
            self.Paragraphs = _Paras()
            self._rng = _Rng()

        def Range(self, start, end):
            return self._rng

    doc = _Doc()
    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: doc)
    monkeypatch.setattr(app, "write_line", lambda text, doc_id=None: None)
    app.add_list(["甲"], style="bullet", doc_id="d1")
    app.add_list(["乙"], style="bulleted", doc_id="d1")
    assert applied == ["bullet", "bullet"]


def test_format_text_paragraph_out_of_range_rejected(monkeypatch):
    from offipy import word

    class _Paras:
        Count = 2  # 文档只有 2 段

    class _Doc:
        Paragraphs = _Paras()

    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: _Doc())
    with pytest.raises(InvalidArgumentError, match="段落在 1..2 范围外"):
        app.format_text(paragraph=99, bold=True, doc_id="d1")


def test_format_paragraph_out_of_range_rejected(monkeypatch):
    from offipy import word

    class _Paras:
        Count = 2

    class _Doc:
        Paragraphs = _Paras()

    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: _Doc())
    with pytest.raises(InvalidArgumentError, match="段落在 1..2 范围外"):
        app.format_paragraph(paragraph=0, alignment="center", doc_id="d1")


def test_format_text_valid_paragraph_reaches_com(monkeypatch):
    # 回归：合法段落索引不误伤，Paragraphs(paragraph) 仍正常下发
    from offipy import word

    calls = []

    class _Font:
        pass

    class _Para:
        def __init__(self):
            self.Range = type("R", (), {"Font": _Font()})()

    class _Paras:
        Count = 1

        def __call__(self, idx):
            calls.append(idx)
            return _Para()

    class _Doc:
        Paragraphs = _Paras()

    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: _Doc())
    app.format_text(paragraph=1, bold=True, doc_id="d1")
    assert calls == [1]


def test_open_doc_missing_file_rejected(tmp_path):
    from offipy import word

    app = word.WordApp.__new__(word.WordApp)
    with pytest.raises(InvalidArgumentError, match="源文件不存在"):
        app.open_doc(str(tmp_path / "missing.docx"))


def test_insert_image_missing_file_rejected(tmp_path):
    from offipy import word

    app = word.WordApp.__new__(word.WordApp)
    with pytest.raises(InvalidArgumentError, match="源文件不存在"):
        app.insert_image(str(tmp_path / "missing.png"), doc_id="d1")


def test_set_table_col_width_rejects_width_below_7pt():
    from offipy import word

    app = word.WordApp.__new__(word.WordApp)
    with pytest.raises(InvalidArgumentError, match="列宽需 ≥ 7pt"):
        app.set_table_col_width(1, 1, 2.0, doc_id="d1")


def test_set_table_col_width_disables_autofit_before_set(monkeypatch):
    # #23 场景 A：autofit 后 AllowAutoFit 仍约束列宽赋值 → 先复位再设宽
    from offipy import word

    class _Col:
        def __init__(self):
            self.Width = None

    class _Table:
        def __init__(self):
            self.AllowAutoFit = None
            self._col = _Col()

        def Columns(self, col):
            return self._col

    class _Doc:
        def __init__(self):
            self._table = _Table()

        def Tables(self, idx):
            return self._table

    doc = _Doc()
    table = doc._table
    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: doc)
    app.set_table_col_width(1, 1, 400.0, doc_id="d1")
    assert table.AllowAutoFit is False
    assert table._col.Width == 400.0


def test_set_table_col_width_merged_region_raises_semantic_error(monkeypatch):
    # #23 场景 B：目标列落在合并区域内（Cell(r, col) 不独立存在）→ 语义化错误，非裸 COM
    if not _COM_ERROR:
        pytest.skip("pywin32 未装，无法构造 com_error")
    from offipy import word

    err = _COM_ERROR((-2147352567, "数值超出范围", None, 0))

    class _MergedCol:
        def __setattr__(self, name, value):
            if name == "Width":
                raise err
            object.__setattr__(self, name, value)

    class _Table:
        def __init__(self):
            self.AllowAutoFit = None
            self._col = _MergedCol()
            self.Rows = type("R", (), {"Count": 1})()

        def Columns(self, col):
            return self._col

        def Cell(self, row, col):
            raise err  # 该行目标列在合并区域内

    class _Doc:
        def __init__(self):
            self._table = _Table()

        def Tables(self, idx):
            return self._table

    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: _Doc())
    with pytest.raises(ComOperationError, match="被合并区域覆盖"):
        app.set_table_col_width(1, 2, 120.0, doc_id="d1")


# ================================================================ B. excel


def test_merge_cells_rejects_malformed_range():
    from offipy import excel

    app = excel.ExcelApp.__new__(excel.ExcelApp)
    with pytest.raises(InvalidArgumentError, match="非法区域"):
        app.merge_cells("数据", "????", doc_id="b1")
    with pytest.raises(InvalidArgumentError):
        app.merge_cells("数据", "A1ZZ", doc_id="b1")


def test_merge_cells_accepts_cell_and_range_forms(monkeypatch):
    from offipy import excel

    class _Rng:
        def __init__(self):
            self.merged = False

        def Merge(self):
            self.merged = True

    class _Ws:
        def __init__(self):
            self.rng = _Rng()

        def Range(self, addr):
            self.addr = addr
            return self.rng

    ws = _Ws()
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    monkeypatch.setattr(app, "_ws", lambda sheet, doc_id=None: ws)
    app.merge_cells("数据", "A1:B2", doc_id="b1")
    assert ws.addr == "A1:B2"
    assert ws.rng.merged is True
    app.merge_cells("数据", "C3", doc_id="b1")  # 单格合并同样合法
    assert ws.addr == "C3"


def test_unmerge_cells_rejects_malformed_range():
    from offipy import excel

    app = excel.ExcelApp.__new__(excel.ExcelApp)
    with pytest.raises(InvalidArgumentError, match="非法区域"):
        app.unmerge_cells("数据", "A1:B2:C3", doc_id="b1")


def test_open_book_missing_file_rejected(tmp_path):
    from offipy import excel

    app = excel.ExcelApp.__new__(excel.ExcelApp)
    with pytest.raises(InvalidArgumentError, match="源文件不存在"):
        app.open_book(str(tmp_path / "missing.xlsx"))


def test_set_range_dimension_mismatch_rejected(monkeypatch):
    # #24：2×3 目标给 1×3 数据 → 校验拦截，不再静默只填第一行
    from offipy import excel

    class _Rng:
        def __init__(self):
            self.Rows = type("C", (), {"Count": 2})()
            self.Columns = type("C", (), {"Count": 3})()
            self.Value = None

    class _Ws:
        def __init__(self):
            self.rng = _Rng()

        def Range(self, addr):
            return self.rng

    ws = _Ws()
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    monkeypatch.setattr(app, "_ws", lambda sheet, doc_id=None: ws)
    with pytest.raises(InvalidArgumentError, match="维度不符"):
        app.set_range("数据", "A2:C3", [[1, 2, 3]], doc_id="b1")
    assert ws.rng.Value is None  # 未写入任何数据


def test_set_range_matching_dimension_passes(monkeypatch):
    from offipy import excel

    class _Rng:
        def __init__(self):
            self.Rows = type("C", (), {"Count": 2})()
            self.Columns = type("C", (), {"Count": 3})()
            self.Value = None

    class _Ws:
        def __init__(self):
            self.rng = _Rng()

        def Range(self, addr):
            return self.rng

    ws = _Ws()
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    monkeypatch.setattr(app, "_ws", lambda sheet, doc_id=None: ws)
    app.set_range("数据", "A2:C3", [[1, 2, 3], [4, 5, 6]], doc_id="b1")
    assert ws.rng.Value == [[1, 2, 3], [4, 5, 6]]


def test_set_range_scalar_broadcasts(monkeypatch):
    # 标量（非 list/tuple）会广播填满整个范围，不校验维度
    from offipy import excel

    class _Rng:
        def __init__(self):
            self.Rows = type("C", (), {"Count": 2})()
            self.Columns = type("C", (), {"Count": 3})()
            self.Value = None

    class _Ws:
        def __init__(self):
            self.rng = _Rng()

        def Range(self, addr):
            return self.rng

    ws = _Ws()
    app = excel.ExcelApp.__new__(excel.ExcelApp)
    monkeypatch.setattr(app, "_ws", lambda sheet, doc_id=None: ws)
    app.set_range("数据", "A2:C3", 7, doc_id="b1")
    assert ws.rng.Value == 7


# ================================================================ C. ppt


def test_add_slide_rejects_non_integer_layout():
    from offipy import ppt

    app = ppt.PptApp.__new__(ppt.PptApp)
    with pytest.raises(InvalidArgumentError, match="非法 layout"):
        app.add_slide(layout="bad", doc_id="p1")
    with pytest.raises(InvalidArgumentError):
        app.add_slide(layout=2.5, doc_id="p1")


def test_add_slide_rejects_out_of_range_layout(monkeypatch):
    # 非整数先行拦截；越界整数走 COM 后转 InvalidArgumentError（不泄露原生枚举错误）
    if not _COM_ERROR:
        pytest.skip("pywin32 未装，无法构造 com_error")
    from offipy import ppt

    class _Slides:
        Count = 0

        def Add(self, index, layout):
            raise _COM_ERROR((-2147352567, "Invalid enumeration value", None, 0))

    class _Pres:
        Slides = _Slides()

    app = ppt.PptApp.__new__(ppt.PptApp)
    monkeypatch.setattr(app, "_require_pres", lambda doc_id: _Pres())
    with pytest.raises(InvalidArgumentError, match="非法 layout"):
        app.add_slide(layout=99, doc_id="p1")


def test_add_slide_accepts_valid_layout(monkeypatch):
    from offipy import ppt

    calls = []

    class _Slides:
        Count = 0

        def Add(self, index, layout):
            calls.append((index, layout))
            self.Count += 1
            return object()

    class _Pres:
        Slides = _Slides()

    app = ppt.PptApp.__new__(ppt.PptApp)
    monkeypatch.setattr(app, "_require_pres", lambda doc_id: _Pres())
    assert app.add_slide(layout=2, doc_id="p1") == 1
    assert calls == [(1, 2)]


def test_slide_index_out_of_range_rejected(monkeypatch):
    from offipy import ppt

    class _Slides:
        Count = 3

        def __call__(self, idx):
            return object()

    class _Pres:
        Slides = _Slides()

    app = ppt.PptApp.__new__(ppt.PptApp)
    monkeypatch.setattr(app, "_require_pres", lambda doc_id: _Pres())
    with pytest.raises(InvalidArgumentError, match="slide 99 越界，演示文稿共 3 页"):
        app.set_title(99, "x", doc_id="p1")
    with pytest.raises(InvalidArgumentError, match="slide 0 越界"):
        app.read_slide_texts(0, doc_id="p1")


def test_slide_index_valid_read_slide_texts(monkeypatch):
    # 回归：合法 slide 索引不误伤，空页返回空列表
    from offipy import ppt

    class _Shapes:
        Count = 0

    class _Slide:
        Shapes = _Shapes()

    class _Slides:
        Count = 3

        def __call__(self, idx):
            return _Slide()

    class _Pres:
        Slides = _Slides()

    app = ppt.PptApp.__new__(ppt.PptApp)
    monkeypatch.setattr(app, "_require_pres", lambda doc_id: _Pres())
    assert app.read_slide_texts(2, doc_id="p1") == []


def test_add_picture_missing_file_rejected(tmp_path):
    from offipy import ppt

    app = ppt.PptApp.__new__(ppt.PptApp)
    with pytest.raises(InvalidArgumentError, match="源文件不存在"):
        app.add_picture(1, str(tmp_path / "missing.png"), 0, 0, 10, 10, doc_id="p1")


def test_close_pres_save_false_closes_and_clears_doc(monkeypatch):
    from contextlib import nullcontext

    from offipy import ppt

    class _Pres:
        def __init__(self):
            self.Saved = None
            self.closed = False

        def Close(self):
            self.closed = True

    pres = _Pres()
    app = ppt.PptApp.__new__(ppt.PptApp)
    monkeypatch.setattr(app, "_require_pres", lambda doc_id: pres)
    monkeypatch.setattr(app, "_alerts_scope", lambda value=1: nullcontext())
    app._docs = {"pres1": pres}
    app._active_id = "pres1"
    out = app.close_pres(save=False, doc_id="pres1")
    assert pres.Saved is True  # 兜底防 Close 弹保存提示
    assert pres.closed is True
    assert out is None
    assert "pres1" not in app._docs
    assert app._active_id is None


def test_close_pres_save_true_saves_then_closes(monkeypatch):
    from contextlib import nullcontext

    from offipy import ppt

    class _Pres:
        Path = ""  # 从未保存 → 走 self.save 落盘

        def __init__(self):
            self.closed = False

        def Close(self):
            self.closed = True

    pres = _Pres()
    app = ppt.PptApp.__new__(ppt.PptApp)
    monkeypatch.setattr(app, "_require_pres", lambda doc_id: pres)
    monkeypatch.setattr(app, "_alerts_scope", lambda value=1: nullcontext())
    monkeypatch.setattr(app, "save", lambda doc_id=None: "C:/saved.pptx")
    app._docs = {"pres1": pres}
    app._active_id = "pres1"
    out = app.close_pres(save=True, doc_id="pres1")
    assert pres.closed is True
    assert out == "C:/saved.pptx"


# ================================================================ D. op() 只读 follow_active（#25）


def test_schema_supports_follow_active():
    from offipy import schema

    # 只读目标 op 显式放行 follow_active
    assert schema.supports_follow_active("excel", "get_cell") is True
    assert schema.supports_follow_active("excel", "read_range") is True
    assert schema.supports_follow_active("word", "read_doc_text") is True
    assert schema.supports_follow_active("ppt", "read_slide_texts") is True
    assert schema.supports_follow_active("ppt", "read_slide_summary") is True
    # destructive / requires_target 自动继承
    assert schema.supports_follow_active("excel", "set_cell") is True
    assert schema.supports_follow_active("ppt", "save_pdf") is True
    # 无目标语义的 op 不放行（new_*/activate/list_docs 不作用在活动目标上）
    assert schema.supports_follow_active("excel", "new_book") is False
    assert schema.supports_follow_active("excel", "activate") is False
    assert schema.supports_follow_active("word", "list_docs") is False


def test_schema_supports_follow_active_matches_flags():
    # 一致性：supports_follow_active ↔ destructive/requires_target/accepts_follow_active
    from offipy import schema

    for app in schema.apps():
        for op in schema.ops(app):
            spec = schema.spec(app, op)
            want = bool(spec.destructive or spec.requires_target or spec.accepts_follow_active)
            assert schema.supports_follow_active(app, op) is want, f"{app}.{op} 不一致"


def test_get_cell_follow_active_resolves_active_doc(monkeypatch):
    from offipy import excel

    captured = {}

    class _Ws:
        def Cells(self, row, col):
            return type("C", (), {"Value": 7})()

    app = excel.ExcelApp.__new__(excel.ExcelApp)
    monkeypatch.setattr(app, "get_target", lambda: {"doc_id": "active-book"})
    monkeypatch.setattr(
        app, "_ws", lambda sheet, doc_id=None: captured.__setitem__("doc_id", doc_id) or _Ws()
    )
    assert app.get_cell(1, "A1", follow_active=True) == 7
    assert captured["doc_id"] == "active-book"


def test_read_range_follow_active_resolves_active_doc(monkeypatch):
    from offipy import excel

    captured = {}

    class _Ws:
        def Range(self, addr):
            return type("R", (), {"Value": [["x"]], "Address": addr})()

    app = excel.ExcelApp.__new__(excel.ExcelApp)
    monkeypatch.setattr(app, "get_target", lambda: {"doc_id": "active-book"})
    monkeypatch.setattr(
        app, "_ws", lambda sheet, doc_id=None: captured.__setitem__("doc_id", doc_id) or _Ws()
    )
    assert app.read_range(1, "A1", follow_active=True) == [["x"]]
    assert captured["doc_id"] == "active-book"


def test_read_doc_text_follow_active_resolves_active_doc(monkeypatch):
    from offipy import word

    captured = {}

    class _Doc:
        @property
        def Content(self):
            return type("C", (), {"Text": "hi"})

    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "get_target", lambda: {"doc_id": "active-doc"})
    monkeypatch.setattr(
        app, "_require_doc", lambda doc_id=None: captured.__setitem__("doc_id", doc_id) or _Doc()
    )
    assert app.read_doc_text(follow_active=True) == "hi"
    assert captured["doc_id"] == "active-doc"


def test_readonly_follow_active_no_active_doc_raises(monkeypatch):
    from offipy import excel

    app = excel.ExcelApp.__new__(excel.ExcelApp)
    monkeypatch.setattr(app, "get_target", lambda: None)
    with pytest.raises(TargetNotFoundError, match="没有活动文档"):
        app.get_cell(1, "A1", follow_active=True)


def test_readonly_follow_active_false_keeps_active_fallback(monkeypatch):
    # follow_active=False（缺省）：doc_id 仍缺省走活动文档（只读语义不变）
    from offipy import excel

    captured = {}

    class _Ws:
        def Cells(self, row, col):
            return type("C", (), {"Value": 7})()

    app = excel.ExcelApp.__new__(excel.ExcelApp)
    monkeypatch.setattr(
        app, "_ws", lambda sheet, doc_id=None: captured.__setitem__("doc_id", doc_id) or _Ws()
    )
    assert app.get_cell(1, "A1") == 7
    assert captured["doc_id"] is None


def test_remote_facade_exposes_follow_active_on_readonly():
    import inspect

    from offipy import api

    r = api.RemoteExcel(base_url="http://127.0.0.1:1")
    assert "follow_active" in inspect.signature(r.get_cell).parameters
    assert "expected_target" not in inspect.signature(r.get_cell).parameters
    assert "follow_active" in inspect.signature(r.read_range).parameters
    assert "expected_target" not in inspect.signature(r.read_range).parameters
    # 破坏性 op 仍同时带 expected_target + follow_active
    assert "expected_target" in inspect.signature(r.set_cell).parameters
    assert "follow_active" in inspect.signature(r.set_cell).parameters
    # 无目标语义的只读 op 不放行
    assert "follow_active" not in inspect.signature(r.new_book).parameters
