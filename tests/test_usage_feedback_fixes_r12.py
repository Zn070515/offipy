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
from offipy.exceptions import ComOperationError, InvalidArgumentError

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
