"""Word add_page_number 三模式分支行为的 fake-COM 单测（不碰 Office）。

用轻量 fake 模拟页脚 Range / Fields / TabStops / PageSetup，验证
replace / append / standalone 的分支、幂等、错误校验与只样式化页码域。
实证依据见 docs/development/probe_word_page_number.md（gitignored）。
"""

import pytest

from offipy.exceptions import InvalidArgumentError
from offipy.word import WordApp, _rgb

_WD_FIELD_PAGE = 33


# --- fake COM 层级 -----------------------------------------------------------


class _FakeFont:
    def __init__(self):
        self.Color = None
        self.Size = None


class _FakeFieldResult:
    def __init__(self, text):
        self.Text = text
        self.Font = _FakeFont()


class _FakeField:
    def __init__(self, fields, range_, type_, result):
        self._fields = fields
        self._range = range_
        self.Type = type_
        self.Result = _FakeFieldResult(result)

    def Delete(self):
        """Word 语义：删除域会一并移除其结果显示文本。"""
        if self in self._fields._items:
            self._fields._items.remove(self)
        if self._range is not None:
            self._range._text = self._range._text.replace(self.Result.Text, "", 1)


class _FakeFields:
    def __init__(self):
        self._items = []

    def __iter__(self):
        return iter(list(self._items))

    def __len__(self):
        return len(self._items)

    def Add(self, range_, type_):
        result = "1"
        fld = _FakeField(self, range_, type_, result)
        self._items.append(fld)
        if range_ is not None:
            text = range_._text
            if text.endswith("\r"):
                range_._text = text[:-1] + result + "\r"
            else:
                range_._text = text + result
        return fld


class _FakeTabStops:
    def __init__(self):
        self.added = []

    def Add(self, position, alignment):
        self.added.append((position, alignment))


class _FakeParagraphFormat:
    def __init__(self):
        self.Alignment = None
        self.TabStops = _FakeTabStops()


class _FakeRange:
    def __init__(self, fields):
        self._fields = fields
        self._text = "\r"
        self.ParagraphFormat = _FakeParagraphFormat()

    @property
    def Fields(self):
        return self._fields

    @property
    def Text(self):
        return self._text

    @Text.setter
    def Text(self, value):
        self._fields._items.clear()  # Word：写 Range.Text 会清除既有域
        self._text = value if value.endswith("\r") else value + "\r"

    def MoveEnd(self, unit, count):
        pass

    def Collapse(self, direction):
        pass


class _FakeFooter:
    def __init__(self):
        self._fields = _FakeFields()
        self._range = _FakeRange(self._fields)

    @property
    def Range(self):
        return self._range


class _FakeSections:
    def __init__(self, footer):
        self._footer = footer

    def Footers(self, idx):
        return self._footer


class _FakePageSetup:
    PageWidth = 612.0
    LeftMargin = 72.0
    RightMargin = 72.0


class _FakeDoc:
    def __init__(self):
        self.footer = _FakeFooter()
        self.PageSetup = _FakePageSetup()

    def Sections(self, idx):
        return _FakeSections(self.footer)


def _app(doc):
    app = WordApp.__new__(WordApp)
    app._require_doc = lambda doc_id=None: doc
    return app


def _page_fields(hf):
    return [f.Type for f in hf.Range.Fields]


# --- replace（默认模式） ------------------------------------------------------


def test_add_page_number_replace_clears_and_aligns_paragraph():
    doc = _FakeDoc()
    app = _app(doc)
    result = app.add_page_number(doc_id="x")
    assert result == "1\r"  # 单段落符，无 '1\r\r' 双段落伪影
    hf = doc.footer
    assert hf.Range.ParagraphFormat.Alignment == 2  # 默认 right
    assert _page_fields(hf) == [_WD_FIELD_PAGE]


def test_add_page_number_replace_center_aligns_paragraph():
    doc = _FakeDoc()
    app = _app(doc)
    app.add_page_number(alignment="center", doc_id="x")
    assert doc.footer.Range.ParagraphFormat.Alignment == 1


def test_add_page_number_replace_styles_field_only():
    doc = _FakeDoc()
    app = _app(doc)
    app.add_page_number(color="#2251FF", size=18, doc_id="x")
    fld = next(f for f in doc.footer.Range.Fields if f.Type == _WD_FIELD_PAGE)
    assert fld.Result.Font.Color == _rgb("#2251FF")
    assert fld.Result.Font.Size == 18


def test_add_page_number_replace_is_idempotent():
    doc = _FakeDoc()
    app = _app(doc)
    app.add_page_number(doc_id="x")
    result = app.add_page_number(doc_id="x")
    assert result == "1\r"
    assert _page_fields(doc.footer) == [_WD_FIELD_PAGE]


# --- append ------------------------------------------------------------------


def test_add_page_number_append_after_text():
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    result = app.add_page_number(mode="append", doc_id="x")
    assert result == "公司名1\r"
    assert _page_fields(doc.footer) == [_WD_FIELD_PAGE]


def test_add_page_number_append_applies_paragraph_alignment():
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    app.add_page_number(mode="append", alignment="center", doc_id="x")
    assert doc.footer.Range.ParagraphFormat.Alignment == 1


def test_add_page_number_append_idempotent_skips_duplicate():
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    app.add_page_number(mode="append", doc_id="x")
    result = app.add_page_number(mode="append", doc_id="x")
    assert result == "公司名1\r"  # 不叠加成 '公司名11\r'
    assert _page_fields(doc.footer) == [_WD_FIELD_PAGE]


def test_add_page_number_append_empty_footer():
    doc = _FakeDoc()  # 空页脚 Range.Text == '\r'
    app = _app(doc)
    result = app.add_page_number(mode="append", doc_id="x")
    assert result == "1\r"
    assert _page_fields(doc.footer) == [_WD_FIELD_PAGE]


def test_add_page_number_append_non_page_field_does_not_block():
    # 幂等只认 PAGE 域（Type==33）：存在非 PAGE 域时仍应追加页码域
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    doc.footer.Range.Fields.Add(doc.footer.Range, 21)  # TIME 域
    app = _app(doc)
    app.add_page_number(mode="append", doc_id="x")
    assert sorted(_page_fields(doc.footer)) == [21, _WD_FIELD_PAGE]


def test_add_page_number_append_styles_field_only():
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    app.add_page_number(mode="append", color="#2251FF", size=18, doc_id="x")
    fld = next(f for f in doc.footer.Range.Fields if f.Type == _WD_FIELD_PAGE)
    assert fld.Result.Font.Color == _rgb("#2251FF")
    assert fld.Result.Font.Size == 18


# --- standalone ---------------------------------------------------------------


def test_add_page_number_standalone_right_tab_zone():
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    result = app.add_page_number(mode="standalone", doc_id="x")
    assert result == "公司名\t1\r"
    # fake PageSetup: 612 - 72 - 72 = 468pt；right 制表位 ≈ 文本宽
    assert doc.footer.Range.ParagraphFormat.TabStops.added == [(468.0, 2)]
    assert _page_fields(doc.footer) == [_WD_FIELD_PAGE]


def test_add_page_number_standalone_center_tab_zone():
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    app.add_page_number(mode="standalone", alignment="center", doc_id="x")
    assert doc.footer.Range.ParagraphFormat.TabStops.added == [(234.0, 1)]


def test_add_page_number_standalone_left_tab_zone():
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    app.add_page_number(mode="standalone", alignment="left", doc_id="x")
    assert doc.footer.Range.ParagraphFormat.TabStops.added == [(0.0, 0)]


def test_add_page_number_standalone_rebuilds_idempotent():
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    app.add_page_number(mode="standalone", doc_id="x")
    assert doc.footer.Range.Text == "公司名\t1\r"
    result = app.add_page_number(mode="standalone", doc_id="x")  # 幂等重建
    assert result == "公司名\t1\r"
    assert _page_fields(doc.footer) == [_WD_FIELD_PAGE]


def test_add_page_number_standalone_empty_footer():
    doc = _FakeDoc()
    app = _app(doc)
    result = app.add_page_number(mode="standalone", doc_id="x")
    assert result == "\t1\r"
    assert _page_fields(doc.footer) == [_WD_FIELD_PAGE]


# --- 校验 ----------------------------------------------------------------------


def test_add_page_number_invalid_mode_rejected():
    doc = _FakeDoc()
    app = _app(doc)
    with pytest.raises(InvalidArgumentError, match="未知模式"):
        app.add_page_number(mode="foo", doc_id="x")
    with pytest.raises(InvalidArgumentError, match="未知模式"):
        app.add_page_number(mode="append_extra", doc_id="x")
    with pytest.raises(InvalidArgumentError, match="未知模式"):
        app.add_page_number(mode="", doc_id="x")


def test_add_page_number_mode_is_case_insensitive():
    # 与 alignment 的 _resolve_style 一致：mode 也按 strip().lower() 归一化
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    result = app.add_page_number(mode="APPEND", doc_id="x")
    assert result == "公司名1\r"


def test_add_page_number_size_nonpositive_rejected():
    doc = _FakeDoc()
    app = _app(doc)
    with pytest.raises(InvalidArgumentError, match="字号必须为正数"):
        app.add_page_number(size=0, doc_id="x")
    with pytest.raises(InvalidArgumentError, match="字号必须为正数"):
        app.add_page_number(size=-1.5, doc_id="x")
    # 正数放行（不抛）
    app.add_page_number(size=12, doc_id="x")
