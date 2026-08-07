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
            footer = self._range._footer
            footer._text = footer._text.replace(self.Result.Text, "", 1)


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
            footer = range_._footer
            text = footer._text
            if text.endswith("\r"):
                footer._text = text[:-1] + result + "\r"
            else:
                footer._text = text + result
        return fld


class _FakeTabStops:
    def __init__(self):
        self.added = []
        self.clear_count = 0

    def Add(self, position, alignment):
        self.added.append((position, alignment))

    def ClearAll(self):
        self.clear_count += 1
        self.added = []


class _FakeParagraphFormat:
    def __init__(self):
        self.Alignment = None
        self.TabStops = _FakeTabStops()


class _FakeRange:
    """每次 hf.Range 访问返回新的 Range 视图，共享页脚文本/域/段落格式。

    折叠态（MoveEnd/Collapse 后）写 Text = 在段落符前插入文本，不清除既有域；
    整段重写（replace 清空）才清域——与 Word 实机行为一致。
    """

    def __init__(self, footer):
        self._footer = footer
        self._collapsed = False

    @property
    def Fields(self):
        return self._footer.fields

    @property
    def ParagraphFormat(self):
        return self._footer.paragraph_format

    @property
    def Text(self):
        return self._footer._text

    @Text.setter
    def Text(self, value):
        if self._collapsed:
            if self._footer._text.endswith("\r"):
                self._footer._text = self._footer._text[:-1] + value + "\r"
            else:
                self._footer._text = self._footer._text + value
        else:
            self._footer.fields._items.clear()
            self._footer._text = value if value.endswith("\r") else value + "\r"

    def MoveEnd(self, unit, count):
        self._collapsed = True

    def MoveStart(self, unit, count):
        self._collapsed = True

    def Collapse(self, direction):
        self._collapsed = True

    def Delete(self):
        """折叠态删除选中文本：建模为删掉尾部一个非段落符字符（standalone 左模式
        清遗留尾随制表符用）。"""
        if self._collapsed:
            t = self._footer._text
            if t.endswith("\r"):
                self._footer._text = t[:-2] + "\r" if len(t) > 1 else "\r"
            else:
                self._footer._text = t[:-1]


class _FakeFooter:
    def __init__(self):
        self.fields = _FakeFields()
        self.paragraph_format = _FakeParagraphFormat()
        self._text = "\r"

    @property
    def Range(self):
        return _FakeRange(self)


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


def test_add_page_number_standalone_left_follows_text():
    # 左对齐无制表区：不插制表符、不设制表位（Word 丢弃 position 0 制表位），
    # 页码紧跟页脚文本自然左排（实探见 docs/development/probe_word_page_number.md）
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    result = app.add_page_number(mode="standalone", alignment="left", doc_id="x")
    assert result == "公司名1\r"
    assert doc.footer.Range.ParagraphFormat.TabStops.added == []
    assert _page_fields(doc.footer) == [_WD_FIELD_PAGE]


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


def test_add_page_number_standalone_clears_stale_tab_stops():
    # 右→左切换：旧右制表位被 ClearAll 清掉，遗留尾随制表符也被移除，
    # 页码紧跟文本左排（不设制表位，Word 丢弃 position 0）
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名\t"  # 模拟上次 right 遗留的尾随制表符
    app = _app(doc)
    # 预置一个「陈旧」右制表位（模拟上次 right 调用遗留）
    doc.footer.Range.ParagraphFormat.TabStops.Add(468.0, 2)
    app.add_page_number(mode="standalone", alignment="left", doc_id="x")
    ts = doc.footer.Range.ParagraphFormat.TabStops
    assert ts.clear_count >= 1
    assert ts.added == []  # 左模式不设制表位
    assert doc.footer.Range.Text == "公司名1\r"  # 尾随制表符已移除，页码紧跟文本


def test_add_page_number_standalone_preserves_non_page_field():
    # standalone 不重写整段：既有的非 PAGE 域（如 DATE）必须原样保留
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    doc.footer.Range.Fields.Add(doc.footer.Range, 21)  # DATE 域
    app = _app(doc)
    app.add_page_number(mode="standalone", doc_id="x")
    assert sorted(_page_fields(doc.footer)) == [21, _WD_FIELD_PAGE]


def test_add_page_number_standalone_resets_paragraph_alignment():
    # standalone 靠制表位定位：必须把段落对齐重置为 left（0），否则先前
    # replace/append 设的对齐会让文本与制表位参考错位（final review #1）
    doc = _FakeDoc()
    doc.footer.Range.Text = "公司名"
    app = _app(doc)
    app.add_page_number(mode="append", alignment="center", doc_id="x")
    assert doc.footer.Range.ParagraphFormat.Alignment == 1  # 先被 append 设为中心
    app.add_page_number(mode="standalone", alignment="right", doc_id="x")
    assert doc.footer.Range.ParagraphFormat.Alignment == 0  # 重置回 left
    assert doc.footer.Range.ParagraphFormat.TabStops.added == [(468.0, 2)]


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


def test_add_page_number_non_str_mode_rejected():
    # 非字符串 mode（int/None）不得落到 .strip() 抛 AttributeError
    doc = _FakeDoc()
    app = _app(doc)
    with pytest.raises(InvalidArgumentError, match="未知模式"):
        app.add_page_number(mode=5, doc_id="x")
    with pytest.raises(InvalidArgumentError, match="未知模式"):
        app.add_page_number(mode=None, doc_id="x")


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


# --- S4 Task 3：line_spacing 数值/字符串双轨（fake-COM） ----------------------


class _FakeParagraphFormat2:
    def __init__(self):
        self.Alignment = None
        self.LineSpacingRule = None
        self.SpaceBefore = None
        self.SpaceAfter = None
        self.LeftIndent = None
        self.FirstLineIndent = None


class _FakeSpacingParas:
    def __init__(self, count=1):
        self.Count = count
        self._fmt = _FakeParagraphFormat2()

    def __call__(self, idx):
        return type("_P", (), {"Format": self._fmt})()


class _FakeSpacingDoc:
    def __init__(self, count=1):
        self.Paragraphs = _FakeSpacingParas(count)


def _spacing_app(doc):
    app = WordApp.__new__(WordApp)
    app._require_doc = lambda doc_id=None: doc
    return app


def test_format_paragraph_line_spacing_1_is_single():
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    app.format_paragraph(paragraph=1, line_spacing=1, doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 0  # wdLineSpaceSingle


def test_format_paragraph_line_spacing_1_0_is_single():
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    app.format_paragraph(paragraph=1, line_spacing=1.0, doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 0


def test_format_paragraph_line_spacing_1_5_is_one_and_half():
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    app.format_paragraph(paragraph=1, line_spacing=1.5, doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 1  # wdLineSpace1pt5


def test_format_paragraph_line_spacing_2_is_double():
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    app.format_paragraph(paragraph=1, line_spacing=2, doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 2  # wdLineSpaceDouble


def test_format_paragraph_line_spacing_2_0_is_double():
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    app.format_paragraph(paragraph=1, line_spacing=2.0, doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 2


def test_format_paragraph_line_spacing_string_still_works():
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    app.format_paragraph(paragraph=1, line_spacing="single", doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 0
    app.format_paragraph(paragraph=1, line_spacing="1.5", doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 1


def test_format_paragraph_line_spacing_cli_numeric_strings():
    # CLI 的 --line_spacing 以字符串传入，'1'/'2' 必须归一成 single/double 键，
    # 否则 '2' 会落入未知行距报错（final review #2）
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    app.format_paragraph(paragraph=1, line_spacing="1", doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 0  # wdLineSpaceSingle
    app.format_paragraph(paragraph=1, line_spacing="1.0", doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 0
    app.format_paragraph(paragraph=1, line_spacing="2", doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 2  # wdLineSpaceDouble
    app.format_paragraph(paragraph=1, line_spacing="2.0", doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule == 2


def test_format_paragraph_line_spacing_invalid_numeric_rejected():
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    with pytest.raises(InvalidArgumentError, match="非法数值行距"):
        app.format_paragraph(paragraph=1, line_spacing=1.7, doc_id="x")


def test_format_paragraph_line_spacing_bool_rejected():
    # bool 是 int 子类：True/False 不得被当作 1/0 行距
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    for bad in (True, False):
        with pytest.raises(InvalidArgumentError, match="非法行距类型"):
            app.format_paragraph(paragraph=1, line_spacing=bad, doc_id="x")


def test_format_paragraph_line_spacing_none_leaves_rule_untouched():
    doc = _FakeSpacingDoc()
    app = _spacing_app(doc)
    app.format_paragraph(paragraph=1, doc_id="x")
    assert doc.Paragraphs(1).Format.LineSpacingRule is None


# --- #43：add_heading level 越界显式拒绝（曾静默降级为 Heading 1） ---


def test_add_heading_rejects_out_of_range_level():
    app = WordApp.__new__(WordApp)

    def _boom(*a, **k):
        raise AssertionError("越界 level 不应触达 COM 路径")

    app.write_line = _boom
    app._require_doc = _boom
    for bad in (0, 4, -1, 99):
        with pytest.raises(InvalidArgumentError, match="level 必须为 1/2/3"):
            app.add_heading("x", level=bad, doc_id="x")
