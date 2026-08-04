"""Agent 只读 op 纯逻辑测试：word/ppt/excel 读回文本层（不触真实 COM）。

用 `__new__` 构造实例跳过 __init__ 的 COM 初始化，再注入 fake 文档对象，
断言返回结构稳定（供 Agent 迭代）。
"""

from types import SimpleNamespace

from offipy import excel, ppt, word


def test_word_read_doc_text_returns_full_text():
    doc = SimpleNamespace(Content=SimpleNamespace(Text="第一段\r\n第二段"))
    w = word.WordApp.__new__(word.WordApp)
    w.active_doc = lambda doc_id=None: doc
    assert w.read_doc_text() == "第一段\r\n第二段"


class _Text:
    def __init__(self, value):
        self.Text = value


class _Shape:
    def __init__(self, text):
        self.TextFrame = SimpleNamespace(TextRange=_Text(text))


def _slide(title_text, body_text, notes_text, has_title=True):
    shapes = SimpleNamespace()
    shapes.HasTitle = has_title
    if has_title:
        shapes.Title = _Shape(title_text)

    def body_placeholders(idx):
        if idx == 2:
            return _Shape(body_text)
        raise AttributeError("no such placeholder")

    def notes_placeholders(idx):
        if idx == 2:
            return _Shape(notes_text)
        raise AttributeError("no such placeholder")

    shapes.Placeholders = body_placeholders
    notes = SimpleNamespace(Shapes=SimpleNamespace(Placeholders=notes_placeholders))
    return SimpleNamespace(Shapes=shapes, NotesPage=notes)


class _Slides:
    def __init__(self, items):
        self.Count = len(items)
        self._items = items

    def __call__(self, idx):
        return self._items[idx - 1]


def test_ppt_read_slide_texts_structure():
    pres = SimpleNamespace(
        Slides=_Slides([_slide("标题A", "正文A", "备注A"), _slide("标题B", "正文B", "备注B")])
    )
    p = ppt.PptApp.__new__(ppt.PptApp)
    p.active_pres = lambda doc_id=None: pres
    result = p.read_slide_texts()
    assert result == [
        {"index": 1, "title": "标题A", "body": "正文A", "notes": "备注A"},
        {"index": 2, "title": "标题B", "body": "正文B", "notes": "备注B"},
    ]


def test_ppt_read_slide_texts_missing_fields_empty():
    # 无标题（HasTitle=False）+ 无正文占位符（Placeholders 抛错）→ 空串兜底
    slides = _Slides([_slide("", "", "备注", has_title=False)])
    pres = SimpleNamespace(Slides=slides)
    p = ppt.PptApp.__new__(ppt.PptApp)
    p.active_pres = lambda doc_id=None: pres
    result = p.read_slide_texts()
    assert result[0]["title"] == ""
    assert result[0]["body"] == ""
    assert result[0]["notes"] == "备注"


def test_excel_read_range_returns_2d_list():
    ws = SimpleNamespace(Range=lambda addr: SimpleNamespace(Value=((1, 2), (3, 4))))
    b = excel.ExcelApp.__new__(excel.ExcelApp)
    b._ws = lambda sheet, doc_id=None: ws
    assert b.read_range("Sheet1", "A1:B2") == [[1, 2], [3, 4]]


def test_excel_normalize_range_forms():
    assert excel._normalize_range(None) == []
    assert excel._normalize_range(100) == [[100]]
    assert excel._normalize_range(((1, 2),)) == [[1, 2]]  # 单行
    assert excel._normalize_range(((1,), (2,))) == [[1], [2]]  # 单列
    assert excel._normalize_range(((1, 2), (3, 4))) == [[1, 2], [3, 4]]
