"""读 op 结构 golden（P2-6）：word/ppt/excel 读回文本/值层的全结构快照。

用 `__new__` 构造实例 + 固定 fake 文档对象，纯逻辑、不触 COM。任何读 op
返回结构的改动（字段名、顺序、缺省值、归一化规则）都会在此被抓住。
"""

from types import SimpleNamespace

from offipy import excel, ppt, word


def test_word_read_doc_text_golden():
    doc = SimpleNamespace(Content=SimpleNamespace(Text="第一段\r\n第二段\r\n第三段"))
    w = word.WordApp.__new__(word.WordApp)
    w.active_doc = lambda doc_id=None: doc
    assert w.read_doc_text() == "第一段\r\n第二段\r\n第三段"


class _Text:
    def __init__(self, value):
        self.Text = value


class _Shape:
    def __init__(self, text):
        self.TextFrame = SimpleNamespace(TextRange=_Text(text))


def _slide(title, body, notes_text, has_title=True):
    shapes = SimpleNamespace()
    shapes.HasTitle = has_title
    if has_title:
        shapes.Title = _Shape(title)

    def body_placeholders(idx):
        if idx == 2:
            return _Shape(body)
        raise AttributeError("no such placeholder")

    def notes_placeholders(idx):
        if idx == 2:
            return _Shape(notes_text)
        raise AttributeError("no such placeholder")

    shapes.Placeholders = body_placeholders
    notes_page = SimpleNamespace(Shapes=SimpleNamespace(Placeholders=notes_placeholders))
    return SimpleNamespace(Shapes=shapes, NotesPage=notes_page)


class _Slides:
    def __init__(self, items):
        self.Count = len(items)
        self._items = items

    def __call__(self, idx):
        return self._items[idx - 1]


def test_ppt_read_slide_texts_golden():
    pres = SimpleNamespace(
        Slides=_Slides(
            [
                _slide("增长概览", "MAU +18%\n留存 41%", "对应第 2 页图表"),
                _slide("季度规划", "上线推荐位改版", "负责人：产品组"),
                _slide("", "", "无标题页备注", has_title=False),
            ]
        )
    )
    p = ppt.PptApp.__new__(ppt.PptApp)
    p.active_pres = lambda doc_id=None: pres
    assert p.read_slide_texts() == [
        {"index": 1, "title": "增长概览", "body": "MAU +18%\n留存 41%", "notes": "对应第 2 页图表"},
        {"index": 2, "title": "季度规划", "body": "上线推荐位改版", "notes": "负责人：产品组"},
        {"index": 3, "title": "", "body": "", "notes": "无标题页备注"},
    ]


def test_excel_read_range_golden():
    ws = SimpleNamespace(Range=lambda addr: SimpleNamespace(Value=((1, 2), (3, 4))))
    b = excel.ExcelApp.__new__(excel.ExcelApp)
    b._ws = lambda sheet, doc_id=None: ws
    assert b.read_range("Sheet1", "A1:B2") == [[1, 2], [3, 4]]


def test_excel_normalize_range_golden():
    assert excel._normalize_range(None) == []
    assert excel._normalize_range(42) == [[42]]
    assert excel._normalize_range(((1, 2),)) == [[1, 2]]
    assert excel._normalize_range(((1,), (2,))) == [[1], [2]]
    assert excel._normalize_range(((1, 2), (3, 4))) == [[1, 2], [3, 4]]
