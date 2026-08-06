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
    # #16：段落符 \r\n / \r 归一化成 \n（golden 随归一化规则更新）
    assert w.read_doc_text() == "第一段\n第二段\n第三段"


class _Text:
    def __init__(self, value):
        self.Text = value


class _TextFrame:
    def __init__(self, text):
        self.TextRange = _Text(text)


class _FakeShape:
    """COM shape 最小 fake：覆盖 read_slide_texts/read_slide_summary 用到的属性面。"""

    def __init__(
        self,
        sid,
        text,
        *,
        name="shp",
        kind=1,
        ph_type=None,
        left=0.0,
        top=0.0,
        width=100.0,
        height=20.0,
        has_text=True,
        group_items=None,
    ):
        self.Id = sid
        self.Name = name
        self.Type = kind  # 1=AutoShape, 6=Group, 14=Placeholder
        self.Left = left
        self.Top = top
        self.Width = width
        self.Height = height
        self.Rotation = 0.0
        self.ZOrderPosition = sid
        self.HasTextFrame = -1 if has_text else 0
        if has_text:
            self.TextFrame = _TextFrame(text)
        if ph_type is not None:
            self.PlaceholderFormat = SimpleNamespace(Type=ph_type)
        if group_items is not None:
            self.GroupItems = group_items


class _FakePlaceholders:
    def __init__(self, shapes):
        self._shapes = list(shapes)
        self.Count = len(self._shapes)

    def __call__(self, idx):
        return self._shapes[idx - 1]


class _FakeShapes:
    def __init__(self, shapes):
        self._shapes = list(shapes)
        self.Count = len(self._shapes)
        self.Placeholders = _FakePlaceholders([s for s in self._shapes if s.Type == 14])

    def __call__(self, idx):
        return self._shapes[idx - 1]


class _Slides:
    def __init__(self, items):
        self.Count = len(items)
        self._items = items

    def __call__(self, idx):
        return self._items[idx - 1]


def _slide(title, body, notes_text, has_title=True):
    shapes = []
    if has_title:
        shapes.append(
            _FakeShape(
                1, title, name="title", kind=14, ph_type=1, left=36, top=18, width=600, height=72
            )
        )
    shapes.append(
        _FakeShape(2, body, name="body", kind=14, ph_type=2, left=36, top=90, width=600, height=396)
    )
    notes = _FakeShape(99, notes_text, name="notes", kind=14, ph_type=2)
    return SimpleNamespace(
        Shapes=_FakeShapes(shapes),
        NotesPage=SimpleNamespace(Shapes=_FakeShapes([notes])),
    )


def test_ppt_read_slide_summary_golden():
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
    assert p.read_slide_summary() == [
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
