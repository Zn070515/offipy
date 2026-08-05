"""PptApp set_title/set_body/set_notes 占位符显式定位（不碰 COM）。

fake slide：按占位符类型找 shape（不硬编码 Placeholders(2) 序号），
找不到自动建文本框并返回实际修改的 shape ID。传 doc_id 放行 destructive
wrapper 后，方法体用 fake _require_pres 命中 fake slide。
"""

from types import SimpleNamespace

import pytest

from offipy.exceptions import InvalidArgumentError
from offipy.ppt import PptApp

# 微软官方 PpPlaceholderType 值（round-10 探针运行时常量 20/20 核实）。
# fixture 必须用官方值构造，不得用项目自身常量自证（防「错误实现+错误测试彼此一致」）。
MS_TITLE = 1  # ppPlaceholderTitle
MS_BODY = 2  # ppPlaceholderBody
MS_CENTER_TITLE = 3  # ppPlaceholderCenterTitle

DID = "x"


class _FakeRange:
    def __init__(self):
        self.Text = None


class _FakeShape:
    def __init__(self, kind, sid):
        self.PlaceholderFormat = SimpleNamespace(Type=kind)
        self.Id = sid
        self.TextFrame = SimpleNamespace(TextRange=_FakeRange())


class _FakePlaceholders:
    def __init__(self, shapes):
        self._shapes = list(shapes)
        self.Count = len(self._shapes)

    def __call__(self, i):
        return self._shapes[i - 1]


class _FakeShapes:
    def __init__(self, shapes, next_id=900):
        self._shapes = list(shapes)
        self._pls = _FakePlaceholders(self._shapes)
        self._next = next_id
        self.added = []

    @property
    def Placeholders(self):
        return self._pls

    def AddTextbox(self, orientation, left, top, width, height):
        self.added.append((orientation, left, top, width, height))
        shape = _FakeShape(None, self._next)
        self._next += 1
        self._shapes.append(shape)
        self._pls = _FakePlaceholders(self._shapes)
        return shape


def _app(slide_shapes, notes_shapes=None):
    slide = SimpleNamespace(
        Shapes=_FakeShapes(slide_shapes),
        NotesPage=SimpleNamespace(Shapes=_FakeShapes(notes_shapes or [])),
    )
    app = PptApp.__new__(PptApp)
    app._require_pres = lambda doc_id=None: SimpleNamespace(Slides=lambda i: slide)
    return app


def test_set_title_hits_title_placeholder():
    shape = _FakeShape(MS_TITLE, 11)
    app = _app([shape])
    assert app.set_title(1, "标题", doc_id=DID) == 11
    assert shape.TextFrame.TextRange.Text == "标题"


def test_set_title_hits_center_title_placeholder():
    shape = _FakeShape(MS_CENTER_TITLE, 12)
    app = _app([shape])
    assert app.set_title(1, "居中标题", doc_id=DID) == 12
    assert shape.TextFrame.TextRange.Text == "居中标题"


def test_set_title_auto_adds_textbox_when_missing():
    body = _FakeShape(MS_BODY, 2)
    app = _app([body])
    shapes = app._require_pres().Slides(1).Shapes
    assert app.set_title(1, "自动建框", doc_id=DID) == 900  # 新 AddTextbox 的 Id
    assert len(shapes.added) == 1
    assert shapes.added[0][0] == 1  # msoTextOrientationHorizontal


def test_set_body_hits_body_placeholder():
    shape = _FakeShape(MS_BODY, 7)
    app = _app([shape])
    assert app.set_body(1, ["a", "b"], doc_id=DID) == 7
    assert shape.TextFrame.TextRange.Text == "a\rb"


def test_set_body_auto_adds_textbox_when_missing():
    title = _FakeShape(MS_TITLE, 1)
    app = _app([title])
    shapes = app._require_pres().Slides(1).Shapes
    assert app.set_body(1, ["x"], doc_id=DID) == 900
    assert len(shapes.added) == 1


def test_set_notes_hits_body_placeholder():
    shape = _FakeShape(MS_BODY, 5)
    app = _app([], notes_shapes=[shape])
    assert app.set_notes(1, "备注", doc_id=DID) == 5
    assert shape.TextFrame.TextRange.Text == "备注"


def test_set_notes_auto_adds_textbox_when_missing():
    app = _app([], notes_shapes=[])
    notes_shapes = app._require_pres().Slides(1).NotesPage.Shapes
    assert app.set_notes(1, "自动建", doc_id=DID) == 900
    assert len(notes_shapes.added) == 1


@pytest.mark.parametrize(
    "method, args",
    [
        ("set_title", (1, "")),
        ("set_title", (1, None)),
        ("set_body", (1, [])),
        ("set_body", (1, None)),
    ],
)
def test_empty_input_still_rejected(method, args):
    app = _app([_FakeShape(MS_TITLE, 1)])
    with pytest.raises(InvalidArgumentError):
        getattr(app, method)(*args, doc_id=DID)
