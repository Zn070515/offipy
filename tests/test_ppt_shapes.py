"""read_shapes 纯逻辑测试（S1 Task 2）。

用 `__new__` 构造 PptApp 跳过 COM 初始化，注入 fake COM 文档树。覆盖：
- 全 shape 类型（文本/图片/线条/自选图形/占位符/group）都会返回
- recursive=True 含 group 容器 + 后代；recursive=False 仅顶层
- parent_shape_id / group_path（外层→内层）/ coordinate_space
- 旋转 group 后代 coordinate_space="unknown"
- 严格 shape_id：读不到抛 ComOperationError（无 0 兜底）
- solid vs 非 solid fill / 可见 vs 隐藏 line → 颜色规则
- 首 run 字体语义 + 空文本字体全 None
- visible 混合态 → None
- group 子元素 z_order 为所在集合 1-based（兄弟排序，抵消绝对偏移）
- read_slide_texts 回归：S1 不改动普通文本读
"""

from types import SimpleNamespace

import pytest

from offipy import ppt
from offipy.exceptions import ComOperationError, InvalidArgumentError

PAGE_W = 720.0
PAGE_H = 540.0

# COM RGB（BGR 打包）→ hex 参考（独立换算，防实现自证）
_RED_RGB = 0xFF  # r=255,g=0,b=0
_PURPLE_RGB = 128 | (0 << 8) | (128 << 16)  # 8388736


class _FakeFont:
    def __init__(self, size=None, name=None, color=None):
        self.Size = size
        self.Name = name
        self.Color = SimpleNamespace(RGB=color)


class _FakeRun:
    def __init__(self, text, font):
        self.Text = text
        self.Font = font


class _FakeTextRange:
    def __init__(self, text, runs=None):
        self.Text = text
        self._runs = runs or [_FakeRun(text, _FakeFont())]

    def Runs(self, idx=None):
        if idx is None:
            return SimpleNamespace(Count=len(self._runs))
        return self._runs[idx - 1]


class _FakeTextFrame:
    def __init__(self, text, runs=None):
        self.TextRange = _FakeTextRange(text, runs)


class _FakeFill:
    def __init__(self, fill_type=1, rgb=None, transparency=None):
        self.Type = fill_type
        if rgb is not None:
            self.ForeColor = SimpleNamespace(RGB=rgb)
        self.Transparency = transparency


class _FakeLine:
    def __init__(self, visible=-1, rgb=None, weight=None):
        self.Visible = visible
        if rgb is not None:
            self.ForeColor = SimpleNamespace(RGB=rgb)
        self.Weight = weight


class _FakePlaceholderFormat:
    def __init__(self, ph_type):
        self.Type = ph_type


class _FakeShape:
    def __init__(
        self,
        sid,
        name,
        text,
        *,
        shape_type=17,
        left=0.0,
        top=0.0,
        width=100.0,
        height=20.0,
        rotation=0.0,
        visible=-1,
        has_text_frame=True,
        ph_type=None,
        z_order=None,
        group_items=None,
        fill=None,
        line=None,
        font_size=None,
        font_name=None,
        font_color=None,
    ):
        self.Id = sid
        self.Name = name
        self.Type = shape_type
        self.Left = left
        self.Top = top
        self.Width = width
        self.Height = height
        self.Rotation = rotation
        self.Visible = visible
        self.ZOrderPosition = z_order if z_order is not None else sid
        self.HasTextFrame = -1 if has_text_frame else 0
        if has_text_frame:
            self.TextFrame = _FakeTextFrame(
                text, [_FakeRun(text, _FakeFont(font_size, font_name, font_color))]
            )
        if ph_type is not None:
            self.PlaceholderFormat = _FakePlaceholderFormat(ph_type)
        if group_items is not None:
            self.GroupItems = _FakeShapes(group_items)
        if fill is not None:
            self.Fill = fill
        if line is not None:
            self.Line = line


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


class _FakeSlides:
    def __init__(self, slides):
        self._slides = list(slides)
        self.Count = len(self._slides)

    def __call__(self, idx):
        return self._slides[idx - 1]


def _pres(*slides):
    return SimpleNamespace(
        Slides=_FakeSlides(slides),
        PageSetup=SimpleNamespace(SlideWidth=PAGE_W, SlideHeight=PAGE_H),
    )


def _app_with(pres):
    app = ppt.PptApp.__new__(ppt.PptApp)
    app._require_pres = lambda doc_id=None: pres
    return app


def _slide(*shapes):
    return SimpleNamespace(Shapes=_FakeShapes(list(shapes)))


# ------------------------------------------------------------------ 基础记录


def test_read_shapes_mixed_types_all_present():
    textbox = _FakeShape(
        2, "TextBox", "你好", left=36, top=36, z_order=1, font_size=24, font_name="Arial"
    )
    picture = _FakeShape(
        3, "Picture 1", "", shape_type=13, has_text_frame=False, left=72, top=72, z_order=2
    )
    line = _FakeShape(
        4, "Line 1", "", shape_type=9, has_text_frame=False, left=10, top=10, z_order=3
    )
    autoshape = _FakeShape(
        5, "Rect 1", "", shape_type=1, has_text_frame=False, left=100, top=100, z_order=4
    )
    app = _app_with(_pres(_slide(textbox, picture, line, autoshape)))

    recs = app.read_shapes(1)
    assert [r["shape_id"] for r in recs] == [2, 3, 4, 5]
    by_id = {r["shape_id"]: r for r in recs}

    t = by_id[2]
    assert t["name"] == "TextBox"
    assert t["shape_type"] == 17
    assert t["shape_type_name"] == "text_box"
    assert t["text"] == "你好"
    assert t["has_text_frame"] is True
    assert t["coordinate_space"] == "slide"
    assert t["coordinate_unit"] == "pt"
    assert t["is_placeholder"] is False
    assert t["placeholder_type"] is None
    assert t["parent_shape_id"] is None
    assert t["group_path"] == []
    assert t["z_order"] == 1
    assert t["visible"] is True

    assert by_id[3]["shape_type_name"] == "picture"
    assert by_id[3]["has_text_frame"] is False
    assert by_id[3]["font_size"] is None  # 无文本 → 字体全 None
    assert by_id[4]["shape_type_name"] == "line"
    assert by_id[5]["shape_type_name"] == "auto_shape"


def test_read_shapes_placeholder_fields():
    ph = _FakeShape(
        2,
        "Title 1",
        "标题",
        shape_type=14,
        ph_type=1,
        left=36,
        top=21.6,
        width=648,
        height=90,
        z_order=1,
    )
    app = _app_with(_pres(_slide(ph)))
    rec = app.read_shapes(1)[0]
    assert rec["is_placeholder"] is True
    assert rec["placeholder_type"] == 1
    assert rec["placeholder_type_name"] == "title"


def test_read_shapes_missing_slide_invalid_argument():
    app = _app_with(_pres(_slide()))
    with pytest.raises(InvalidArgumentError):
        app.read_shapes(5)


# ------------------------------------------------------------------ recursive


def test_read_shapes_recursive_returns_group_container_and_descendants():
    child = _FakeShape(302, "Child", "组内文本", left=180, top=180, z_order=14)
    grp = _FakeShape(
        300,
        "Group1",
        "",
        shape_type=6,
        has_text_frame=False,
        left=72,
        top=72,
        width=432,
        height=216,
        z_order=13,
        group_items=[child],
    )
    app = _app_with(_pres(_slide(grp)))

    # recursive=True：group 容器 + 后代
    recs = app.read_shapes(1)
    assert [r["shape_id"] for r in recs] == [300, 302]
    g, c = recs
    assert g["shape_type_name"] == "group"
    assert c["parent_shape_id"] == 300
    assert c["group_path"] == [300]
    assert c["coordinate_space"] == "slide"

    # recursive=False：仅顶层 group 容器，后代不展开
    recs = app.read_shapes(1, recursive=False)
    assert [r["shape_id"] for r in recs] == [300]


def test_read_shapes_nested_group_path_outer_to_inner():
    inner_title = _FakeShape(302, "NestedTitle", "嵌套组内标题", left=180, top=180, z_order=1)
    inner_grp = _FakeShape(
        301,
        "GroupInner",
        "",
        shape_type=6,
        has_text_frame=False,
        left=108,
        top=108,
        width=360,
        height=144,
        z_order=1,
        group_items=[inner_title],
    )
    outer_grp = _FakeShape(
        300,
        "GroupOuter",
        "",
        shape_type=6,
        has_text_frame=False,
        left=72,
        top=72,
        width=432,
        height=216,
        z_order=1,
        group_items=[inner_grp],
    )
    app = _app_with(_pres(_slide(outer_grp)))
    recs = app.read_shapes(1)
    assert [r["shape_id"] for r in recs] == [300, 301, 302]
    nested = recs[2]
    assert nested["parent_shape_id"] == 301
    assert nested["group_path"] == [300, 301]  # 外层→内层


def test_read_shapes_rotated_group_descendant_unknown():
    child = _FakeShape(302, "Child", "旋转组内文本", left=180, top=180, z_order=1)
    rotated_grp = _FakeShape(
        300,
        "RotatedGroup",
        "",
        shape_type=6,
        has_text_frame=False,
        rotation=45.0,
        left=72,
        top=72,
        width=432,
        height=216,
        z_order=1,
        group_items=[child],
    )
    app = _app_with(_pres(_slide(rotated_grp)))
    recs = app.read_shapes(1)
    assert [r["shape_id"] for r in recs] == [300, 302]
    assert recs[0]["coordinate_space"] == "slide"  # group 自身不旋转 → slide
    assert recs[1]["coordinate_space"] == "unknown"  # 旋转 group 后代


# ------------------------------------------------------------------ 严格 id


def test_read_shapes_strict_unreadable_id_raises():
    class _NoIdShape:
        Name = "no-id"
        Type = 17
        Left = Top = 0.0
        Width = Height = 10.0
        Rotation = 0.0
        Visible = -1
        ZOrderPosition = 1
        HasTextFrame = 0

    app = _app_with(_pres(_slide(_NoIdShape())))
    with pytest.raises(ComOperationError):
        app.read_shapes(1)


def test_read_shapes_no_zero_id_fallback():
    # 新 API 严格：任何 shape 的 Id 读不到 → 抛错，绝不出 shape_id=0
    class _NoIdShape:
        Name = "no-id"
        Type = 1
        Left = Top = 0.0
        Width = Height = 10.0
        Rotation = 0.0
        Visible = -1
        ZOrderPosition = 1
        HasTextFrame = 0

    app = _app_with(_pres(_slide(_NoIdShape())))
    recs = None
    with pytest.raises(ComOperationError):
        recs = app.read_shapes(1)
    assert recs is None  # 绝不以部分结果带 0 兜底返回


# ------------------------------------------------------------------ fill / line


def test_read_shapes_fill_solid_gives_color():
    sh = _FakeShape(
        2,
        "Rect",
        "",
        shape_type=1,
        has_text_frame=False,
        left=10,
        top=10,
        z_order=1,
        fill=_FakeFill(fill_type=1, rgb=_RED_RGB, transparency=0.25),
    )
    app = _app_with(_pres(_slide(sh)))
    rec = app.read_shapes(1)[0]
    assert rec["fill_color"] == "#FF0000"
    assert rec["fill_transparency"] == 0.25


def test_read_shapes_fill_non_solid_color_none():
    # gradient(3) / pattern(2) → 颜色诚实给 None（不把「不知道」编成假值）
    grad = _FakeShape(
        2,
        "Grad",
        "",
        shape_type=1,
        has_text_frame=False,
        z_order=1,
        fill=_FakeFill(fill_type=3, rgb=_RED_RGB),
    )
    pat = _FakeShape(
        3,
        "Pat",
        "",
        shape_type=1,
        has_text_frame=False,
        z_order=2,
        fill=_FakeFill(fill_type=2, rgb=_RED_RGB),
    )
    app = _app_with(_pres(_slide(grad, pat)))
    by_id = {r["shape_id"]: r for r in app.read_shapes(1)}
    assert by_id[2]["fill_color"] is None
    assert by_id[3]["fill_color"] is None


def test_read_shapes_no_fill_object_color_none():
    # 无 Fill 对象（读 Fill 抛错）→ (None, None)，诚实
    sh = _FakeShape(2, "NoFill", "", shape_type=1, has_text_frame=False, z_order=1)
    app = _app_with(_pres(_slide(sh)))
    rec = app.read_shapes(1)[0]
    assert rec["fill_color"] is None
    assert rec["fill_transparency"] is None


def test_read_shapes_line_visible_gives_color_width():
    sh = _FakeShape(
        2,
        "Line",
        "",
        shape_type=9,
        has_text_frame=False,
        left=10,
        top=10,
        z_order=1,
        line=_FakeLine(visible=-1, rgb=_PURPLE_RGB, weight=2.5),
    )
    app = _app_with(_pres(_slide(sh)))
    rec = app.read_shapes(1)[0]
    assert rec["line_color"] == "#800080"
    assert rec["line_width"] == 2.5


def test_read_shapes_line_hidden_no_color_no_width():
    sh = _FakeShape(
        2,
        "NoLine",
        "",
        shape_type=1,
        has_text_frame=False,
        z_order=1,
        line=_FakeLine(visible=0, rgb=_RED_RGB, weight=2.0),
    )
    app = _app_with(_pres(_slide(sh)))
    rec = app.read_shapes(1)[0]
    assert rec["line_color"] is None
    assert rec["line_width"] is None


# ------------------------------------------------------------------ font


def test_read_shapes_font_first_run_semantics():
    run1 = _FakeRun("Alpha", _FakeFont(24, "Arial", _RED_RGB))
    run2 = _FakeRun("Beta", _FakeFont(14, "Times New Roman", _PURPLE_RGB))
    tb = _FakeShape(2, "Tb", "AlphaBeta", left=36, top=36, z_order=1)
    tb.TextFrame = _FakeTextFrame("AlphaBeta", [run1, run2])
    app = _app_with(_pres(_slide(tb)))
    rec = app.read_shapes(1)[0]
    assert rec["font_size"] == 24.0
    assert rec["font_name"] == "Arial"
    assert rec["font_color"] == "#FF0000"


def test_read_shapes_font_empty_text_all_none():
    tb = _FakeShape(2, "Empty", "", left=36, top=36, z_order=1)
    app = _app_with(_pres(_slide(tb)))
    rec = app.read_shapes(1)[0]
    assert rec["text"] == ""
    assert rec["font_size"] is None
    assert rec["font_name"] is None
    assert rec["font_color"] is None


def test_read_shapes_visible_mixed_state_none():
    sh = _FakeShape(2, "Mixed", "", has_text_frame=False, z_order=1)
    sh.Visible = -2
    app = _app_with(_pres(_slide(sh)))
    rec = app.read_shapes(1)[0]
    assert rec["visible"] is None


# ------------------------------------------------------------------ z_order


def test_read_shapes_group_child_z_order_local_rank():
    # 探针 #6：group 子元素 ZOrderPosition 带偏移（group.Z + local）；read_shapes 的
    # z_order 语义是「所在集合内 1-based」，须按兄弟排序还原为 1..Count。
    child1 = _FakeShape(302, "c1", "", has_text_frame=False, z_order=14)
    child2 = _FakeShape(303, "c2", "", has_text_frame=False, z_order=15)
    child3 = _FakeShape(304, "c3", "", has_text_frame=False, z_order=16)
    grp = _FakeShape(
        300,
        "G",
        "",
        shape_type=6,
        has_text_frame=False,
        z_order=13,
        group_items=[child1, child2, child3],
    )
    app = _app_with(_pres(_slide(grp)))
    recs = app.read_shapes(1)
    # group 自身 z_order 在 slide.Shapes 内 = 1
    assert recs[0]["z_order"] == 1
    # 三个子元素按兄弟排序 → 本地 1..3（尽管原始 ZOrderPosition 是 14/15/16）
    assert [r["z_order"] for r in recs[1:]] == [1, 2, 3]


# ------------------------------------------------------------------ 回归


def test_read_slide_texts_regression_unchanged():
    # S1 新增 read_shapes 后，read_slide_texts 输出必须原样
    child = _FakeShape(302, "Child", "组内文本", left=180, top=180, z_order=1)
    grp = _FakeShape(
        300,
        "Group1",
        "",
        shape_type=6,
        has_text_frame=False,
        left=72,
        top=72,
        width=432,
        height=216,
        z_order=1,
        group_items=[child],
    )
    title = _FakeShape(2, "Title 1", "标题", shape_type=14, ph_type=1, left=36, top=21.6, z_order=1)
    app = _app_with(_pres(_slide(grp, title)))
    texts = app.read_slide_texts(1)
    assert [r["shape_id"] for r in texts] == [302, 2]
    assert texts[0]["parent_shape_id"] == 300
    assert texts[0]["group_path"] == [300]
