"""编辑定位器 + 校验器 + 破坏性编辑 op 纯逻辑测试（S1 Task 3/4）。

用 `__new__` 不触碰 COM；直接测模块级私有定位器/校验器与 PptApp 编辑 op。覆盖：
- 递归 shape_id 定位（顶层 + 嵌套 group），返回 _LocatedShape 携带所在集合
- 找不到 shape_id → TargetNotFoundError；遍历途中 Id 读不到 → ComOperationError
- 文本操作打在无文本能力 shape（图片/线条）→ InvalidArgumentError
- 颜色/透明度/几何/字号校验 → InvalidArgumentError
- 填充/轮廓能力护栏 → InvalidArgumentError
- set_shape_geometry 部分/全量更新、校验、旋转 group 后代 left/top 拒绝、绝对坐标
- set_shape_text 替换文本保留样式、无文本能力拒绝、空串允许
- set_shape_font 各属性部分更新、整范围传播多 run、校验拒绝
"""

from types import SimpleNamespace

import pytest

from offipy import ppt
from offipy.exceptions import ComOperationError, InvalidArgumentError, TargetNotFoundError


class _FakeFont:
    def __init__(self, size=None, name=None, color=None, bold=None, italic=None):
        self.Size = size
        self.Name = name
        self.Color = SimpleNamespace(RGB=color)
        self.Bold = bold
        self.Italic = italic


class _FakeRun:
    def __init__(self, text, font):
        self.Text = text
        self.Font = font


class _FakeColorProxy:
    """整范围 Font.Color 代理：.RGB 写同步到全部 run 的 Font.Color（探针 #3）。"""

    def __init__(self, fonts):
        self._fonts = list(fonts)

    @property
    def RGB(self):
        return self._fonts[0].Color.RGB

    @RGB.setter
    def RGB(self, value):
        for f in self._fonts:
            f.Color.RGB = value


class _FakeRangeFont:
    """整范围 Font 代理：属性写同步到全部 run 的 Font（模型化探针 #3 COM 传播）。

    读走第一个 run 的 Font（首 run 语义，与 spec 一致）。
    """

    def __init__(self, runs):
        self._runs = list(runs)

    def __getattr__(self, name):
        if name == "Color":
            return _FakeColorProxy([r.Font for r in self._runs])
        return getattr(self._runs[0].Font, name)

    def __setattr__(self, name, value):
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        for run in self._runs:
            setattr(run.Font, name, value)


class _FakeTextRange:
    def __init__(self, text, runs=None):
        self.Text = text
        self._runs = runs or [_FakeRun(text, _FakeFont())]

    @property
    def Font(self):
        return _FakeRangeFont(self._runs)

    def Runs(self, idx=None):
        if idx is None:
            return SimpleNamespace(Count=len(self._runs))
        return self._runs[idx - 1]


class _FakeTextFrame:
    def __init__(self, text, runs=None):
        self.TextRange = _FakeTextRange(text, runs)


class _FakePlaceholderFormat:
    def __init__(self, ph_type):
        self.Type = ph_type


class _FakeFill:
    def __init__(self, fill_type=1, rgb=None, transparency=None, visible=-1):
        self.Type = fill_type
        self.Transparency = transparency
        self.Visible = visible
        self.ForeColor = SimpleNamespace(RGB=rgb)

    def Solid(self):
        self.Type = 1


class _FakeLine:
    def __init__(self, visible=-1, rgb=None, weight=None):
        self.Visible = visible
        self.Weight = weight
        self.ForeColor = SimpleNamespace(RGB=rgb)


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
        runs=None,
        font_size=None,
        font_name=None,
        font_color=None,
        bold=None,
        italic=None,
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
            if runs is not None:
                self.TextFrame = _FakeTextFrame(text, runs)
            else:
                font = _FakeFont(font_size, font_name, font_color, bold, italic)
                self.TextFrame = _FakeTextFrame(text, [_FakeRun(text, font)])
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


class _FakeSlide:
    def __init__(self, *shapes):
        self.Shapes = _FakeShapes(list(shapes))


class _FakeSlides:
    def __init__(self, slides):
        self._slides = list(slides)
        self.Count = len(self._slides)

    def __call__(self, idx):
        return self._slides[idx - 1]


class _FakePres:
    def __init__(self, slides):
        self.Slides = _FakeSlides(slides)


def _app(slide_or_shape):
    """构造跳过 COM 初始化的 PptApp，doc_id="doc" 绑定到单页 fake pres。

    传 shape 自动包成 _FakeSlide；传 _FakeSlide 直接用。
    """
    slide = slide_or_shape if hasattr(slide_or_shape, "Shapes") else _FakeSlide(slide_or_shape)
    pres = _FakePres([slide])
    app = ppt.PptApp.__new__(ppt.PptApp)
    app._require_pres = lambda doc_id=None: pres
    return app, pres


def _pic(sid, name):
    return _FakeShape(sid, name, "", shape_type=13, has_text_frame=False)


def _line(sid, name):
    return _FakeShape(sid, name, "", shape_type=9, has_text_frame=False)


def _txt(sid, name, text="", **kw):
    return _FakeShape(sid, name, text, **kw)


def _grp(sid, name, *children, rotation=0.0):
    return _FakeShape(
        sid,
        name,
        "",
        shape_type=6,
        has_text_frame=False,
        rotation=rotation,
        group_items=list(children),
    )


class _NoCapabilityShape:
    """访问 .Fill / .Line 即抛，模拟不支持填充/轮廓的 shape。"""

    def __init__(self, sid, *, has_fill=True, has_line=True):
        self.Id = sid
        self.Name = "NoCap"
        self.Type = 1
        self.HasTextFrame = 0
        self._has_fill = has_fill
        self._has_line = has_line
        if has_fill:
            self.Fill = SimpleNamespace(Type=1)
        if has_line:
            self.Line = SimpleNamespace(Visible=-1)

    def __getattr__(self, name):
        if name == "Fill" and not self._has_fill:
            raise ComOperationError("Fill 不可用")
        if name == "Line" and not self._has_line:
            raise ComOperationError("Line 不可用")
        raise AttributeError(name)


# ------------------------------------------------------------------ 定位器


def test_locate_top_level_returns_correct_collection():
    slide = _FakeSlide(_txt(2, "A"), _txt(3, "B"), _txt(4, "C"))
    hit = ppt._find_shape_by_id(slide, 3)
    assert hit.shape.Name == "B"
    assert hit.containing_collection is slide.Shapes
    assert hit.parent_shape_id is None
    assert hit.group_path == ()
    assert hit.rotated_group_ancestor is False


def test_locate_group_child_returns_parent_group_items():
    child = _txt(302, "Child")
    grp = _grp(300, "G", child)
    slide = _FakeSlide(grp)
    hit = ppt._find_shape_by_id(slide, 302)
    assert hit.shape is child
    assert hit.containing_collection is grp.GroupItems
    assert hit.parent_shape_id == 300
    assert hit.group_path == (300,)
    assert hit.rotated_group_ancestor is False


def test_locate_nested_group_path_outer_to_inner():
    inner_title = _txt(302, "NestedTitle")
    inner = _grp(301, "Inner", inner_title)
    outer = _grp(300, "Outer", inner)
    slide = _FakeSlide(outer)
    hit = ppt._find_shape_by_id(slide, 302)
    assert hit.shape is inner_title
    assert hit.containing_collection is inner.GroupItems
    assert hit.parent_shape_id == 301
    assert hit.group_path == (300, 301)


def test_locate_rotated_group_descendant_flags_unknown():
    child = _txt(302, "Child")
    rotated = _grp(300, "Rot", child, rotation=45.0)
    slide = _FakeSlide(rotated)
    hit = ppt._find_shape_by_id(slide, 302)
    assert hit.rotated_group_ancestor is True


def test_locate_wrong_id_raises_target_not_found():
    slide = _FakeSlide(_txt(2, "A"), _grp(300, "G", _txt(302, "C")))
    with pytest.raises(TargetNotFoundError):
        ppt._find_shape_by_id(slide, 999)


def test_locate_unreadable_id_raises_com_error():
    class _NoIdShape:
        Name = "no-id"
        Type = 1
        HasTextFrame = 0

    slide = _FakeSlide(_NoIdShape())
    with pytest.raises(ComOperationError):
        ppt._find_shape_by_id(slide, 2)


# ------------------------------------------------------------------ 文本能力护栏


def test_require_text_frame_passes_for_textbox():
    ppt._require_text_frame(_txt(2, "T", "hi"), "测试")  # 不抛


def test_require_text_frame_rejects_picture():
    with pytest.raises(InvalidArgumentError):
        ppt._require_text_frame(_pic(3, "Pic"), "设置文本")


def test_require_text_frame_rejects_line():
    with pytest.raises(InvalidArgumentError):
        ppt._require_text_frame(_line(4, "Ln"), "设置文本")


# ------------------------------------------------------------------ 填充 / 轮廓能力护栏


def test_require_fill_capability_rejects_no_fill():
    with pytest.raises(InvalidArgumentError):
        ppt._require_fill_capability(_NoCapabilityShape(2, has_fill=False), "设置填充")


def test_require_fill_capability_passes_with_fill():
    ppt._require_fill_capability(_FakeShape(2, "S", "", fill=SimpleNamespace(Type=1)), "设置填充")


def test_require_line_capability_rejects_no_line():
    with pytest.raises(InvalidArgumentError):
        ppt._require_line_capability(_NoCapabilityShape(2, has_line=False), "设置轮廓")


def test_require_line_capability_passes_with_line():
    sh = _FakeShape(2, "S", "", line=SimpleNamespace(Visible=-1))
    ppt._require_line_capability(sh, "设置轮廓")


# ------------------------------------------------------------------ 校验器


def test_validate_hex_color_accepts_and_normalizes():
    assert ppt._validate_hex_color("#ff0000") == "#FF0000"
    assert ppt._validate_hex_color("  #a0b1c2  ") == "#A0B1C2"


@pytest.mark.parametrize(
    "bad",
    ["FF0000", "#GG0000", "#FFF", "#12345", "#1234567", "", 123, None],
)
def test_validate_hex_color_rejects(bad):
    with pytest.raises(InvalidArgumentError):
        ppt._validate_hex_color(bad)


def test_validate_fraction_0_1_accepts_bounds():
    assert ppt._validate_fraction_0_1(0.0) == 0.0
    assert ppt._validate_fraction_0_1(1.0) == 1.0
    assert ppt._validate_fraction_0_1(0.25) == 0.25


@pytest.mark.parametrize("bad", [1.5, -0.1, float("nan"), float("inf"), "abc", None])
def test_validate_fraction_0_1_rejects(bad):
    with pytest.raises(InvalidArgumentError):
        ppt._validate_fraction_0_1(bad)


def test_validate_positive_float_accepts():
    assert ppt._validate_positive_float(10, "width") == 10.0
    assert ppt._validate_positive_float(0.5, "height") == 0.5


@pytest.mark.parametrize("bad", [0, -1, float("nan"), float("inf"), "x", None])
def test_validate_positive_float_rejects(bad):
    with pytest.raises(InvalidArgumentError):
        ppt._validate_positive_float(bad, "width")


def test_validate_finite_float_accepts():
    assert ppt._validate_finite_float(45.0, "rotation") == 45.0
    assert ppt._validate_finite_float(0.0, "left") == 0.0


@pytest.mark.parametrize("bad", [float("nan"), float("inf"), "x", None])
def test_validate_finite_float_rejects(bad):
    with pytest.raises(InvalidArgumentError):
        ppt._validate_finite_float(bad, "left")


# ------------------------------------------------------------------ 颜色换算


def test_rgb_to_com_packs_bgr():
    # #FF0000 → COM BGR 低字节 R：0x0000FF
    assert ppt._rgb_to_com("#FF0000") == 0xFF
    # #800080 → r=0x80, g=0x00, b=0x80
    assert ppt._rgb_to_com("#800080") == 0x800080
    # 往返：_rgb_to_hex(_rgb_to_com(c)) == c
    for c in ("#000000", "#FFFFFF", "#A0B1C2"):
        assert ppt._rgb_to_hex(ppt._rgb_to_com(c)) == c


# ------------------------------------------------------------------ set_shape_geometry


def test_set_shape_geometry_partial_update():
    sh = _txt(2, "T", "hi", left=10, top=20, width=100, height=50, rotation=0)
    app, pres = _app(sh)
    app.set_shape_geometry(1, 2, left=15, top=25, doc_id="doc")
    got = pres.Slides(1).Shapes(1)
    assert got.Left == 15
    assert got.Top == 25
    assert got.Width == 100  # 未传属性不动
    assert got.Height == 50
    assert got.Rotation == 0


def test_set_shape_geometry_full_update():
    sh = _txt(2, "T", "hi")
    app, pres = _app(sh)
    app.set_shape_geometry(1, 2, left=1, top=2, width=300, height=200, rotation=90, doc_id="doc")
    got = pres.Slides(1).Shapes(1)
    assert (got.Left, got.Top, got.Width, got.Height, got.Rotation) == (1, 2, 300, 200, 90)


def test_set_shape_geometry_no_args_rejects():
    app, _ = _app(_txt(2, "T", "hi"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_geometry(1, 2, doc_id="doc")


@pytest.mark.parametrize("kw", [{"width": 0}, {"height": -5}, {"width": float("nan")}])
def test_set_shape_geometry_invalid_value_rejects(kw):
    app, _ = _app(_txt(2, "T", "hi"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_geometry(1, 2, doc_id="doc", **kw)


def test_set_shape_geometry_rotated_group_child_rejects_left_top():
    child = _txt(302, "Child", "x")
    grp = _grp(300, "Rot", child, rotation=45.0)
    app, _ = _app(_FakeSlide(grp))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_geometry(1, 302, left=10, doc_id="doc")
    with pytest.raises(InvalidArgumentError):
        app.set_shape_geometry(1, 302, top=10, doc_id="doc")
    # width/height/rotation 仍允许
    app.set_shape_geometry(1, 302, width=200, height=100, rotation=15, doc_id="doc")
    assert child.Width == 200
    assert child.Height == 100
    assert child.Rotation == 15


def test_set_shape_geometry_group_child_absolute_coords():
    child = _txt(302, "Child", "x", left=400, top=200)
    grp = _grp(300, "G", child)
    app, _ = _app(_FakeSlide(grp))
    app.set_shape_geometry(1, 302, left=420, top=210, doc_id="doc")
    assert child.Left == 420
    assert child.Top == 210


def test_set_shape_geometry_wrong_id_raises_target_not_found():
    app, _ = _app(_txt(2, "T", "hi"))
    with pytest.raises(TargetNotFoundError):
        app.set_shape_geometry(1, 999, left=1, doc_id="doc")


# ------------------------------------------------------------------ set_shape_text


def test_set_shape_text_replaces_and_preserves_style():
    sh = _txt(2, "T", "hi", font_name="Georgia", font_size=28, font_color=0x800080, bold=-1)
    app, pres = _app(sh)
    app.set_shape_text(1, 2, "Replaced", doc_id="doc")
    tr = pres.Slides(1).Shapes(1).TextFrame.TextRange
    assert tr.Text == "Replaced"
    assert tr.Runs(1).Font.Name == "Georgia"
    assert tr.Runs(1).Font.Size == 28
    assert tr.Runs(1).Font.Color.RGB == 0x800080
    assert tr.Runs(1).Font.Bold == -1


def test_set_shape_text_empty_allowed():
    app, pres = _app(_txt(2, "T", "hi"))
    app.set_shape_text(1, 2, "", doc_id="doc")
    assert pres.Slides(1).Shapes(1).TextFrame.TextRange.Text == ""


def test_set_shape_text_on_picture_rejects():
    app, _ = _app(_pic(3, "Pic"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_text(1, 3, "x", doc_id="doc")


def test_set_shape_text_on_line_rejects():
    app, _ = _app(_line(4, "Ln"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_text(1, 4, "x", doc_id="doc")


def test_set_shape_text_non_string_rejects():
    app, _ = _app(_txt(2, "T", "hi"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_text(1, 2, 123, doc_id="doc")


# ------------------------------------------------------------------ set_shape_font


def test_set_shape_font_partial_each_option():
    sh = _txt(2, "T", "hi")
    app, _ = _app(sh)
    app.set_shape_font(1, 2, size=30, doc_id="doc")
    assert sh.TextFrame.TextRange.Font.Size == 30
    app.set_shape_font(1, 2, font_name="Arial", doc_id="doc")
    assert sh.TextFrame.TextRange.Font.Name == "Arial"
    app.set_shape_font(1, 2, bold=True, doc_id="doc")
    assert sh.TextFrame.TextRange.Font.Bold == -1
    app.set_shape_font(1, 2, italic=True, doc_id="doc")
    assert sh.TextFrame.TextRange.Font.Italic == -1
    app.set_shape_font(1, 2, color="#FF0000", doc_id="doc")
    assert sh.TextFrame.TextRange.Font.Color.RGB == 0xFF


def test_set_shape_font_combined_update():
    sh = _txt(2, "T", "hi")
    app, _ = _app(sh)
    app.set_shape_font(1, 2, size=18, bold=True, color="#800080", doc_id="doc")
    f = sh.TextFrame.TextRange.Font
    assert f.Size == 18
    assert f.Bold == -1
    assert f.Color.RGB == 0x800080


def test_set_shape_font_whole_range_propagates_all_runs():
    run1 = _FakeRun("Alpha", _FakeFont(24, "Arial"))
    run2 = _FakeRun("Beta", _FakeFont(14, "Times"))
    sh = _FakeShape(2, "T", "AlphaBeta", runs=[run1, run2])
    app, _ = _app(sh)
    app.set_shape_font(1, 2, size=30, doc_id="doc")
    assert sh.TextFrame.TextRange.Runs(1).Font.Size == 30
    assert sh.TextFrame.TextRange.Runs(2).Font.Size == 30


def test_set_shape_font_no_args_rejects():
    app, _ = _app(_txt(2, "T", "hi"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_font(1, 2, doc_id="doc")


@pytest.mark.parametrize("kw", [{"size": 0}, {"size": -1}, {"size": float("nan")}])
def test_set_shape_font_invalid_size_rejects(kw):
    app, _ = _app(_txt(2, "T", "hi"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_font(1, 2, doc_id="doc", **kw)


def test_set_shape_font_invalid_color_rejects():
    app, _ = _app(_txt(2, "T", "hi"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_font(1, 2, color="red", doc_id="doc")


def test_set_shape_font_non_bool_bold_rejects():
    app, _ = _app(_txt(2, "T", "hi"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_font(1, 2, bold=1, doc_id="doc")


def test_set_shape_font_on_picture_rejects():
    app, _ = _app(_pic(3, "Pic"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_font(1, 3, size=30, doc_id="doc")


# ------------------------------------------------------------------ set_shape_fill


def _filled(sid=2, name="S", **kw):
    return _FakeShape(sid, name, "", fill=_FakeFill(**kw))


def test_set_shape_fill_color_forces_solid_and_shows():
    fill = _FakeFill(fill_type=3, rgb=0x0, transparency=0.5)
    sh = _FakeShape(2, "S", "", fill=fill)
    app, _ = _app(sh)
    app.set_shape_fill(1, 2, color="#FF0000", doc_id="doc")
    assert fill.Type == 1  # Solid() 强制 solid
    assert fill.Visible == -1
    assert fill.ForeColor.RGB == 0xFF
    assert fill.Transparency == 0.5  # 未传属性保留


def test_set_shape_fill_transparency_only_keeps_color():
    fill = _FakeFill(fill_type=1, rgb=0x80, transparency=0.0)
    sh = _FakeShape(2, "S", "", fill=fill)
    app, _ = _app(sh)
    app.set_shape_fill(1, 2, transparency=0.25, doc_id="doc")
    assert fill.Transparency == 0.25
    assert fill.ForeColor.RGB == 0x80  # 颜色保留


def test_set_shape_fill_both():
    fill = _FakeFill()
    sh = _FakeShape(2, "S", "", fill=fill)
    app, _ = _app(sh)
    app.set_shape_fill(1, 2, color="#800080", transparency=0.5, doc_id="doc")
    assert fill.Type == 1
    assert fill.Visible == -1
    assert fill.ForeColor.RGB == 0x800080
    assert fill.Transparency == 0.5


def test_set_shape_fill_no_args_clears():
    fill = _FakeFill(rgb=0xFF)
    sh = _FakeShape(2, "S", "", fill=fill)
    app, _ = _app(sh)
    app.set_shape_fill(1, 2, doc_id="doc")
    assert fill.Visible == 0  # 清除填充


def test_set_shape_fill_on_line_rejects():
    sh = _line(4, "Ln")  # 无 Fill → 能力护栏
    app, _ = _app(sh)
    with pytest.raises(InvalidArgumentError):
        app.set_shape_fill(1, 4, color="#FF0000", doc_id="doc")


def test_set_shape_fill_invalid_color_rejects():
    app, _ = _app(_filled())
    with pytest.raises(InvalidArgumentError):
        app.set_shape_fill(1, 2, color="red", doc_id="doc")


def test_set_shape_fill_invalid_transparency_rejects():
    app, _ = _app(_filled())
    with pytest.raises(InvalidArgumentError):
        app.set_shape_fill(1, 2, transparency=1.5, doc_id="doc")


# ------------------------------------------------------------------ set_shape_outline


def _lined(sid=2, name="S", **kw):
    return _FakeShape(sid, name, "", line=_FakeLine(**kw))


def test_set_shape_outline_color_only():
    line = _FakeLine()
    sh = _FakeShape(2, "S", "", line=line)
    app, _ = _app(sh)
    app.set_shape_outline(1, 2, color="#800080", doc_id="doc")
    assert line.ForeColor.RGB == 0x800080


def test_set_shape_outline_width_only():
    line = _FakeLine(weight=1.0)
    sh = _FakeShape(2, "S", "", line=line)
    app, _ = _app(sh)
    app.set_shape_outline(1, 2, width=2.5, doc_id="doc")
    assert line.Weight == 2.5


def test_set_shape_outline_visible_true():
    line = _FakeLine(visible=0)
    sh = _FakeShape(2, "S", "", line=line)
    app, _ = _app(sh)
    app.set_shape_outline(1, 2, visible=True, doc_id="doc")
    assert line.Visible == -1


def test_set_shape_outline_visible_false_hides():
    line = _FakeLine(visible=-1)
    sh = _FakeShape(2, "S", "", line=line)
    app, _ = _app(sh)
    app.set_shape_outline(1, 2, visible=False, doc_id="doc")
    assert line.Visible == 0


def test_set_shape_outline_combined():
    line = _FakeLine()
    sh = _FakeShape(2, "S", "", line=line)
    app, _ = _app(sh)
    app.set_shape_outline(1, 2, color="#FF0000", width=3.0, visible=False, doc_id="doc")
    assert line.ForeColor.RGB == 0xFF
    assert line.Weight == 3.0
    assert line.Visible == 0  # 显式 visible 最后生效


def test_set_shape_outline_no_args_rejects():
    app, _ = _app(_lined())
    with pytest.raises(InvalidArgumentError):
        app.set_shape_outline(1, 2, doc_id="doc")


def test_set_shape_outline_invalid_width_rejects():
    app, _ = _app(_lined())
    with pytest.raises(InvalidArgumentError):
        app.set_shape_outline(1, 2, width=0, doc_id="doc")


def test_set_shape_outline_invalid_color_rejects():
    app, _ = _app(_lined())
    with pytest.raises(InvalidArgumentError):
        app.set_shape_outline(1, 2, color="#GG0000", doc_id="doc")


def test_set_shape_outline_on_no_line_shape_rejects():
    app, _ = _app(_txt(2, "T", "hi"))  # 无 Line → 能力护栏
    with pytest.raises(InvalidArgumentError):
        app.set_shape_outline(1, 2, width=1.0, doc_id="doc")


# ------------------------------------------------------------------ set_shape_visible


def test_set_shape_visible_true_and_false():
    sh = _txt(2, "T", "hi", visible=0)
    app, pres = _app(sh)
    app.set_shape_visible(1, 2, True, doc_id="doc")
    assert pres.Slides(1).Shapes(1).Visible == -1
    app.set_shape_visible(1, 2, False, doc_id="doc")
    assert pres.Slides(1).Shapes(1).Visible == 0


@pytest.mark.parametrize("bad", [1, 0, None, "yes"])
def test_set_shape_visible_non_bool_rejects(bad):
    app, _ = _app(_txt(2, "T", "hi"))
    with pytest.raises(InvalidArgumentError):
        app.set_shape_visible(1, 2, bad, doc_id="doc")
