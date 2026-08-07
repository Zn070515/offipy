"""编辑定位器 + 能力护栏 + 校验器纯逻辑测试（S1 Task 3）。

用 `__new__` 不触碰 COM；直接测模块级私有定位器/校验器。覆盖：
- 递归 shape_id 定位（顶层 + 嵌套 group），返回 _LocatedShape 携带所在集合
- 找不到 shape_id → TargetNotFoundError
- 遍历途中 Id 读不到 → ComOperationError（与 read_shapes 严格一致）
- 文本操作打在无文本能力 shape（图片/线条）→ InvalidArgumentError
- 颜色/透明度/几何/字号校验 → InvalidArgumentError
- 填充/轮廓能力护栏 → InvalidArgumentError
"""

from types import SimpleNamespace

import pytest

from offipy import ppt
from offipy.exceptions import ComOperationError, InvalidArgumentError, TargetNotFoundError


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
            self.TextFrame = _FakeTextFrame(text)
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


def _pic(sid, name):
    return _FakeShape(sid, name, "", shape_type=13, has_text_frame=False)


def _line(sid, name):
    return _FakeShape(sid, name, "", shape_type=9, has_text_frame=False)


def _txt(sid, name, text=""):
    return _FakeShape(sid, name, text)


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
