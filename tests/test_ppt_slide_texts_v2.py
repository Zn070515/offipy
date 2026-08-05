"""read_slide_texts（v2 按页 per-shape）与 read_slide_summary 纯逻辑测试。

用 `__new__` 构造 PptApp 跳过 COM 初始化，注入 fake 文档树镜像
tests/fixtures/ppt/minimal_text_shapes.pptx 的结构（坐标单位磅）。
覆盖：
- per-shape SlideTextRecord 全字段（shape_id/name/text/坐标/占位符信息）
- include_empty 只包含有 TextFrame 的 shape（图片不返回）
- group 递归 + group_path/parent_shape_id + coordinate_space（非旋转 slide / 旋转 unknown）
- summary 占位符直读 / 纯文本框阅读顺序回退 / 豁免集（页码/页脚/日期 + 页码候选）
- 占位符类型完整映射 + unknown_{n} 兜底
- MsoTriState 正规化 + _shape_has_text_frame 双重兜底
"""

from types import SimpleNamespace

from offipy import ppt

PAGE_W = 720.0  # 10in
PAGE_H = 540.0  # 7.5in


class _FakeTextRange:
    def __init__(self, text):
        self.Text = text


class _FakeTextFrame:
    def __init__(self, text):
        self.TextRange = _FakeTextRange(text)


class _FakePlaceholderFormat:
    def __init__(self, ph_type):
        self.Type = ph_type


class _FakeShape:
    """COM shape 最小 fake：覆盖读全用到的属性面。"""

    def __init__(
        self,
        sid,
        name,
        text,
        *,
        shape_type=17,  # 17=文本框；14=占位符；6=group；13=图片
        left=0.0,
        top=0.0,
        width=100.0,
        height=20.0,
        rotation=0.0,
        has_text_frame=True,
        ph_type=None,
        z_order=None,
        group_items=None,
    ):
        self.Id = sid
        self.Name = name
        self.Type = shape_type
        self.Left = left
        self.Top = top
        self.Width = width
        self.Height = height
        self.Rotation = rotation
        self.ZOrderPosition = z_order if z_order is not None else sid
        self.HasTextFrame = -1 if has_text_frame else 0
        if has_text_frame:
            self.TextFrame = _FakeTextFrame(text)
        if ph_type is not None:
            self.PlaceholderFormat = _FakePlaceholderFormat(ph_type)
        if group_items is not None:
            self.GroupItems = _FakeShapes(group_items)


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


# ------------------------------------------------------------------ per-shape 记录


def test_read_slide_texts_per_shape_records():
    title = _FakeShape(
        2,
        "Title 1",
        "产品发布计划",
        shape_type=14,
        ph_type=1,
        left=36,
        top=21.6,
        width=648,
        height=90,
        z_order=1,
    )
    body = _FakeShape(
        3,
        "Content Placeholder 2",
        "第一季度\n第二季度\n第三季度",
        shape_type=14,
        ph_type=2,
        left=36,
        top=126,
        width=648,
        height=356.4,
        z_order=2,
    )
    app = _app_with(_pres(SimpleNamespace(Shapes=_FakeShapes([title, body]))))

    records = app.read_slide_texts(1)
    assert [r["shape_id"] for r in records] == [2, 3]
    t, b = records

    assert t["name"] == "Title 1"
    assert t["text"] == "产品发布计划"
    assert t["left"] == 36.0 and t["top"] == 21.6
    assert t["width"] == 648.0 and t["height"] == 90.0
    assert t["coordinate_space"] == "slide"
    assert t["coordinate_unit"] == "pt"
    assert t["is_placeholder"] is True
    assert t["placeholder_type"] == 1
    assert t["placeholder_type_name"] == "title"
    assert t["parent_shape_id"] is None
    assert t["group_path"] == []

    assert b["text"] == "第一季度\n第二季度\n第三季度"
    assert b["placeholder_type"] == 2
    assert b["placeholder_type_name"] == "body"


def test_read_slide_texts_include_empty_only_text_shapes():
    # 有文本 / 空文本文本框 + 图片（无 TextFrame）——include_empty 只应涉及前两者
    filled = _FakeShape(2, "Filled", "内容", left=36, top=36, z_order=1)
    empty = _FakeShape(3, "Empty", "", left=36, top=100, z_order=2)
    picture = _FakeShape(
        4, "Picture 1", "", shape_type=13, has_text_frame=False, left=72, top=72, z_order=3
    )
    app = _app_with(_pres(SimpleNamespace(Shapes=_FakeShapes([filled, empty, picture]))))

    # 缺省 include_empty=False：空文本文本框不返回
    assert [r["shape_id"] for r in app.read_slide_texts(1)] == [2]
    # include_empty=True：有 TextFrame 的空文本也返回；图片始终不返回
    assert [r["shape_id"] for r in app.read_slide_texts(1, include_empty=True)] == [2, 3]


def test_read_slide_texts_picture_never_returned():
    # 图片 shape 无 TextFrame：include_empty=True 也不返回（那是 read_shapes 的职责）
    picture = _FakeShape(
        4, "Picture 1", "", shape_type=13, has_text_frame=False, left=72, top=72, z_order=1
    )
    app = _app_with(_pres(SimpleNamespace(Shapes=_FakeShapes([picture]))))
    assert app.read_slide_texts(1, include_empty=True) == []


# ------------------------------------------------------------------ group 递归


def _group_slide():
    """镜像 fixture 页 3：外层 group(300) > 内层 group(301) > 文本 + 外层直接文本。"""
    nested_title = _FakeShape(
        302, "NestedTitle", "嵌套组内标题", left=180, top=180, width=288, height=36, z_order=1
    )
    nested_body = _FakeShape(
        303, "NestedBody", "嵌套组内正文", left=180, top=237.6, width=288, height=43.2, z_order=2
    )
    inner_grp = _FakeShape(
        301,
        "GroupInner",
        "",
        shape_type=6,
        left=108,
        top=108,
        width=360,
        height=144,
        z_order=1,
        group_items=[nested_title, nested_body],
    )
    outer_text = _FakeShape(
        304, "OuterText", "外层组内文本", left=93.6, top=259.2, width=288, height=36, z_order=3
    )
    outer_grp = _FakeShape(
        300,
        "GroupOuter",
        "",
        shape_type=6,
        left=72,
        top=72,
        width=432,
        height=216,
        z_order=1,
        group_items=[inner_grp, outer_text],
    )
    return SimpleNamespace(Shapes=_FakeShapes([outer_grp]))


def test_read_slide_texts_group_recursion():
    app = _app_with(_pres(_group_slide()))
    records = app.read_slide_texts(1)

    by_id = {r["shape_id"]: r for r in records}
    assert set(by_id) == {302, 303, 304}

    nested_title = by_id[302]
    assert nested_title["text"] == "嵌套组内标题"
    assert nested_title["parent_shape_id"] == 301
    assert nested_title["group_path"] == [300, 301]  # 外层→内层
    assert nested_title["coordinate_space"] == "slide"
    assert nested_title["is_placeholder"] is False
    assert nested_title["placeholder_type"] is None

    outer_text = by_id[304]
    assert outer_text["parent_shape_id"] == 300
    assert outer_text["group_path"] == [300]


def test_read_slide_texts_group_recursive_false_top_level_only():
    # recursive=False：group 自身无 TextFrame → 无记录（子元素不展开）
    app = _app_with(_pres(_group_slide()))
    assert app.read_slide_texts(1, recursive=False) == []


def test_read_slide_texts_rotated_group_unknown_coordinate_space():
    # 旋转 group 内子元素读值不可信（探针 P0-2）→ coordinate_space="unknown"
    child = _FakeShape(302, "Child", "旋转组内文本", left=180, top=180, z_order=1)
    rotated_grp = _FakeShape(
        300,
        "RotatedGroup",
        "",
        shape_type=6,
        rotation=45.0,
        left=72,
        top=72,
        width=432,
        height=216,
        z_order=1,
        group_items=[child],
    )
    app = _app_with(_pres(SimpleNamespace(Shapes=_FakeShapes([rotated_grp]))))
    records = app.read_slide_texts(1)
    assert len(records) == 1
    assert records[0]["text"] == "旋转组内文本"
    assert records[0]["group_path"] == [300]
    assert records[0]["coordinate_space"] == "unknown"


# ------------------------------------------------------------------ summary


def test_read_slide_summary_placeholder_page():
    # 镜像 fixture 页 1：title=1 / body=2 占位符直读
    title = _FakeShape(
        2,
        "Title 1",
        "产品发布计划",
        shape_type=14,
        ph_type=1,
        left=36,
        top=21.6,
        width=648,
        height=90,
        z_order=1,
    )
    body = _FakeShape(
        3,
        "Content Placeholder 2",
        "第一季度\n第二季度\n第三季度",
        shape_type=14,
        ph_type=2,
        left=36,
        top=126,
        width=648,
        height=356.4,
        z_order=2,
    )
    notes = _FakeShape(
        9, "Notes 1", "演讲者备注：本页为公司级发布计划", shape_type=14, ph_type=2, z_order=1
    )
    slide = SimpleNamespace(
        Shapes=_FakeShapes([title, body]),
        NotesPage=SimpleNamespace(Shapes=_FakeShapes([notes])),
    )
    app = _app_with(_pres(slide))
    summary = app.read_slide_summary()
    assert summary == [
        {
            "index": 1,
            "title": "产品发布计划",
            "body": "第一季度\n第二季度\n第三季度",
            "notes": "演讲者备注：本页为公司级发布计划",
        }
    ]


def test_read_slide_summary_textbox_page_fallback_and_stable():
    # 镜像 fixture 页 2：无占位符，title/body 按稳定阅读顺序回退
    title = _FakeShape(
        2, "TextBoxTitle", "纯文本框标题", left=36, top=36, width=432, height=36, z_order=1
    )
    body = _FakeShape(
        3,
        "TextBoxBody",
        "纯文本框正文第一行\n纯文本框正文第二行",
        left=36,
        top=108,
        width=432,
        height=86.4,
        z_order=2,
    )
    slide = SimpleNamespace(
        Shapes=_FakeShapes([title, body]), NotesPage=SimpleNamespace(Shapes=_FakeShapes([]))
    )
    app = _app_with(_pres(slide))
    summary = app.read_slide_summary()
    assert summary[0]["title"] == "纯文本框标题"
    assert summary[0]["body"] == "纯文本框正文第一行\n纯文本框正文第二行"


def test_read_slide_summary_title_fallback_skips_empty_text():
    # 回归：P6 真机 bug——header 背景矩形带空 TextFrame（top=0）按阅读顺序排在
    # 真标题（top=12）前，title 回退 cands 必须过滤空文本，否则 title 拿到空串。
    empty_bar = _FakeShape(
        1, "HeaderBar", "", left=0, top=0, width=720, height=10, z_order=1
    )
    title = _FakeShape(
        2, "RealTitle", "真实标题", left=36, top=12, width=400, height=30, z_order=2
    )
    body = _FakeShape(
        3, "BodyText", "正文内容", left=36, top=60, width=400, height=40, z_order=3
    )
    slide = SimpleNamespace(
        Shapes=_FakeShapes([empty_bar, title, body]),
        NotesPage=SimpleNamespace(Shapes=_FakeShapes([])),
    )
    app = _app_with(_pres(slide))
    s = app.read_slide_summary()[0]
    assert s["title"] == "真实标题"
    assert s["body"] == "正文内容"  # 空文本 shape 也不进 body


def test_read_slide_summary_textbox_page_iteration_order_stable():
    # 同一批 shape 换迭代顺序（body 在前），摘要不变：阅读顺序 key 保证确定性
    title = _FakeShape(
        2, "TextBoxTitle", "纯文本框标题", left=36, top=36, width=432, height=36, z_order=1
    )
    body = _FakeShape(
        3,
        "TextBoxBody",
        "纯文本框正文第一行\n纯文本框正文第二行",
        left=36,
        top=108,
        width=432,
        height=86.4,
        z_order=2,
    )
    rev = SimpleNamespace(
        Shapes=_FakeShapes([body, title]), NotesPage=SimpleNamespace(Shapes=_FakeShapes([]))
    )
    app = _app_with(_pres(rev))
    s = app.read_slide_summary()[0]
    assert s["title"] == "纯文本框标题"
    assert s["body"] == "纯文本框正文第一行\n纯文本框正文第二行"


def test_read_slide_summary_exempt_placeholders_and_page_number():
    # 镜像 fixture 页 4：页码/页脚/日期占位符 + 角上纯数字 → 全部豁免，不进 title/body
    title = _FakeShape(2, "TitleText", "季度财报", left=36, top=36, width=432, height=36, z_order=1)
    page_num = _FakeShape(
        3, "TextBoxPageNum", "7", left=612, top=496.8, width=43.2, height=21.6, z_order=2
    )
    footer = _FakeShape(
        400,
        "Footer 1",
        "机密",
        shape_type=14,
        ph_type=15,
        left=72,
        top=504,
        width=108,
        height=21.6,
        z_order=3,
    )
    date_ph = _FakeShape(
        401,
        "Date 1",
        "2026-08-05",
        shape_type=14,
        ph_type=16,
        left=504,
        top=504,
        width=108,
        height=21.6,
        z_order=4,
    )
    sld_num = _FakeShape(
        402,
        "SlideNumberPlaceholder 1",
        "4",
        shape_type=14,
        ph_type=13,
        left=648,
        top=504,
        width=108,
        height=21.6,
        z_order=5,
    )
    slide = SimpleNamespace(
        Shapes=_FakeShapes([title, page_num, footer, date_ph, sld_num]),
        NotesPage=SimpleNamespace(Shapes=_FakeShapes([])),
    )
    app = _app_with(_pres(slide))
    s = app.read_slide_summary()[0]
    assert s["title"] == "季度财报"
    assert s["body"] == ""
    assert "机密" not in s["body"] and "2026-08-05" not in s["body"] and "7" not in s["body"]


def test_read_slide_summary_center_title_placeholder():
    # center_title(3) 也是标题候选（standard 布局的标题页）
    ctitle = _FakeShape(
        2,
        "Title 1",
        "居中标题",
        shape_type=14,
        ph_type=3,
        left=90,
        top=216,
        width=540,
        height=100,
        z_order=1,
    )
    subtitle = _FakeShape(
        3,
        "Subtitle 2",
        "副标题",
        shape_type=14,
        ph_type=4,
        left=90,
        top=340,
        width=540,
        height=60,
        z_order=2,
    )
    slide = SimpleNamespace(
        Shapes=_FakeShapes([ctitle, subtitle]), NotesPage=SimpleNamespace(Shapes=_FakeShapes([]))
    )
    app = _app_with(_pres(slide))
    s = app.read_slide_summary()[0]
    assert s["title"] == "居中标题"
    assert s["body"] == "副标题"


# ------------------------------------------------------------------ 占位符类型映射


def test_placeholder_type_name_official_mapping():
    from offipy.models import placeholder_type_name

    expected = {
        -2: "mixed",
        1: "title",
        2: "body",
        3: "center_title",
        4: "subtitle",
        5: "vertical_title",
        6: "vertical_body",
        7: "object",
        8: "chart",
        9: "bitmap",
        10: "media_clip",
        11: "org_chart",
        12: "table",
        13: "slide_number",
        14: "header",
        15: "footer",
        16: "date",
        17: "vertical_object",
        18: "picture",
        19: "cameo",
    }
    for value, name in expected.items():
        assert placeholder_type_name(value) == name


def test_placeholder_type_name_unknown_fallback():
    from offipy.models import placeholder_type_name

    assert placeholder_type_name(99) == "unknown_99"
    assert placeholder_type_name(-5) == "unknown_-5"


def test_read_slide_texts_unknown_placeholder_type():
    # 未知占位符类型回 unknown_{n}（不回 None），不影响文本读取
    ph = _FakeShape(2, "Weird", "未知类型", shape_type=14, ph_type=99, left=36, top=36, z_order=1)
    app = _app_with(_pres(SimpleNamespace(Shapes=_FakeShapes([ph]))))
    records = app.read_slide_texts(1)
    assert records[0]["is_placeholder"] is True
    assert records[0]["placeholder_type"] == 99
    assert records[0]["placeholder_type_name"] == "unknown_99"


# ------------------------------------------------------------------ MsoTriState / TextFrame 兜底


def test_tri_state_to_bool_normalization():
    assert ppt._tri_state_to_bool(-1) is True
    assert ppt._tri_state_to_bool(1) is True
    assert ppt._tri_state_to_bool(0) is False
    assert ppt._tri_state_to_bool(-2) is None
    assert ppt._tri_state_to_bool(-3) is None


def test_shape_has_text_frame_fallbacks():
    # HasTextFrame 返回 -2（混合态）但 TextFrame 存在 → True
    mixed = SimpleNamespace(HasTextFrame=-2, TextFrame=object())
    assert ppt._shape_has_text_frame(mixed) is True

    # HasTextFrame=0（无文本）即便 TextFrame 对象存在 → False（状态优先）
    no_state = SimpleNamespace(HasTextFrame=0, TextFrame=object())
    assert ppt._shape_has_text_frame(no_state) is False

    # HasTextFrame 读取抛错但 TextFrame 可访问 → True（P2-1 双重兜底）
    class _StateBoom:
        HasTextFrame = property(lambda self: (_ for _ in ()).throw(RuntimeError("state")))

        def __init__(self):
            self.TextFrame = object()

    assert ppt._shape_has_text_frame(_StateBoom()) is True

    # HasTextFrame 与 TextFrame 都不可访问 → False
    class _BothBoom:
        HasTextFrame = property(lambda self: (_ for _ in ()).throw(RuntimeError("state")))

        @property
        def TextFrame(self):
            raise RuntimeError("frame")

    assert ppt._shape_has_text_frame(_BothBoom()) is False
