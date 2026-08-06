"""命名布局库测试：布局元数据、CSS/HTML 骨架、data-layout 引用注入。"""

import pytest

from offipy import layouts
from offipy.layouts import LAYOUTS, inject_layouts, layout_css, layout_html, referenced_layouts

# ---------------------------------------------------------------- 布局元数据


def test_registered_layout_count():
    assert len(LAYOUTS) == 11


def test_layouts_metadata():
    lst = layouts.layouts()
    names = {entry["name"] for entry in lst}
    assert "hero-title" in names
    assert "closer" in names
    assert all({"name", "title", "description"} <= set(entry) for entry in lst)


def test_expected_layout_names():
    expected = {
        "hero-title",
        "split-2col",
        "cards-3",
        "big-number",
        "quote-frame",
        "timeline",
        "comparison",
        "chart-dominant",
        "portrait-feature",
        "closer",
        "icons-row",
    }
    assert set(LAYOUTS) == expected


# ---------------------------------------------------------------- CSS / HTML


def test_layout_css_contains_common_and_specific():
    css = layout_css("hero-title")
    assert ".kicker" in css  # 公共排版
    assert ".hero { display: flex" in css  # 布局自身


def test_layout_css_uses_tokens_only():
    for name in LAYOUTS:
        css = layout_css(name)
        assert "var(--" in css, f"{name} 的 CSS 应只用 design token"
        assert "#" not in css.replace("var(--", "").replace(":", ""), f"{name} 不应出现硬编码色值"


def test_layout_html_returns_skeleton():
    html = layout_html("hero-title")
    assert "data-pptx-slide" in html
    assert 'data-layout="hero-title"' in html
    assert "<!--" in html  # 有填空注释


def test_split_2col_css_has_wrapper_display_contents():
    # 防呆：.cols 包装层透明化，Claude 包一层 wrapper 也能正确两栏
    css = layout_css("split-2col")
    assert ".split-2col > .cols { display: contents; }" in css


def test_split_2col_css_only_col_stretches():
    # #9：align-items: stretch 会把眉标/标题等非 col 的 flex 子元素拉成整页高
    # （超高文本框）。容器改 flex-start，只让 .col 用 align-self: stretch 填满。
    css = layout_css("split-2col")
    assert "align-items: flex-start" in css
    assert "align-items: stretch" not in css
    assert ".split-2col .col {" in css
    assert "align-self: stretch" in css


def test_split_2col_html_notes_flex_item_constraint():
    # HTML 模板注释提醒 flex-item 约束，防 Claude 包 wrapper 破坏布局
    html = layout_html("split-2col")
    assert "不要加 wrapper" in html
    assert "flex item" in html


def test_unknown_layout_raises():
    with pytest.raises(KeyError):
        layout_css("nope")
    with pytest.raises(KeyError):
        layout_html("nope")


# ---------------------------------------------------------------- 引用解析


def test_referenced_layouts_dedup_preserves_order():
    html = """
    <section class="slide" data-pptx-slide data-layout="hero-title"></section>
    <section class="slide" data-pptx-slide data-layout="cards-3"></section>
    <section class="slide" data-pptx-slide data-layout="hero-title"></section>
    """
    assert referenced_layouts(html) == ["hero-title", "cards-3"]


def test_referenced_layouts_none():
    assert referenced_layouts("<html><body>no layout</body></html>") == []


# ---------------------------------------------------------------- 注入


def test_inject_replaces_style_placeholder():
    html = """
    <html><head><style data-theme="mckinsey"></style><style data-layouts></style></head>
    <body><section class="slide" data-pptx-slide data-layout="hero-title"></section></body>
    </html>
    """
    out = inject_layouts(html)
    assert out.count("<style data-layouts>") == 1
    assert ".kicker" in out
    assert ".hero { display: flex" in out


def test_inject_appends_before_head_close():
    html = """
    <html><head><title>t</title></head>
    <body><section class="slide" data-pptx-slide data-layout="cards-3"></section></body>
    </html>
    """
    out = inject_layouts(html)
    assert "<style data-layouts>" in out
    assert out.index("<style data-layouts>") < out.index("</head>")


def test_inject_only_referenced_layouts():
    html = """
    <html><head></head>
    <body>
      <section class="slide" data-pptx-slide data-layout="big-number"></section>
    </body></html>
    """
    out = inject_layouts(html)
    assert ".big-number .big" in out  # 被引用的布局注入
    assert "cards-3" not in out  # 未被引用的布局不注入
    assert "comparison" not in out


def test_inject_no_references_returns_unchanged():
    html = "<html><head></head><body><p>plain</p></body></html>"
    assert inject_layouts(html) == html


def test_inject_no_head_puts_block_at_top():
    html = '<section class="slide" data-pptx-slide data-layout="closer">x</section>'
    out = inject_layouts(html)
    assert out.startswith("<style data-layouts>")


# ---------------------------------------------------------------- icons-row（M3）


def test_icons_row_layout_registered():
    assert "icons-row" in LAYOUTS


def test_icons_row_css_uses_tokens_only():
    css = LAYOUTS["icons-row"].css
    assert "var(--" in css
    assert "#" not in css  # 禁止硬编码色值
    assert "font-size: var(--body)" in css


def test_icons_row_wraps_multiple_rows():
    """icons-row 支持多行：flex-wrap + 固定 3 列/行 basis（>3 个图标自动换行）。"""
    css = LAYOUTS["icons-row"].css
    assert "flex-wrap: wrap" in css
    assert "calc((100% - 2 * var(--gap)) / 3)" in css
    assert "flex: 0 0" in css


def test_inject_icons_row_layout():
    html = (
        "<html><head><style data-layouts></style></head><body>"
        '<section class="slide icons-row" data-pptx-slide data-layout="icons-row">'
        '<div class="icon-row"><div class="icon-item">'
        '<svg class="icon" data-icon="ph:check" viewBox="0 0 256 256"></svg>'
        '<div class="label">已核</div></div></div>'
        "</section></body></html>"
    )
    out = inject_layouts(html)
    assert ".icons-row .icon-row" in out
    assert "data-layouts" in out


def test_icons_row_css_not_injected_when_unreferenced():
    html = (
        "<html><head><style data-layouts></style></head><body>"
        '<section class="slide" data-pptx-slide></section></body></html>'
    )
    out = inject_layouts(html)
    assert ".icons-row" not in out
