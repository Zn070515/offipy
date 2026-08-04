"""命名布局库测试：布局元数据、CSS/HTML 骨架、data-layout 引用注入。"""

import pytest

from offipy import layouts
from offipy.layouts import LAYOUTS, inject_layouts, layout_css, layout_html, referenced_layouts

# ---------------------------------------------------------------- 布局元数据


def test_registered_layout_count():
    assert len(LAYOUTS) == 10


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
