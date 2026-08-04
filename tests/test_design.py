"""设计系统测试：token 完整性、主题 CSS 生成、占位注入。"""

from offipy import design


def test_three_builtin_themes():
    names = design.themes()
    assert {n["name"] for n in names} == {"mckinsey", "academic", "dark-tech"}


def test_all_themes_have_required_tokens():
    for theme in design.THEMES.values():
        for key in (
            "--bg",
            "--surface",
            "--ink",
            "--muted",
            "--accent",
            "--accent-soft",
            "--divider",
            "--font",
            "--font-serif",
            "--font-sans",
        ):
            assert key in theme.base_vars, f"{theme.name} 缺 {key}"


def test_theme_css_contains_root_and_variant():
    css = design.theme_css("mckinsey")
    assert ":root {" in css
    assert "--bg: #FFFFFF;" in css
    assert "--accent: #2251FF;" in css
    assert ".slide.dark {" in css
    assert "--bg: #051C2C;" in css


def test_theme_css_contains_slide_base_styles():
    for name in design.THEMES:
        css = design.theme_css(name)
        assert "width: 1920px; height: 1080px" in css
        assert ".slide {" in css
        assert "background: var(--bg)" in css
        assert "padding: var(--pad)" in css


def test_light_default_theme_has_light_variant():
    css = design.theme_css("dark-tech")
    assert ".slide.light {" in css
    assert "--bg: #F8FAFC;" in css


def test_unknown_theme_raises():
    import pytest

    with pytest.raises(KeyError):
        design.theme_css("nonexistent")


def test_base_tokens_shared_across_themes():
    # 字号 / 间距刻度是公共的，主题只覆盖颜色字体圆角
    css = design.theme_css("academic")
    assert "--title: 52px;" in css
    assert "--body: 24px;" in css
    assert "--pad: 96px;" in css


def test_root_block_includes_shared_tokens():
    # 公共刻度必须在 :root 里（布局 CSS 的 var(--title) 依赖），不能只在 variant
    css = design.theme_css("mckinsey")
    root_block = css.split(":root {", 1)[1].split("}", 1)[0]
    assert "--title: 52px;" in root_block
    assert "--body: 24px;" in root_block
    assert "--pad: 96px;" in root_block
    assert "--gap: 24px;" in root_block


def test_inject_theme_replaces_placeholder():
    html = '<html><head><style data-theme="mckinsey"></style></head><body>deck</body></html>'
    out = design.inject_theme(html, "mckinsey")
    assert '<style data-theme="mckinsey"></style>' not in out  # 空占位被填掉
    assert 'data-theme="mckinsey"' in out  # 标记保留可追溯
    assert ":root {" in out
    assert "--accent: #2251FF;" in out
    # 注入后仍是合法可用的单 <head>
    assert out.count("<head>") == 1


def test_inject_theme_appends_when_no_placeholder():
    html = "<html><head><title>t</title></head><body>deck</body></html>"
    out = design.inject_theme(html, "dark-tech")
    assert ":root {" in out
    # 注入块出现在 </head> 之前
    assert out.index(":root {") < out.index("</head>")


def test_inject_theme_unknown_raises():
    import pytest

    with pytest.raises(KeyError):
        design.inject_theme("<html><head></head><body></body></html>", "nope")
