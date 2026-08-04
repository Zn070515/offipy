"""自动选型测试：内容信号 → 布局/主题推荐，显式 data-layout 优先。"""

from offipy import autopick
from offipy.autopick import pick, pick_file, pick_layouts, pick_theme


def _section(content):
    return f'<section class="slide" data-pptx-slide>\n{content}\n</section>'


def _deck(*extra):
    """index 1 固定封面，被测页放 index 2 起。"""
    cover = _section('<h1 class="title">封面</h1>')
    body = "".join(extra)
    return "<html><head></head><body>" + cover + body + "</body></html>"


def _picks(html):
    return {p.index: p for p in pick_layouts(html)}


# ---------------------------------------------------------------- 单页布局


def test_first_slide_always_hero():
    html = _deck()
    assert _picks(html)[1].layout == "hero-title"


def test_big_number_page():
    body = '<div class="kicker">KPI</div><div class="big">+18%</div>'
    body += '<h2 class="title">增长</h2>'
    html = _deck(_section(body))
    assert _picks(html)[2].layout == "big-number"


def test_comparison_page():
    html = _deck(_section('<div class="cmp-row"><div class="cmp-cell">成本</div></div>'))
    assert _picks(html)[2].layout == "comparison"


def test_quote_page():
    body = '<blockquote class="quote">一句话</blockquote>'
    body += '<div class="attribution">出处</div>'
    html = _deck(_section(body))
    assert _picks(html)[2].layout == "quote-frame"


def test_closer_page():
    html = _deck(_section('<h1 class="title">Thank You</h1><div class="contacts">me@x.com</div>'))
    assert _picks(html)[2].layout == "closer"


def test_timeline_page():
    html = _deck(_section('<div class="tl-item"><div class="tl-date">2025</div></div>'))
    assert _picks(html)[2].layout == "timeline"


def test_chart_page():
    body = '<h2 class="title">洞察</h2><div class="chart">'
    body += '<div class="bar"></div></div>'
    html = _deck(_section(body))
    assert _picks(html)[2].layout == "chart-dominant"


def test_cards_page():
    body = '<div class="cards"><div class="card">'
    body += '<div class="num">01</div></div></div>'
    html = _deck(_section(body))
    assert _picks(html)[2].layout == "cards-3"


def test_split_2col_page():
    html = _deck(_section('<div class="col">A</div><div class="col">B</div>'))
    assert _picks(html)[2].layout == "split-2col"


def test_plain_content_defaults_hero():
    html = _deck(_section('<h2 class="title">普通章节</h2><p>正文</p>'))
    assert _picks(html)[2].layout == "hero-title"


# ---------------------------------------------------------------- 显式 data-layout


def test_explicit_layout_wins():
    target = _section('<h1 class="title">Thank You</h1>').replace(
        "data-pptx-slide", 'data-pptx-slide data-layout="cards-3"', 1
    )
    html = _deck(target)
    page = _picks(html)[2]
    assert page.layout == "cards-3"
    assert page.explicit is True


def test_pick_layouts_has_reasons():
    html = _deck(_section('<div class="big">95%</div>'))
    page = _picks(html)[2]
    assert page.reason  # 每页都有理由
    assert "signals" in page.to_dict()


# ---------------------------------------------------------------- 主题


def test_theme_tech_keywords():
    html = _deck(_section("<p>cloud data engine</p>"))
    theme, _ = pick_theme(html)
    assert theme == "dark-tech"


def test_theme_academic_keywords():
    html = _deck(_section("<p>research paper study</p>"))
    theme, _ = pick_theme(html)
    assert theme == "academic"


def test_theme_default_mckinsey():
    html = _deck(_section("<p>我们做了一件事</p>"))
    theme, _ = pick_theme(html)
    assert theme == "mckinsey"


def test_theme_ai_not_substring():
    # "ai" 不应误匹配 "daily"（词边界）
    html = _deck(_section("<p>daily routine</p>"))
    theme, _ = pick_theme(html)
    assert theme == "mckinsey"


# ---------------------------------------------------------------- 总入口


def test_pick_returns_complete_deckpick():
    html = _deck(_section('<div class="big">+18%</div>'))
    result = pick(html)
    assert result.theme in autopick.available_themes()
    assert len(result.slides) == 2
    assert result.slides[0].index == 1
    assert result.slides[1].layout == "big-number"


def test_pick_markdown():
    html = _deck(_section('<div class="big">+18%</div>'))
    md = pick(html).markdown()
    assert "自动选型推荐" in md
    assert "big-number" in md


def test_pick_file(tmp_path):
    html = tmp_path / "deck.html"
    html.write_text(_deck(), encoding="utf-8")
    result = pick_file(str(html))
    assert result.theme == "mckinsey"
    assert len(result.slides) == 1


def test_available_lists():
    assert "mckinsey" in autopick.available_themes()
    assert "hero-title" in autopick.available_layouts()
