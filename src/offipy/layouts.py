"""命名布局库：11 种可复用页面组件 + 按引用注入。

Claude 写 deck 时给 slide 打 `data-layout="<name>"` 引用布局（如
`<section class="slide hero" data-pptx-slide data-layout="hero-title">`）。
inject_layouts 检测引用，把对应布局 CSS（+ 公共排版）注入 <head>。

所有布局 CSS 只用设计 token（var(--bg) / var(--accent) / var(--title)…，
见 design.py），因此任何主题下都成立。布局模板 HTML 供 Claude 复制后填内容，
用 HTML 注释标注填空位置。

布局子元素就是 flex item：section 里除了注释，不要多加 wrapper 层（会破坏
flex 布局）。split-2col 对 .cols 包装做了 display:contents 透明化兜底，但
眉标/标题等其它子元素仍须直接作 section 子元素。

布局来源：docs/ppt_design_research.md §5.2（agentara/skills 的 10 种命名布局）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class Layout:
    name: str
    title: str
    description: str
    css: str
    html: str


# 公共排版：多个布局共享的组件 class（kicker / title / subtitle / rule）
_COMMON_CSS = """
/* ===== 公共排版（所有布局共享） ===== */
.kicker {
  font-size: var(--kicker);
  letter-spacing: 0.18em;
  text-transform: uppercase;
  color: var(--accent);
  margin-bottom: 24px;
  font-family: var(--font-sans);
}
.title { font-size: var(--title); font-weight: 700; line-height: 1.15; margin-bottom: 32px; }
.subtitle { font-size: var(--body); color: var(--muted); max-width: 60ch; line-height: 1.6; }
.rule { font-size: var(--body); color: var(--muted); line-height: 1.6; }
.footer {
  position: absolute; bottom: 56px; left: var(--pad); right: var(--pad);
  font-size: var(--caption); color: var(--muted);
  display: flex; justify-content: space-between;
}
"""

LAYOUTS: dict[str, Layout] = {}


def _register(layout: Layout) -> Layout:
    LAYOUTS[layout.name] = layout
    return layout


_register(
    Layout(
        name="hero-title",
        title="封面 / 章节页",
        description="全页居中：眉标 + 大标题 + 副标题。留白最大，一页一观点。",
        css="""
.hero { display: flex; flex-direction: column; justify-content: center; align-items: flex-start; }
.hero .title { font-size: var(--title); }
.hero .subtitle { margin-top: 8px; }
""",
        html="""<section class="slide hero" data-pptx-slide data-layout="hero-title">
  <div class="kicker"><!-- 眉标：如 Quarterly Review --></div>
  <h1 class="title"><!-- 主标题，直接给结论 --></h1>
  <p class="subtitle"><!-- 一句话副标题 --></p>
</section>""",
    )
)

_register(
    Layout(
        name="split-2col",
        title="左右两栏",
        description="A/B 两栏并排对比内容，中间留 gap。",
        css="""
.split-2col { display: flex; gap: var(--gap); align-items: flex-start; }
/* 防呆：.cols 包装层透明化（display: contents），包一层 wrapper 也能正确两栏。
   但眉标/标题等其它子元素仍直接作 section 子元素（它们是 flex item）。
   只让 .col 用 align-self: stretch 填满整页，眉标/标题保持自然高度——
   容器级拉伸会把非 col 的 flex 子元素拉成超高文本框。 */
.split-2col > .cols { display: contents; }
.split-2col .col {
  flex: 1; align-self: stretch; background: var(--surface); border-radius: var(--radius);
  padding: 40px; border-top: 4px solid var(--accent);
}
.split-2col .col-title { font-size: var(--col-title); font-weight: 700; margin-bottom: 20px; }
.split-2col .col p { font-size: var(--body); color: var(--ink); line-height: 1.7; }
""",
        html="""<section class="slide split-2col" data-pptx-slide data-layout="split-2col">
  <!-- 布局子元素即 flex item，不要加 wrapper：.col 必须是 section 直接子元素；
       想加眉标/标题，也直接作 section 子元素（放在两栏之前），别包进 .cols。 -->
  <div class="col"><h2 class="col-title"><!-- 左栏标题 --></h2><p><!-- 左栏正文 --></p></div>
  <div class="col"><h2 class="col-title"><!-- 右栏标题 --></h2><p><!-- 右栏正文 --></p></div>
</section>""",
    )
)

_register(
    Layout(
        name="cards-3",
        title="三卡片",
        description="眉标 + 标题 + 三张并列卡片，每卡带序号与要点。",
        css="""
.cards-3 .cards { display: flex; gap: var(--gap); margin-top: 48px; }
.cards-3 .card {
  flex: 1; background: var(--surface); border-radius: var(--radius); padding: 40px;
  border-left: 6px solid var(--accent);
}
.cards-3 .num {
  font-size: var(--num); color: var(--accent); font-weight: 700;
  margin-bottom: 14px;
}
.cards-3 .txt { font-size: var(--body); color: var(--ink); line-height: 1.6; }
""",
        html="""<section class="slide cards-3" data-pptx-slide data-layout="cards-3">
  <div class="kicker"><!-- 眉标 --></div>
  <h2 class="title"><!-- 标题 --></h2>
  <div class="cards">
    <div class="card"><div class="num">01</div><div class="txt"><!-- 要点 --></div></div>
    <div class="card"><div class="num">02</div><div class="txt"><!-- 要点 --></div></div>
    <div class="card"><div class="num">03</div><div class="txt"><!-- 要点 --></div></div>
  </div>
</section>""",
    )
)

_register(
    Layout(
        name="big-number",
        title="大数字页",
        description="超大数字 + 一句结论，highlight the one。",
        css="""
.big-number .big {
  font-size: var(--display); font-weight: 800; color: var(--accent);
  line-height: 1; margin: 40px 0 24px;
}
.big-number .rule { margin-top: 16px; }
""",
        html="""<section class="slide big-number" data-pptx-slide data-layout="big-number">
  <div class="kicker"><!-- 眉标 --></div>
  <div class="big"><!-- 大数字：+18% --></div>
  <h2 class="title"><!-- 结论 --></h2>
  <p class="rule"><!-- 补充说明 --></p>
</section>""",
    )
)

_register(
    Layout(
        name="quote-frame",
        title="引言页",
        description="大引言 + 出处，留白为主，用于过渡/总结。",
        css="""
.quote-frame {
  display: flex; flex-direction: column; justify-content: center; align-items: flex-start;
}
.quote-frame .quote {
  font-family: var(--font-serif); font-size: var(--quote); font-weight: 600;
  line-height: 1.4; color: var(--ink); max-width: 26ch; margin-bottom: 32px;
}
.quote-frame .quote::before {
  content: "“"; color: var(--accent);
  font-size: var(--quote-mark); line-height: 0;
}
.quote-frame .attribution {
  font-size: var(--body); color: var(--muted); font-family: var(--font-sans);
}
""",
        html="""<section class="slide quote-frame" data-pptx-slide data-layout="quote-frame">
  <blockquote class="quote"><!-- 引言正文 --></blockquote>
  <div class="attribution"><!-- — 出处 --></div>
</section>""",
    )
)

_register(
    Layout(
        name="timeline",
        title="时间线",
        description="竖排时间线：日期 + 要点，展示进程/里程碑。",
        css="""
.timeline .tl-item {
  display: flex; gap: 32px; align-items: baseline;
  padding: 16px 0; border-bottom: 1px solid var(--divider);
}
.timeline .tl-date {
  font-family: var(--font-sans); font-size: var(--tl-date); font-weight: 700;
  color: var(--accent); min-width: 160px;
}
.timeline .tl-txt { font-size: var(--body); color: var(--ink); line-height: 1.6; }
""",
        html="""<section class="slide timeline" data-pptx-slide data-layout="timeline">
  <div class="kicker"><!-- 眉标 --></div>
  <h2 class="title"><!-- 标题 --></h2>
  <div class="tl-item">
    <div class="tl-date"><!-- 时间 --></div>
    <div class="tl-txt"><!-- 要点 --></div>
  </div>
  <div class="tl-item">
    <div class="tl-date"><!-- 时间 --></div>
    <div class="tl-txt"><!-- 要点 --></div>
  </div>
</section>""",
    )
)

_register(
    Layout(
        name="comparison",
        title="对比表",
        description="A/B 对比：首行表头 + 逐行对比，直标数据。",
        css="""
.comparison .cmp-row { display: flex; gap: 0; align-items: stretch; }
.comparison .cmp-row + .cmp-row { border-top: 1px solid var(--divider); }
.comparison .cmp-cell {
  flex: 1; padding: 20px 24px; font-size: var(--body); line-height: 1.5;
}
.comparison .cmp-cell:first-child { flex: 0 0 220px; color: var(--muted); }
.comparison .cmp-head { background: var(--surface); font-weight: 700; color: var(--ink); }
.comparison .cmp-col.win { color: var(--accent); font-weight: 700; }
""",
        html="""<section class="slide comparison" data-pptx-slide data-layout="comparison">
  <div class="kicker"><!-- 眉标 --></div>
  <h2 class="title"><!-- 标题：直接给对比结论 --></h2>
  <div class="cmp-row">
    <div class="cmp-cell cmp-head">维度</div>
    <div class="cmp-cell cmp-head">方案 A</div>
    <div class="cmp-cell cmp-head">方案 B</div>
  </div>
  <div class="cmp-row">
    <div class="cmp-cell">成本</div>
    <div class="cmp-cell"><!-- 数据 --></div>
    <div class="cmp-cell cmp-col win"><!-- 数据 --></div>
  </div>
</section>""",
    )
)

_register(
    Layout(
        name="chart-dominant",
        title="图表主导页",
        description="标题即洞察，图表占大面积，右/下侧放数据说明与来源。",
        css="""
.chart-dominant .chart-area { display: flex; gap: var(--gap); margin-top: 40px; }
.chart-dominant .chart {
  flex: 1; background: var(--surface); border-radius: var(--radius);
  padding: 32px; min-height: 460px;
  display: flex; align-items: flex-end; justify-content: space-around;
}
.chart-dominant .chart .bar {
  width: 96px; background: var(--accent); border-radius: 6px 6px 0 0;
}
.chart-dominant .chart .bar.muted { background: var(--muted); }
.chart-dominant .chart-note {
  flex: 0 0 360px; font-size: var(--caption); color: var(--muted); line-height: 1.7;
  padding-top: 8px;
}
""",
        html="""<section class="slide chart-dominant" data-pptx-slide data-layout="chart-dominant">
  <div class="kicker"><!-- 眉标 --></div>
  <h2 class="title"><!-- 标题即洞察 --></h2>
  <div class="chart-area">
    <div class="chart">
      <!-- 直标数据：div 高度=值；柱状图默认 -->
      <div class="bar" style="height: 60%;"></div>
      <div class="bar muted" style="height: 40%;"></div>
    </div>
    <div class="chart-note"><!-- 数据说明与来源 --></div>
  </div>
</section>""",
    )
)

_register(
    Layout(
        name="icons-row",
        title="图标行",
        description="眉标 + 标题 + 一排图标（各带标签），用于能力/价值/要点罗列。",
        css="""
.icons-row .icon-row {
  display: flex; flex-wrap: wrap; gap: var(--gap);
  justify-content: space-between; margin-top: 48px;
}
.icons-row .icon-item {
  flex: 0 0 calc((100% - 2 * var(--gap)) / 3); text-align: center;
}
.icons-row .icon-item .icon { width: 72px; height: 72px; color: var(--accent); }
.icons-row .icon-item .label {
  font-size: var(--body); color: var(--ink); margin-top: 16px; font-weight: 600;
}
""",
        html="""<section class="slide icons-row" data-pptx-slide data-layout="icons-row">
  <div class="kicker"><!-- 眉标 --></div>
  <h2 class="title"><!-- 标题 --></h2>
  <div class="icon-row">
    <div class="icon-item">
      <svg class="icon" data-icon="ph:check-circle" viewBox="0 0 256 256"
        width="72" height="72"></svg>
      <div class="label"><!-- 标签 --></div>
    </div>
    <div class="icon-item">
      <svg class="icon" data-icon="ph:trend-up" viewBox="0 0 256 256"
        width="72" height="72"></svg>
      <div class="label"><!-- 标签 --></div>
    </div>
  </div>
</section>""",
    )
)

_register(
    Layout(
        name="portrait-feature",
        title="人物 / 形象页",
        description="左图右文：形象特写 + 眉标 + 名字 + 一段介绍。",
        css="""
.portrait-feature { display: flex; gap: 72px; align-items: center; }
.portrait-feature .portrait {
  flex: 0 0 560px; height: 720px; background: var(--surface);
  border-radius: var(--radius); display: flex; align-items: center; justify-content: center;
  color: var(--muted); font-size: var(--body);
}
.portrait-feature .bio { flex: 1; }
.portrait-feature .bio .title { margin-bottom: 24px; }
.portrait-feature .bio p { font-size: var(--body); color: var(--ink); line-height: 1.7; }
""",
        html="""<section class="slide portrait-feature" data-pptx-slide
  data-layout="portrait-feature">
  <div class="portrait"><!-- 图片，或占位色块 --></div>
  <div class="bio">
    <div class="kicker"><!-- 眉标 --></div>
    <h2 class="title"><!-- 姓名 / 主题 --></h2>
    <p><!-- 介绍 --></p>
  </div>
</section>""",
    )
)

_register(
    Layout(
        name="closer",
        title="收尾页",
        description="Thank You + 联系方式，浅色变体常用，呼应封面。",
        css="""
.closer { display: flex; flex-direction: column; justify-content: center; align-items: flex-start; }
.closer .title { font-size: var(--title); }
.closer .contacts {
  margin-top: 32px; font-size: var(--body); color: var(--muted); line-height: 1.8;
}
""",
        html="""<section class="slide closer" data-pptx-slide data-layout="closer">
  <div class="kicker">Thank You</div>
  <h1 class="title"><!-- 欢迎提问与反馈 --></h1>
  <div class="contacts"><!-- 联系方式 --></div>
</section>""",
    )
)


def layouts() -> list[dict]:
    """所有布局的元数据清单（给自动选型 / 文档用）。"""
    return [
        {"name": layout.name, "title": layout.title, "description": layout.description}
        for layout in LAYOUTS.values()
    ]


def layout_css(name: str) -> str:
    """单布局 CSS（含公共排版）。"""
    if name not in LAYOUTS:
        raise KeyError(f"未知布局: {name!r}（可选: {', '.join(LAYOUTS)}）")
    return _COMMON_CSS + LAYOUTS[name].css


def layout_html(name: str) -> str:
    """布局 HTML 骨架模板（Claude 复制后填内容）。"""
    if name not in LAYOUTS:
        raise KeyError(f"未知布局: {name!r}（可选: {', '.join(LAYOUTS)}）")
    return LAYOUTS[name].html


def referenced_layouts(html: str) -> list[str]:
    """找出 HTML 里被 data-layout 引用的布局名（保序去重）。"""
    names = re.findall(r'data-layout=["\']([a-z0-9-]+)["\']', html)
    seen: list[str] = []
    for n in names:
        if n not in seen:
            seen.append(n)
    return seen


def chart_dominant_slide_indices(html_text: str) -> list[int]:
    """找出声明 chart-dominant 且内部含图表的 slide 序号（1-based，文档序）。

    条件：section 同时满足 (a) 开标签声明 `data-layout="chart-dominant"`
    （单/双引号皆可），且 (b) 内部含图表容器（`class="chart"` 或 `data-chart`）。
    跳过 (b) 避免「声明了布局但没有图表」的页误报。Deck HTML 不嵌套 section，
    故非贪婪 `(.*?)</section>` 正确。
    """
    indices: list[int] = []
    pattern = re.compile(r"(<section\b[^>]*data-pptx-slide[^>]*>)(.*?)</section>", re.S)
    for i, m in enumerate(pattern.finditer(html_text), start=1):
        opening, body = m.group(1), m.group(2)
        if re.search(r'data-layout=["\']chart-dominant["\']', opening) and (
            'class="chart"' in body or "data-chart" in body
        ):
            indices.append(i)
    return indices


def inject_layouts(html: str) -> str:
    """把 HTML 里引用的布局 CSS 注入到 <head>。

    有 `<style data-layouts></style>` 占位则替换；无则 append 到 </head> 前。
    只注入被引用的布局（+ 公共排版），未被引用的布局不进产物。
    """
    names = referenced_layouts(html)
    # 未知名称无 CSS 可注入：跳过不崩，单个拼写错误不打断整个管线
    names = [n for n in names if n in LAYOUTS]
    if not names:
        return html
    css = _COMMON_CSS + "".join(LAYOUTS[n].css for n in names)
    block = f"<style data-layouts>\n{css}</style>"
    pattern = re.compile(r"(<style data-layouts>)(.*?)(</style>)", re.S)
    if pattern.search(html):
        return pattern.sub(lambda m: f"{m.group(1)}\n{css}</style>", html, count=1)
    if "</head>" in html:
        return html.replace("</head>", block + "\n</head>", 1)
    if "<body" in html:
        return html.replace("<body", block + "\n<body", 1)
    return block + "\n" + html
