"""内容工作流：markdown outline → 逐页结构化内容 → HTML deck 骨架。

Claude/用户在对话里写一份 markdown 大纲（`# 标题` + 每个 `## 段` = 一页，
页内 `- `/`* ` 为要点、普通行为正文），outline 模块解析成结构化内容，并
可直接转成 HTML deck 骨架，省去手写骨架的重复劳动。每页布局可由行尾
`@layout: <name>` 或页内指令显式指定，否则交给 autopick 按内容信号推断。

大纲格式约定：
    # 主标题
    > 副标题（可选）

    ## 第一页标题 @layout: big-number
    @kicker: 眉标（可选）
    - 要点一
    - 要点二

    ## 第二页标题
    一段正文说明。
    @notes: 备注（可选，仅进 JSON 不进 HTML）
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from html import escape

from .autopick import pick_layouts
from .design import inject_theme

_DIRECTIVE_RE = re.compile(r"^\s*@([\w-]+)\s*:\s*(.*)$")
_LAYOUT_INLINE_RE = re.compile(r"\s*@layout:\s*([a-z0-9-]+)\s*$")
# 布局名白名单：拼进 class/data-layout 属性，必须防注入（引号/空格逃逸）
_LAYOUT_NAME_RE = re.compile(r"[a-z0-9-]+")
_CHART_TYPES = ("bar", "line", "pie")  # 本地常量，避免 import charts（内部惰性 import python-pptx）
_ICON_SETS = ("ph", "lu")  # 本地常量，避免 import icons
_ICON_VIEWBOX = {"ph": "0 0 256 256", "lu": "0 0 24 24"}
_ICON_NAME_RE = re.compile(r"[a-z0-9][a-z0-9-]*")  # 与 icons.py 一致：首字符必须字母数字


@dataclass
class SlideContent:
    index: int
    title: str
    body: list[str] = field(default_factory=list)
    bullets: list[str] = field(default_factory=list)
    kicker: str = ""
    note: str = ""
    layout: str = ""
    chart_type: str = ""
    chart_data: str = ""  # raw JSON 字符串（未转义原样存储）
    icons: list[tuple[str, str]] = field(default_factory=list)  # [(data_icon, label), ...]

    def to_dict(self) -> dict:
        d: dict = {"index": self.index, "title": self.title}
        if self.kicker:
            d["kicker"] = self.kicker
        if self.layout:
            d["layout"] = self.layout
        if self.chart_type:
            d["chart_type"] = self.chart_type
        if self.chart_data:
            d["chart_data"] = self.chart_data
        if self.body:
            d["body"] = self.body
        if self.bullets:
            d["bullets"] = self.bullets
        if self.icons:
            d["icons"] = [[data_icon, label] for data_icon, label in self.icons]
        if self.note:
            d["note"] = self.note
        return d


@dataclass
class DeckOutline:
    title: str
    subtitle: str = ""
    slides: list[SlideContent] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "title": self.title,
            **({"subtitle": self.subtitle} if self.subtitle else {}),
            "slides": [s.to_dict() for s in self.slides],
        }

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)

    def markdown(self) -> str:
        """逆序列化：DeckOutline → 同格式大纲，可再 parse_outline 还原。"""
        lines = [f"# {self.title}"]
        if self.subtitle:
            lines.append(f"> {self.subtitle}")
        for s in self.slides:
            lines.append("")
            layout = f" @layout: {s.layout}" if s.layout else ""
            lines.append(f"## {s.title}{layout}")
            if s.kicker:
                lines.append(f"@kicker: {s.kicker}")
            lines.extend(s.body)
            lines.extend(f"- {b}" for b in s.bullets)
            if s.chart_type or s.chart_data:
                if s.chart_type:
                    lines.append(f"@chart: {s.chart_type}")
                if s.chart_data:
                    lines.append(f"@chart-data: {s.chart_data}")
            if s.icons:
                items = [
                    f"{data_icon}:{label}" if label else data_icon for data_icon, label in s.icons
                ]
                lines.append(f"@icons: {'; '.join(items)}")
            if s.note:
                lines.append(f"@notes: {s.note}")
        return "\n".join(lines) + "\n"


def parse_outline(md: str) -> DeckOutline:
    """解析 markdown 大纲 → DeckOutline。格式约定见模块 docstring。"""
    title = ""
    subtitle = ""
    slides: list[SlideContent] = []
    cur: SlideContent | None = None
    pending: dict[str, str] = {}
    for raw in md.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        m = _DIRECTIVE_RE.match(line)
        if m:
            key, val = m.group(1).lower(), m.group(2).strip()
            if key not in ("layout", "kicker", "notes", "chart", "chart-data", "icons"):
                raise ValueError(
                    f"未知指令 @{key}（可选: @layout/@kicker/@notes/@chart/@chart-data/@icons）"
                )
            if key == "layout" and not _LAYOUT_NAME_RE.fullmatch(val):
                raise ValueError(
                    f"非法布局名 @layout: {val!r}（限小写字母/数字/连字符，如 big-number）"
                )
            if key == "chart" and val not in _CHART_TYPES:
                raise ValueError(f"非法图表类型 @chart: {val!r}（可选: bar/line/pie）")
            if cur is not None:
                if key == "layout":
                    cur.layout = val
                elif key == "kicker":
                    cur.kicker = val
                elif key == "notes":
                    cur.note = val
                elif key == "chart":
                    cur.chart_type = val
                    if not cur.layout:
                        cur.layout = "chart-dominant"
                elif key == "chart-data":
                    cur.chart_data = val
                elif key == "icons":
                    cur.icons = _parse_icons_value(val)
                    if not cur.layout:
                        cur.layout = "icons-row"
            else:
                pending[key] = val
            continue
        if line.startswith("> "):
            if cur is None and not subtitle:
                subtitle = line[2:].strip()
            elif cur is not None:
                cur.body.append(line[2:].strip())
            continue
        if line.startswith("# "):
            title = line[2:].strip()
            continue
        if line.startswith("## "):
            if cur is not None:
                slides.append(cur)
            text = line[3:].strip()
            minline = _LAYOUT_INLINE_RE.search(text)
            if minline:
                layout = minline.group(1)
                text = text[: minline.start()].rstrip()
            else:
                layout = ""
            chart_type = pending.pop("chart", "")
            icons = _parse_icons_value(pending.pop("icons", "")) if "icons" in pending else []
            cur = SlideContent(
                index=len(slides) + 1,
                title=text,
                layout=(
                    layout
                    or pending.pop("layout", "")
                    or (chart_type and "chart-dominant")
                    or ("icons-row" if icons else "")
                ),
                kicker=pending.pop("kicker", ""),
                note=pending.pop("notes", ""),
                chart_type=chart_type,
                chart_data=pending.pop("chart-data", ""),
                icons=icons,
            )
            continue
        if cur is None:
            continue
        if line.startswith(("- ", "* ")):
            cur.bullets.append(line[2:].strip())
        else:
            cur.body.append(line.strip())
    if cur is not None:
        slides.append(cur)
    if not title:
        raise ValueError("大纲缺少 # 主标题")
    return DeckOutline(title=title, subtitle=subtitle, slides=slides)


def _parse_icons_value(val: str) -> list[tuple[str, str]]:
    """@icons 值 → [(data_icon, label)]。每项 'ph:name:label' 或 'name:label' 或 'name'。

    前缀缺省 ph；label 缺省空。set/name 白名单校验（name 拼进 data-icon 属性）。
    label 不得含 ';' 或 ':'（分项/分段分隔符）；首段恰为 ph/lu 时按 set 解释
    （如 'lu:zap' = set lu + 图标 zap）。
    """
    icons: list[tuple[str, str]] = []
    for item in val.split(";"):
        item = item.strip()
        if not item:
            continue
        parts = [p.strip() for p in item.split(":")]
        if len(parts) == 1:
            set_, name, label = "ph", parts[0], ""
        elif len(parts) == 2:
            if parts[0] in _ICON_SETS:
                set_, name, label = parts[0], parts[1], ""
            else:
                set_, name, label = "ph", parts[0], parts[1]
        elif len(parts) == 3:
            set_, name, label = parts[0], parts[1], parts[2]
        else:
            raise ValueError(f"非法图标项 @icons: {item!r}（格式 'ph:name:label'）")
        if set_ not in _ICON_SETS:
            raise ValueError(f"非法图标集 @icons: {set_!r}（可选: {'/'.join(_ICON_SETS)}）")
        if not _ICON_NAME_RE.fullmatch(name):
            raise ValueError(
                f"非法图标名 @icons: {name!r}（限小写字母/数字/连字符，如 check-circle）"
            )
        icons.append((f"{set_}:{name}", label))
    if not icons:
        raise ValueError("@icons 不能为空")
    return icons


def _esc(text: str) -> str:
    return escape(text, quote=True)


def _slide_section(s: SlideContent) -> str:
    """单页 HTML 骨架。bullets 渲染成 .cards>.card（触发 autopick 的 cards 信号，
    且 DOM 与 cards-3 布局模板对齐），正文渲染成 .col（split 信号），
    保持与 autopick._inspect_signals 对齐。"""
    parts = ['<section class="slide" data-pptx-slide>']
    if s.kicker:
        parts.append(f'  <div class="kicker">{_esc(s.kicker)}</div>')
    parts.append(f'  <h2 class="title">{_esc(s.title)}</h2>')
    if s.chart_type:
        data_attr = f' data-chart-data="{_esc(s.chart_data)}"' if s.chart_data else ""
        parts.append(f'  <div class="chart" data-chart="{_esc(s.chart_type)}"{data_attr}></div>')
    elif s.icons:
        items = []
        for data_icon, label in s.icons:
            set_, _ = data_icon.split(":", 1)
            vb = _ICON_VIEWBOX.get(set_, "0 0 256 256")
            svg = (
                f'<svg class="icon" data-icon="{_esc(data_icon)}" '
                f'viewBox="{vb}" width="64" height="64"></svg>'
            )
            label_html = f'<div class="label">{_esc(label)}</div>' if label else ""
            items.append('    <div class="icon-item">' + svg + label_html + "</div>")
        row = "\n".join(items)
        parts.append(f'  <div class="icon-row">\n{row}\n  </div>')
    else:
        parts.extend(f'  <div class="col"><p>{_esc(b)}</p></div>' for b in s.body)
        if s.bullets:
            cards = "\n".join(
                f'    <div class="card"><div class="txt">{_esc(b)}</div></div>' for b in s.bullets
            )
            parts.append(f'  <div class="cards">\n{cards}\n  </div>')
    parts.append("</section>")
    return "\n".join(parts)


def to_deck_html(outline: DeckOutline, theme: str | None = None) -> str:
    """outline → 完整 HTML deck 骨架。

    布局：显式 @layout 优先；否则按生成骨架的内容信号交给
    autopick.pick_layouts 推断（首页→hero-title、要点→cards-3、
    正文→split-2col、含%→big-number…）。theme 给定则注入内置主题
    （design.inject_theme），否则留待后续 deck.make --theme 注入。
    """
    blocks = [_slide_section(s) for s in outline.slides]
    skeleton = (
        "<!DOCTYPE html>\n<html>\n<head>\n</head>\n<body>\n"
        + "\n".join(blocks)
        + "\n</body>\n</html>"
    )
    picks = {p.index: p.layout for p in pick_layouts(skeleton)}
    for i, (s, block) in enumerate(zip(outline.slides, blocks, strict=True)):
        # 按位置序数（i+1）关联 autopick 结果，不依赖 SlideContent.index 不变式
        layout = s.layout or picks.get(i + 1, "hero-title")
        blocks[i] = re.sub(
            r'<section class="([^"]*)" data-pptx-slide',
            rf'<section class="\1 {layout}" data-pptx-slide data-layout="{layout}"',
            block,
            count=1,
        )
    html = (
        "<!DOCTYPE html>\n<html>\n<head>\n</head>\n<body>\n"
        + "\n".join(blocks)
        + "\n</body>\n</html>"
    )
    if theme:
        html = inject_theme(html, theme)
    return html
