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

_DIRECTIVE_RE = re.compile(r"^\s*@(\w+)\s*:\s*(.*)$")
_LAYOUT_INLINE_RE = re.compile(r"\s*@layout:\s*([a-z0-9-]+)\s*$")


@dataclass
class SlideContent:
    index: int
    title: str
    body: list[str] = field(default_factory=list)
    bullets: list[str] = field(default_factory=list)
    kicker: str = ""
    note: str = ""
    layout: str = ""

    def to_dict(self) -> dict:
        d: dict = {"index": self.index, "title": self.title}
        if self.kicker:
            d["kicker"] = self.kicker
        if self.layout:
            d["layout"] = self.layout
        if self.body:
            d["body"] = self.body
        if self.bullets:
            d["bullets"] = self.bullets
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
            for b in s.body:
                lines.append(b)
            for b in s.bullets:
                lines.append(f"- {b}")
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
            if key not in ("layout", "kicker", "notes"):
                raise ValueError(f"未知指令 @{key}（可选: @layout/@kicker/@notes）")
            if cur is not None:
                if key == "layout":
                    cur.layout = val
                elif key == "kicker":
                    cur.kicker = val
                elif key == "notes":
                    cur.note = val
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
            cur = SlideContent(
                index=len(slides) + 1,
                title=text,
                layout=layout or pending.pop("layout", ""),
                kicker=pending.pop("kicker", ""),
                note=pending.pop("notes", ""),
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
