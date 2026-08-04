"""自动选型：从 deck HTML 内容结构推断主题 + 每页布局，并解释理由。

对标 mckinsey-pptx 的「选模板 subagent」思路（见 docs/ppt_design_research.md
P2），但用纯规则、零依赖实现——Claude 写 deck 后调用 pick()，拿到推荐结果，
自己决定采纳与否。

推荐原则：
- 主题看全文关键词分布（科技/学术/商务 → dark-tech / academic / mckinsey）。
- 每页布局看内容信号（大数字→big-number、对比→comparison、引言→quote-frame…）。
- slide 已显式打 data-layout 则尊重显式选择，不再推断。

返回结果可序列化为 JSON，也可渲染成 Markdown 供 Claude 迭代。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .design import THEMES
from .layouts import LAYOUTS

# ---------------------------------------------------------------- 信号常量

# 主题关键词：按「每词 +1 分」累计，谁高选谁
_TECH_WORDS = (
    "data",
    "engine",
    "cloud",
    "tech",
    "ai",
    "ml",
    "product",
    "api",
    "compute",
    "model",
    "software",
    "数据",
    "引擎",
    "云",
    "模型",
    "算法",
)
_ACADEMIC_WORDS = (
    "research",
    "study",
    "paper",
    "academic",
    "theory",
    "literature",
    "literature review",
    "finding",
    "研究",
    "文献",
    "学术",
    "论文",
    "理论",
)
_DEFAULT_THEME = "mckinsey"


def _inner_text(html: str) -> str:
    """剥掉标签，合并空白，拿到可见文本（用于关键词/信号统计）。"""
    text = re.sub(r"<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", text).strip()


def _has_class(block: str, name: str) -> bool:
    return re.search(r'class="[^"]*\b' + re.escape(name) + r'\b[^"]*"', block) is not None


def _slide_blocks(html: str) -> list[str]:
    """按 data-pptx-slide 切出每页 HTML 块，含 section 开标签（data-layout 在标签上）。

    deck 约定 section 不嵌套，非贪婪匹配安全。
    """
    pattern = re.compile(r"<section[^>]*data-pptx-slide[^>]*>.*?</section>", re.S)
    return pattern.findall(html)


def _explicit_layout(block: str) -> str | None:
    m = re.search(r'data-layout="([a-z0-9-]+)"', block)
    return m.group(1) if m else None


# ---------------------------------------------------------------- 单页布局


def _inspect_signals(block: str) -> dict:
    text = _inner_text(block)
    return {
        "has_kicker": _has_class(block, "kicker"),
        "has_h1": "<h1" in block,
        "has_percent": bool(re.search(r"[+-]?\d{1,3}(?:\.\d+)?\s*%", text)),
        "has_chart": _has_class(block, "chart") or _has_class(block, "bar"),
        "has_comparison": (
            _has_class(block, "cmp") or ("对比" in text) or bool(re.search(r"\bvs\b", text.lower()))
        ),
        "has_timeline": _has_class(block, "tl-") or _has_class(block, "timeline"),
        "has_cards": _has_class(block, "card"),
        "has_quote": "<blockquote" in block or _has_class(block, "quote"),
        "has_portrait": _has_class(block, "portrait"),
        "has_cols": _has_class(block, "col"),
        "has_closer": bool(re.search(r"thank\s*you|谢谢|感谢|联系", text, re.I)),
        "text_len": len(text),
    }


def _pick_layout_for_slide(index: int, block: str) -> tuple[str, str, dict]:
    """推断单页布局 → (layout, reason, signals)。"""
    signals = _inspect_signals(block)
    if index == 1:
        return "hero-title", "首屏封面：一页一观点，留白最大", signals
    if signals["has_closer"]:
        return "closer", "收尾页：Thank You + 联系方式，呼应封面", signals
    if signals["has_quote"]:
        return "quote-frame", "引言/过渡页：大引言 + 出处", signals
    if signals["has_comparison"]:
        return "comparison", "A/B 对比表：首行表头 + 直标数据", signals
    if signals["has_timeline"]:
        return "timeline", "时间线：日期 + 要点，展示进程/里程碑", signals
    if signals["has_percent"] and signals["text_len"] < 200:
        return "big-number", "单一关键数字 + 一句结论，highlight the one", signals
    if signals["has_chart"]:
        return "chart-dominant", "图表主导：标题即洞察，图表占大面积", signals
    if signals["has_portrait"]:
        return "portrait-feature", "人物/形象页：左图右文", signals
    if signals["has_cards"]:
        return "cards-3", "三卡片：眉标 + 标题 + 并列要点", signals
    if signals["has_cols"]:
        return "split-2col", "左右两栏：A/B 并排对比内容", signals
    return "hero-title", "章节页：大标题 + 副标题，一页一观点", signals


def _count_keywords(words: tuple[str, ...], text: str) -> int:
    """英文词按 \b 边界匹配（避免 'ai' 误中 'daily'），中文直接子串匹配。"""
    n = 0
    for w in words:
        if w.isascii():
            if re.search(rf"\b{re.escape(w)}\b", text):
                n += 1
        elif w in text:
            n += 1
    return n


def _pick_theme(html: str) -> tuple[str, str]:
    """按全文关键词分布选主题 → (theme, reason)。"""
    text = _inner_text(html).lower()
    tech = _count_keywords(_TECH_WORDS, text)
    academic = _count_keywords(_ACADEMIC_WORDS, text)
    if tech > academic and tech > 0:
        return (
            "dark-tech",
            f"检测到科技关键词 ×{tech}（> 学术 ×{academic}）：深色底 + 高对比强调色",
        )
    if academic > 0:
        return "academic", f"检测到学术关键词 ×{academic}：米白底 + 深蓝强调，克制排版"
    return _DEFAULT_THEME, "未检测到强风格信号，默认商务蓝（可改 theme= 覆盖）"


# ---------------------------------------------------------------- 结果结构


@dataclass
class SlidePick:
    index: int
    layout: str
    reason: str
    signals: dict = field(default_factory=dict)
    explicit: bool = False

    def to_dict(self) -> dict:
        return {
            "index": self.index,
            "layout": self.layout,
            "reason": self.reason,
            "explicit": self.explicit,
            **({"signals": self.signals} if self.signals else {}),
        }


@dataclass
class DeckPick:
    theme: str
    theme_reason: str
    slides: list[SlidePick] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "theme": self.theme,
            "theme_reason": self.theme_reason,
            "slides": [s.to_dict() for s in self.slides],
        }

    def markdown(self) -> str:
        lines = ["# 自动选型推荐", "", f"- 主题：**{self.theme}** —— {self.theme_reason}", ""]
        lines.append("| 页 | 布局 | 理由 |")
        lines.append("|---|------|------|")
        for s in self.slides:
            mark = "（显式指定）" if s.explicit else ""
            lines.append(f"| {s.index} | `{s.layout}`{mark} | {s.reason} |")
        return "\n".join(lines)


# ---------------------------------------------------------------- 入口


def pick_layouts(html: str) -> list[SlidePick]:
    """逐页推断布局。已有 data-layout 的 slide 尊重显式选择。"""
    picks: list[SlidePick] = []
    for index, block in enumerate(_slide_blocks(html), start=1):
        explicit = _explicit_layout(block)
        if explicit and explicit in LAYOUTS:
            picks.append(SlidePick(index, explicit, "slide 已显式指定 data-layout", explicit=True))
            continue
        layout, reason, signals = _pick_layout_for_slide(index, block)
        picks.append(SlidePick(index, layout, reason, signals=signals))
    return picks


def pick_theme(html: str) -> tuple[str, str]:
    """推断整份 deck 的主题 → (theme, reason)。"""
    return _pick_theme(html)


def pick(html: str) -> DeckPick:
    """总入口：主题 + 每页布局 + 理由。"""
    theme, theme_reason = _pick_theme(html)
    return DeckPick(theme=theme, theme_reason=theme_reason, slides=pick_layouts(html))


def pick_file(html_path: str) -> DeckPick:
    """从 HTML 文件路径选型。"""
    with open(html_path, encoding="utf-8") as f:
        return pick(f.read())


# ---------------------------------------------------------------- 主题/布局清单


def available_themes() -> list[str]:
    """内置主题名（design.py）。"""
    return list(THEMES)


def available_layouts() -> list[str]:
    """内置布局名（layouts.py）。"""
    return list(LAYOUTS)
