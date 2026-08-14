"""draw.io 图 → PPTX 原生可编辑形状（保留作者版式与配色）。

输入 .drawio 文件 → vendored drawio_extract.py 提取 IR（节点自带绝对坐标与样式）
→ offipy 薄层：选页 + 坐标归一化 fit（layout_drawio）→ 复用 diagrams.render_to_slide
渲染成可编辑形状（autoshape + connector）。可独立生成整页 PPTX（drawio_to_pptx），
也可作为 deck 后处理注入（postprocess_drawio：HTML <div class="drawio"
data-drawio="..."> → 替换为可编辑形状）。

vendored 提取器用 importlib 加载，保持上游文件原样；安全边界（DTD/ENTITY 拒绝、
压缩上限）继承自 vendor 的 parse_file，offipy 不自行解析 draw.io XML。
"""

from __future__ import annotations

import importlib.util
import sys
from dataclasses import dataclass
from pathlib import Path

from offipy.diagrams import DiagramLayout, PlacedEdge, PlacedNode, render_to_slide

_EXTRACT_REL = Path("_vendor/diagram-design/skills/diagram-design/scripts/drawio_extract.py")

_extractor = None


def _load_extractor():
    """惰性加载 vendored drawio_extract（importlib，保持上游原样）。"""
    global _extractor
    if _extractor is None:
        script = Path(__file__).resolve().parent / _EXTRACT_REL
        if not script.exists():
            raise RuntimeError(f"vendored drawio_extract 缺失: {script}")
        name = "offipy_vendored_drawio_extract"
        spec = importlib.util.spec_from_file_location(name, script)
        mod = importlib.util.module_from_spec(spec)
        sys.modules[name] = mod
        assert spec.loader is not None
        spec.loader.exec_module(mod)
        _extractor = mod
    return _extractor


@dataclass
class DrawioNode:
    id: str
    label: str = ""
    shape: str = "rect"
    parent: str | None = None
    x: float = 0.0
    y: float = 0.0
    w: float = 0.0
    h: float = 0.0
    fill: str = ""
    stroke: str = ""
    font_color: str = ""
    dashed: bool = False
    rounded: bool = False
    container: bool = False


@dataclass
class DrawioEdge:
    source: str
    target: str
    label: str = ""
    dashed: bool = False
    bidirectional: bool = False
    undirected: bool = False
    stroke: str = ""


@dataclass
class DrawioDiagram:
    nodes: list[DrawioNode]
    edges: list[DrawioEdge]


def parse_drawio(source, *, page=None) -> DrawioDiagram:
    """.drawio 文件 → DrawioDiagram（vendored parse_file + select_pages 选页）。

    page: int → 页码（0 起，对齐 draw.io 索引）；str → 数字串按 index、否则按页名
    原样传给 vendored select_pages；None → 第一页。单页输出，不暴露 "all"。

    安全边界：只调 vendored parse_file（继承 DTD/ENTITY 拒绝与压缩上限），
    绝不自行 ET.fromstring 未经验证的原始输入。
    """
    path = Path(source)
    if not path.is_file():
        raise FileNotFoundError(f"源文件不存在: {path}")
    ex = _load_extractor()
    selector = None if page is None else (str(page) if isinstance(page, int) else page)
    try:
        pages = ex.parse_file(path)
        selected = ex.select_pages(pages, selector)
    except SystemExit as e:
        code = e.code if isinstance(e.code, str) else str(e)
        raise ValueError(f"无法解析 draw.io 源码: {code}") from None
    if not selected:
        raise ValueError("无法解析 draw.io 源码: 未找到所选页")
    p = selected[0]
    nodes = [
        DrawioNode(
            id=n.id,
            label=n.label,
            shape=n.shape,
            parent=n.parent,
            x=n.x,
            y=n.y,
            w=n.w,
            h=n.h,
            fill=n.fill,
            stroke=n.stroke,
            font_color=n.font_color,
            dashed=n.dashed,
            rounded=n.rounded,
            container=n.container,
        )
        for n in p.nodes
    ]
    # 悬空边（vendored 已把指向不存在节点的 source/target 置 None）直接过滤
    edges = [
        DrawioEdge(
            source=e.source,
            target=e.target,
            label=e.label,
            dashed=e.dashed,
            bidirectional=e.bidirectional,
            undirected=e.undirected,
            stroke=e.stroke,
        )
        for e in p.edges
        if e.source and e.target
    ]
    if not nodes and not edges:
        raise ValueError("无法解析 draw.io 源码: 未提取到任何节点或边")
    return DrawioDiagram(nodes, edges)


def _render_shape(shape: str, rounded: bool) -> str:
    """canonical shape → 渲染层 shape 名（rounded 决定 rect 圆角与否）。"""
    if shape == "rect":
        return "round" if rounded else "rectangle"
    if shape in (
        "ellipse",
        "triangle",
        "rhombus",
        "hexagon",
        "parallelogram",
        "cylinder",
    ):
        return shape
    return "round" if rounded else "rectangle"


def layout_drawio(
    diagram: DrawioDiagram, *, max_w: float = 12.0, max_h: float = 6.75
) -> DiagramLayout:
    """把 draw.io IR 布局成坐标（inches）：归一化到原点 + 等比 fit + shape 名合成 + 颜色透传。

    vendored parse_page 已解析绝对坐标（父链累加）；这里只做：平移让最小坐标为 0、
    整体等比缩放 fit 到 max_w×max_h（不改变相对位置与配色）、canonical shape → 渲染层
    shape 名。容器节点保留自身几何（draw.io 已摆好），render 层画背景框。
    """
    nodes = diagram.nodes
    if not nodes:
        raise ValueError("无法布局 draw.io 图: 无节点")
    xs = [n.x for n in nodes] + [n.x + n.w for n in nodes]
    ys = [n.y for n in nodes] + [n.y + n.h for n in nodes]
    x0, y0 = min(xs), min(ys)
    x1, y1 = max(xs), max(ys)
    raw_w, raw_h = x1 - x0, y1 - y0
    scale = min(
        1.0,
        max_w / raw_w if raw_w > 0 else 1.0,
        max_h / raw_h if raw_h > 0 else 1.0,
    )
    placed = [
        PlacedNode(
            id=n.id,
            label=n.label,
            shape=_render_shape(n.shape, n.rounded),
            x=(n.x - x0) * scale,
            y=(n.y - y0) * scale,
            w=n.w * scale,
            h=n.h * scale,
            is_container=n.container,
            parent=n.parent,
            fill=n.fill,
            stroke=n.stroke,
            font_color=n.font_color,
        )
        for n in nodes
    ]
    edges = _layout_edges(diagram, placed)
    return DiagramLayout(placed, edges, raw_w * scale, raw_h * scale)


def _edge_anchors(s: PlacedNode, t: PlacedNode) -> tuple[tuple[float, float], tuple[float, float]]:
    """按源/目标相对位置选边缘中点（drawio 无方向概念，固定规则，不做 BT/RL 反转）。"""
    scx, scy = s.x + s.w / 2, s.y + s.h / 2
    tcx, tcy = t.x + t.w / 2, t.y + t.h / 2
    if abs(tcx - scx) >= abs(tcy - scy):
        if tcx > scx:
            return (s.x + s.w, s.y + s.h / 2), (t.x, t.y + t.h / 2)
        return (s.x, s.y + s.h / 2), (t.x + t.w, t.y + t.h / 2)
    if tcy > scy:
        return (s.x + s.w / 2, s.y + s.h), (t.x + t.w / 2, t.y)
    return (s.x + s.w / 2, s.y), (t.x + t.w / 2, t.y + t.h)


def _layout_edges(diagram: DrawioDiagram, placed: list[PlacedNode]) -> list[PlacedEdge]:
    by_id = {n.id: n for n in placed}
    out: list[PlacedEdge] = []
    for e in diagram.edges:
        s, t = by_id.get(e.source), by_id.get(e.target)
        if s is None or t is None:
            continue  # 悬空边（parse 已过滤，此处防御）
        a1, a2 = _edge_anchors(s, t)
        style = "dashed" if e.dashed else "solid"
        out.append(
            PlacedEdge(
                e.source,
                e.target,
                e.label,
                style,
                "arrow",
                e.undirected,
                a1[0],
                a1[1],
                a2[0],
                a2[1],
                stroke=e.stroke,
            )
        )
    return out


def drawio_to_pptx(source, out_path: str, *, page=None) -> str:
    """.drawio 文件 → 可编辑 PPTX（16:9 整页）。返回 out_path。

    python-pptx 惰性 import（deck extra），未安装时给可操作错误。
    """
    diagram = parse_drawio(source, page=page)
    lay = layout_drawio(diagram)
    try:
        from pptx import Presentation
        from pptx.util import Emu
    except ImportError as e:  # pragma: no cover
        raise RuntimeError("drawio_to_pptx 需要 python-pptx，请安装 offipy[deck]") from e
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    render_to_slide(slide, lay)
    prs.save(out_path)
    return str(out_path)
