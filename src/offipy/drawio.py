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
