"""Mermaid 图 → PPTX 原生可编辑形状（结构图渲染，非原生图表）。

输入 Mermaid flowchart 源码 → vendored mermaid_extract.py 提取 IR（拓扑）
→ offipy 分层布局（layout_diagram）→ python-pptx 渲染成可编辑形状
（autoshape + connector，render_to_slide）。可独立生成整页 PPTX
（mermaid_to_pptx），也可作为 deck 后处理注入（postprocess_mermaid：
HTML <pre class="mermaid"> 块 → 替换为可编辑形状）。

vendored 提取器用 importlib 加载，保持上游文件原样；仅支持 flowchart
（TD/TB/LR/RL/BT）；sequence/state/er 及不支持语法 → ValueError。
"""

from __future__ import annotations

import importlib.util
import re
import sys
from collections import deque
from dataclasses import dataclass
from pathlib import Path

_EXTRACT_REL = Path("_vendor/diagram-design/skills/diagram-design/scripts/mermaid_extract.py")
SUPPORTED_DIRECTIONS = {"TD", "TB", "LR", "RL", "BT"}
# vendored extractor 契约：graph/flowchart 必须带显式方向（否则 _kind_and_direction
# 直接 _fail("not a Mermaid file")）。裸 graph 是常见用户输入，parse_mermaid 用它做
# 预检，把误导性的 "not a Mermaid file" 换成清晰错误。
_FLOW_HEADER_WITHOUT_DIRECTION = re.compile(r"^\s*(graph|flowchart)\s*$", re.I)

_extractor = None


def _load_extractor():
    """惰性加载 vendored mermaid_extract（importlib，保持上游原样）。"""
    global _extractor
    if _extractor is None:
        script = Path(__file__).resolve().parent / _EXTRACT_REL
        if not script.exists():
            raise RuntimeError(f"vendored mermaid_extract 缺失: {script}")
        name = "offipy_vendored_mermaid_extract"
        spec = importlib.util.spec_from_file_location(name, script)
        mod = importlib.util.module_from_spec(spec)
        sys.modules[name] = mod
        assert spec.loader is not None
        spec.loader.exec_module(mod)
        _extractor = mod
    return _extractor


def parse_mermaid(text: str):
    """Mermaid 文本 → vendored Diagram IR。仅接受 flowchart。

    mermaid_extract 对坏输入/不支持语法用 SystemExit(2) 退出，这里捕获并
    归一化成 ValueError（调用方是 deck 后处理/独立 API，都按用户输入错误处理）。
    """
    # 预检裸 graph/flowchart（无方向）：extractor 会报误导性的 "not a Mermaid file"，
    # 这里先拦成清晰错误（Mermaid 标准默认 TD，但 vendor 契约要求显式方向）。
    first = text.splitlines()[0] if text.splitlines() else ""
    if _FLOW_HEADER_WITHOUT_DIRECTION.match(first):
        raise ValueError(
            "flowchart/graph 需要显式方向（TD/TB/LR/RL/BT），收到缺方向的 graph/flowchart"
        )
    ex = _load_extractor()
    block = ex.SourceBlock(0, text, 1)
    try:
        diagram = ex.parse_block(block)
    except SystemExit as e:
        code = e.code if isinstance(e.code, str) else str(e)
        raise ValueError(f"无法解析 Mermaid 源码: {code}") from None
    if diagram.kind != "flowchart":
        raise ValueError(
            f"仅支持 flowchart/graph，收到 {diagram.kind}（sequence/state/er 后续迭代）"
        )
    # vendor 对部分坏输入不抛 SystemExit，而是静默产出空图（无节点无边）。
    # 空 flowchart 无渲染意义，统一按用户输入错误处理（plan 期望：坏语法 → ValueError）。
    if not diagram.nodes and not diagram.edges:
        raise ValueError("无法解析 Mermaid 源码: 未提取到任何节点或边")
    return diagram


_NODE_W = 2.0
_NODE_H = 0.6
_GAP_X = 0.9
_GAP_Y = 1.1
_MARGIN = 0.5


@dataclass
class PlacedNode:
    id: str
    label: str
    shape: str
    x: float
    y: float
    w: float
    h: float
    is_container: bool = False
    parent: str | None = None


@dataclass
class PlacedEdge:
    source: str
    target: str
    label: str = ""
    style: str = "solid"
    arrowhead: str = "arrow"
    undirected: bool = False
    ax1: float = 0.0
    ay1: float = 0.0
    ax2: float = 0.0
    ay2: float = 0.0


@dataclass
class DiagramLayout:
    nodes: list[PlacedNode]
    edges: list[PlacedEdge]
    canvas_w: float
    canvas_h: float


def _kahn_layers(ids: list[str], edges) -> tuple[dict[str, int], dict[int, list[str]]]:
    """Kahn 拓扑分层。返回 (node→layer, layer→node_id 列表)。环残留追加尾层。"""
    adj: dict[str, list[str]] = {i: [] for i in ids}
    indeg = {i: 0 for i in ids}
    for e in edges:
        if e.source in adj and e.target in adj:
            adj[e.source].append(e.target)
            indeg[e.target] += 1
    q = deque(i for i in ids if indeg[i] == 0)
    order: list[str] = []
    while q:
        i = q.popleft()
        order.append(i)
        for t in adj[i]:
            indeg[t] -= 1
            if indeg[t] == 0:
                q.append(t)
    for i in ids:
        if i not in order:
            order.append(i)  # 环残留：容忍
    layer = {i: 0 for i in ids}
    for i in order:
        for t in adj[i]:
            layer[t] = max(layer[t], layer[i] + 1)
    layers: dict[int, list[str]] = {}
    for i in order:
        layers.setdefault(layer[i], []).append(i)
    return layer, layers


def layout_diagram(
    diagram, *, direction: str | None = None, max_w: float = 12.0, max_h: float = 6.75
) -> DiagramLayout:
    """把 flowchart IR 布局成坐标（inches）。方向 TD/TB/LR/RL/BT；整体 fit 到 max_w/max_h。

    - 叶子（非 container）按 Kahn 分层；layer 间距、层内列距。
    - container 节点不占独立坐标，作为背景框覆盖其子孙叶子 bbox（Task 4 画）。
    - 反向方向（BT/RL）反转层序/列序。
    """
    dr = (direction or diagram.direction or "TD").upper()
    if dr not in SUPPORTED_DIRECTIONS:
        raise ValueError(f"不支持方向 {dr}（可选 TD/TB/LR/RL/BT）")
    leaves = [n for n in diagram.nodes if not (n.container or n.shape == "container")]
    leaves_by_id = {n.id: n for n in leaves}

    layer_of, layers = _kahn_layers(list(leaves_by_id), diagram.edges)
    col_of: dict[str, int] = {}
    for ids in layers.values():
        for col, i in enumerate(ids):
            col_of[i] = col
    max_row = max(layers) if layers else 0
    max_col = max((len(v) for v in layers.values()), default=0)

    vertical = dr in ("TD", "TB", "BT")
    if vertical:
        raw_w = _MARGIN + max_col * (_NODE_W + _GAP_X) + _MARGIN
        raw_h = _MARGIN + (max_row + 1) * (_NODE_H + _GAP_Y) + _MARGIN
    else:
        raw_w = _MARGIN + (max_row + 1) * (_NODE_W + _GAP_X) + _MARGIN
        raw_h = _MARGIN + max_col * (_NODE_H + _GAP_Y) + _MARGIN
    scale = min(1.0, max_w / raw_w if raw_w > 0 else 1.0, max_h / raw_h if raw_h > 0 else 1.0)
    W, H = _NODE_W * scale, _NODE_H * scale
    GX, GY = _GAP_X * scale, _GAP_Y * scale
    M = _MARGIN * scale

    placed: list[PlacedNode] = []
    for i, n in leaves_by_id.items():
        lyr = layer_of[i]
        col = col_of[i]
        if vertical:
            row_idx = (max_row - lyr) if dr == "BT" else lyr
            px = M + col * (W + GX)
            py = M + row_idx * (H + GY)
        else:
            # RL 沿主方向（x/layer）反转：layer0 在最右；垂直分布（y/col）不镜像。
            # 对应 _layout_edges 的 RL 分支（a1=源左边缘、a2=目标右边缘）依赖此反转。
            px = M + (max_row - lyr) * (W + GX) if dr == "RL" else M + lyr * (W + GX)
            py = M + col * (H + GY)
        placed.append(PlacedNode(i, n.label, n.shape, px, py, W, H, False, n.parent))

    placed_edges = _layout_edges(diagram, placed, vertical, dr)
    canvas_w = M + (max_col * (W + GX)) if vertical else M + ((max_row + 1) * (W + GX))
    canvas_h = M + ((max_row + 1) * (H + GY)) if vertical else M + (max_col * (H + GY))
    return DiagramLayout(placed, placed_edges, canvas_w, canvas_h)


def _layout_edges(
    diagram, placed: list[PlacedNode], vertical: bool, direction: str
) -> list[PlacedEdge]:
    by_id = {n.id: n for n in placed}
    out: list[PlacedEdge] = []
    for e in diagram.edges:
        # 可达分支：容器端点边在 Task 2 布局阶段被跳过（容器不占坐标，无框可锚）。
        # Task 4 把 placed_nodes 传给本函数后容器框入 by_id，此分支转为真正死路径。
        if e.source not in by_id or e.target not in by_id:
            continue
        s, t = by_id[e.source], by_id[e.target]
        if vertical:
            if direction == "BT":
                a1 = (s.x + s.w / 2, s.y)
                a2 = (t.x + t.w / 2, t.y + t.h)
            else:
                a1 = (s.x + s.w / 2, s.y + s.h)
                a2 = (t.x + t.w / 2, t.y)
        else:
            if direction == "RL":
                a1 = (s.x, s.y + s.h / 2)
                a2 = (t.x + t.w, t.y + t.h / 2)
            else:
                a1 = (s.x + s.w, s.y + s.h / 2)
                a2 = (t.x, t.y + t.h / 2)
        out.append(
            PlacedEdge(
                e.source,
                e.target,
                e.label,
                e.style,
                e.arrowhead,
                e.undirected,
                a1[0],
                a1[1],
                a2[0],
                a2[1],
            )
        )
    return out
