"""offipy.diagrams — Mermaid 解析（vendored mermaid_extract 包装）。"""

from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from offipy.diagrams import layout_diagram, parse_mermaid, render_to_slide


def test_parse_flowchart_basic():
    diagram = parse_mermaid("graph TD\n    A[开始] --> B[处理]")
    assert diagram.kind == "flowchart"
    assert diagram.direction == "TD"
    assert [n.id for n in diagram.nodes] == ["A", "B"]
    assert [e.source for e in diagram.edges] == ["A"]


def test_parse_chinese_label_preserved():
    diagram = parse_mermaid("graph LR\n    X[数据管道] --> Y[结果]")
    labels = {n.id: n.label for n in diagram.nodes}
    assert labels["X"] == "数据管道"


def test_parse_direction_lr():
    diagram = parse_mermaid("graph LR\n    A --> B")
    assert diagram.direction == "LR"


def test_unsupported_kind_rejected():
    with pytest.raises(ValueError, match="仅支持 flowchart"):
        parse_mermaid("sequenceDiagram\n    A->>B: hi")


def test_bad_syntax_raises_value_error():
    with pytest.raises(ValueError):
        parse_mermaid("graph TD\n    this is not mermaid !!!")


def test_extractor_systemexit_normalized_to_value_error():
    # gantt 不是 extractor 支持的 kind，_kind_and_direction 直接 _fail → SystemExit(2)。
    # 消息判别器「无法解析 Mermaid 源码」证明走的是 SystemExit 归一化分支
    # （区别于空图守卫的「未提取到任何节点或边」与 kind 检查的「仅支持 flowchart」）。
    with pytest.raises(ValueError, match="无法解析 Mermaid 源码"):
        parse_mermaid("gantt\ntitle x")


def test_mermaid_missing_direction_clear_error():
    # extractor 契约：graph/flowchart 必须带显式方向，裸 graph 会报 "not a Mermaid file"；
    # offipy 包装预检后给清晰 ValueError（含"方向"），不透传误导性消息。
    with pytest.raises(ValueError, match="方向"):
        parse_mermaid("graph\n    A[开始] --> B[处理]")


def test_layout_td_layers_and_coords():
    diagram = parse_mermaid(
        "graph TD\n    A[开始] --> B[处理]\n    B --> C[输出]\n    B --> D[结束]"
    )
    lay = layout_diagram(diagram)
    # A 在 (0,0)，B 在第二层，C/D 在第三层（各占一层内的列）
    a = next(n for n in lay.nodes if n.id == "A")
    b = next(n for n in lay.nodes if n.id == "B")
    c = next(n for n in lay.nodes if n.id == "C")
    d = next(n for n in lay.nodes if n.id == "D")
    assert b.y > a.y  # 分层：A/B 不同行
    assert c.y > b.y
    assert c.y == pytest.approx(d.y)  # C/D 同层同 y
    assert c.x != d.x  # 同层不同列
    assert all(n.w > 0 and n.h > 0 for n in lay.nodes)


def test_layout_lr_swaps_axis():
    diagram = parse_mermaid("graph LR\n    A --> B")
    lay = layout_diagram(diagram)
    a = next(n for n in lay.nodes if n.id == "A")
    b = next(n for n in lay.nodes if n.id == "B")
    assert b.x > a.x  # LR: 水平前进
    assert b.y == a.y  # 同层同列


def test_layout_fits_max():
    diagram = parse_mermaid("graph TD\n    A --> B\n    B --> C\n    C --> D")
    lay = layout_diagram(diagram, max_w=3.0, max_h=2.0)
    assert lay.canvas_w <= 3.0 + 1e-9
    assert lay.canvas_h <= 2.0 + 1e-9


def test_layout_cycle_tolerated():
    diagram = parse_mermaid("graph TD\n    A --> B\n    B --> A")
    lay = layout_diagram(diagram)
    assert {n.id for n in lay.nodes} == {"A", "B"}


def test_layout_routes_container_endpoint_edges():
    # 容器框落地后，容器端点边 A→Z 被路由到容器框（源锚底部中心，TD 方向）。
    diagram = parse_mermaid("graph TD\n    subgraph A\n        X --> Y\n    end\n    A --> Z")
    lay = layout_diagram(diagram)
    box = next(n for n in lay.nodes if n.is_container)
    assert box.id == "A"
    assert box.x >= 0 and box.y >= 0
    x = next(n for n in lay.nodes if n.id == "X")
    y = next(n for n in lay.nodes if n.id == "Y")
    assert box.x <= x.x and box.y <= y.y
    assert box.x + box.w >= x.x + x.w
    assert box.y + box.h >= y.y + y.h
    e = next(e for e in lay.edges if e.source == "A" and e.target == "Z")
    assert e.ax1 == pytest.approx(box.x + box.w / 2)
    assert e.ay1 == pytest.approx(box.y + box.h)
    assert len(lay.edges) == 2


def test_layout_subgraph_container_bbox():
    diagram = parse_mermaid(
        "graph TD\n    subgraph 数据处理\n        A[清洗] --> B[聚合]\n    end\n    B --> C[输出]"
    )
    lay = layout_diagram(diagram)
    cont = [n for n in lay.nodes if n.is_container]
    assert len(cont) == 1
    box = cont[0]
    # 容器框必须包住其子节点 A/B
    a = next(n for n in lay.nodes if n.id == "A")
    b = next(n for n in lay.nodes if n.id == "B")
    assert box.x <= a.x and box.y <= a.y
    assert box.x + box.w >= b.x + b.w
    assert box.y + box.h >= b.y + b.h
    # 顶/左缘不越过画布原点（防容器标题被裁到画布顶外）
    assert box.x >= 0 and box.y >= 0


def test_render_subgraph_container_shape():
    diagram = parse_mermaid("graph TD\n    subgraph 数据处理\n        A[清洗] --> B[聚合]\n    end")
    prs, slide = _render(diagram)
    texts = {sh.text_frame.text for sh in slide.shapes if sh.has_text_frame}
    assert "数据处理" in texts
    assert texts >= {"清洗", "聚合"}
    # 容器背景框（RECTANGLE）与叶节点都是 autoshape
    auto = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    assert len(auto) == 3


def test_render_edge_label_text():
    diagram = parse_mermaid("graph TD\n    A[发] -->|是| B[收]")
    prs, slide = _render(diagram)
    texts = {sh.text_frame.text for sh in slide.shapes if sh.has_text_frame}
    assert "是" in texts


def test_layout_nested_subgraph_bbox():
    # 嵌套 subgraph：外层框必须包住内层框（内层是外层的跨层后代），内层框包住叶子。
    diagram = parse_mermaid(
        "graph TD\n"
        "    subgraph 外层\n"
        "        subgraph 内层\n"
        "            A[开始] --> B[结束]\n"
        "        end\n"
        "    end"
    )
    lay = layout_diagram(diagram)
    boxes = {n.label: n for n in lay.nodes if n.is_container}
    assert set(boxes) == {"外层", "内层"}
    outer, inner = boxes["外层"], boxes["内层"]
    assert outer.x <= inner.x and outer.y <= inner.y
    assert outer.x + outer.w >= inner.x + inner.w
    assert outer.y + outer.h >= inner.y + inner.h
    a = next(n for n in lay.nodes if n.id == "A")
    b = next(n for n in lay.nodes if n.id == "B")
    assert inner.x <= a.x and inner.y <= a.y
    assert inner.x + inner.w >= b.x + b.w
    assert inner.y + inner.h >= b.y + b.h


@pytest.mark.parametrize(
    "direction",
    ["BT", "RL"],
)
def test_layout_reversed_directions_anchors(direction):
    # 反向方向：BT 层 1 在层 0 上方（y 更小），RL 层 1 在层 0 左侧（x 更小）。
    # 同时锁边锚点（BT: 源顶部/目标底部；RL: 源左缘/目标右缘）。
    diagram = parse_mermaid(f"graph {direction}\n    A --> B")
    lay = layout_diagram(diagram)
    a = next(n for n in lay.nodes if n.id == "A")
    b = next(n for n in lay.nodes if n.id == "B")
    if direction == "BT":
        assert b.y < a.y
        assert b.x == a.x
    else:
        assert b.x < a.x
        assert b.y == a.y
    edge = lay.edges[0]
    if direction == "BT":
        assert edge.ax1 == pytest.approx(a.x + a.w / 2)
        assert edge.ay1 == pytest.approx(a.y)
        assert edge.ax2 == pytest.approx(b.x + b.w / 2)
        assert edge.ay2 == pytest.approx(b.y + b.h)
    else:  # RL
        assert edge.ax1 == pytest.approx(a.x)
        assert edge.ay1 == pytest.approx(a.y + a.h / 2)
        assert edge.ax2 == pytest.approx(b.x + b.w)
        assert edge.ay2 == pytest.approx(b.y + b.h / 2)


def _render(diagram, max_w=12.0, max_h=6.75):
    lay = layout_diagram(diagram, max_w=max_w, max_h=max_h)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    render_to_slide(slide, lay)
    return prs, slide


def _shape_map(slide):
    return {sh.shape_type: sh for sh in slide.shapes}


def test_render_creates_editable_autoshapes():
    diagram = parse_mermaid("graph TD\n    A[开始] --> B{判断} --> C[结束]")
    prs, slide = _render(diagram)
    auto = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    conn = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.LINE]
    assert len(auto) == 3
    assert len(conn) == 2
    texts = {sh.text_frame.text for sh in slide.shapes if sh.has_text_frame}
    assert texts >= {"开始", "判断", "结束"}


def test_render_shape_mapping_rhombus():
    diagram = parse_mermaid("graph TD\n    A{决策} --> B[结果]")
    prs, slide = _render(diagram)
    diamond = [
        sh
        for sh in slide.shapes
        if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and sh.auto_shape_type == MSO_SHAPE.DIAMOND
    ]
    assert len(diamond) == 1
    assert diamond[0].text_frame.text == "决策"


def test_render_arrowheads_present():
    diagram = parse_mermaid("graph TD\n    A --> B")
    prs, slide = _render(diagram)
    conns = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.LINE]
    assert len(conns) == 1
    ln = conns[0].line._get_or_add_ln()
    tail = ln.find(qn("a:tailEnd"))
    assert tail is not None
    assert tail.get("type") == "triangle"


def test_render_dashed_edge_when_dotted():
    diagram = parse_mermaid("graph TD\n    A -.-> B")
    prs, slide = _render(diagram)
    conns = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.LINE]
    assert conns and conns[0].line.dash_style == MSO_LINE_DASH_STYLE.DASH


def test_render_chinese_font_set():
    diagram = parse_mermaid("graph TD\n    A[数据] --> B[管道]")
    prs, slide = _render(diagram)
    assert slide.shapes
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            assert para.runs
            for run in para.runs:
                rPr = run._r.find(qn("a:rPr"))
                assert rPr is not None
                for tag in ("a:latin", "a:ea", "a:cs"):
                    el = rPr.find(qn(tag))
                    assert el is not None, f"{tag} 未设置"
                    assert el.get("typeface") == "Microsoft YaHei"


def test_render_no_arrowhead_when_undirected():
    diagram = parse_mermaid("graph TD\n    A --- B")
    prs, slide = _render(diagram)
    conns = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.LINE]
    assert conns
    ln = conns[0].line._get_or_add_ln()
    assert ln.find(qn("a:tailEnd")) is None
