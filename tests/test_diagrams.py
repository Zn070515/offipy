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


def test_layout_skips_container_endpoint_edges():
    # 容器不占坐标，容器端点边（A→Z）在 Task 2 布局阶段被跳过。
    # Task 4 容器框落地后改为路由（届时把本测试改成断言 A→Z 存在且锚在框上）。
    diagram = parse_mermaid("graph TD\n    subgraph A\n        X --> Y\n    end\n    A --> Z")
    lay = layout_diagram(diagram)
    placed_ids = {n.id for n in lay.nodes}
    assert placed_ids == {"X", "Y", "Z"}
    assert all(e.source in placed_ids and e.target in placed_ids for e in lay.edges)
    assert not any(e.source == "A" or e.target == "A" for e in lay.edges)


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
    for sh in slide.shapes:
        if sh.has_text_frame:
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    rPr = run._r.find(qn("a:rPr"))
                    if rPr is not None:
                        ea = rPr.find(qn("a:ea"))
                        if ea is not None:
                            assert ea.get("typeface") == "Microsoft YaHei"


def test_render_no_arrowhead_when_undirected():
    diagram = parse_mermaid("graph TD\n    A --- B")
    prs, slide = _render(diagram)
    conns = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.LINE]
    assert conns
    ln = conns[0].line._get_or_add_ln()
    assert ln.find(qn("a:tailEnd")) is None
