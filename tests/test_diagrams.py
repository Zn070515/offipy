"""offipy.diagrams — Mermaid 解析（vendored mermaid_extract 包装）。"""

from __future__ import annotations

import pytest

from offipy.diagrams import layout_diagram, parse_mermaid


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
