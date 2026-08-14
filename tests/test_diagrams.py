"""offipy.diagrams — Mermaid 解析（vendored mermaid_extract 包装）。"""

from __future__ import annotations

import pytest

from offipy.diagrams import parse_mermaid


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


def test_mermaid_missing_direction_clear_error():
    # extractor 契约：graph/flowchart 必须带显式方向，裸 graph 会报 "not a Mermaid file"；
    # offipy 包装预检后给清晰 ValueError（含"方向"），不透传误导性消息。
    with pytest.raises(ValueError, match="方向"):
        parse_mermaid("graph\n    A[开始] --> B[处理]")
