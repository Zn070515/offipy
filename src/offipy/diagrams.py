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
import os
import re
import sys
from collections import deque
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path

_EXTRACT_REL = Path(
    "_vendor/diagram-design/skills/diagram-design/scripts/mermaid_extract.py"
)
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
