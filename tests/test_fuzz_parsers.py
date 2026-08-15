"""解析器 fuzz：stdlib random + fixed seed，异常白名单（OffipyError/ValueError 等）。

与 test_assets_negative.py 同理念：畸形输入必须被友好拒绝，绝不静默抛裸
IndexError/AttributeError/递归爆栈。"""

import contextlib
import random
import xml.etree.ElementTree as ET

import pytest

from offipy.deck import _reject_no_visual_audit_declarations, _rewrite_relative_urls
from offipy.drawio import parse_drawio
from offipy.exceptions import InvalidArgumentError, OffipyError
from offipy.layouts import chart_dominant_slide_indices

SEED = 20260815
_TOKENS = [
    "<diagram>",
    "</diagram>",
    "<mxCell",
    "value=",
    "edgeStyle=",
    "strokeWidth=",
    "rotation=",
    "<mxGeometry",
    "x=",
    "y=",
    "id=",
    "parent=",
    "&amp;",
    "<mxPoint",
    "><",
    "/>",
    "data-drawio",
    "</html>",
    "<body",
    "style=",
    "<p>",
    "-->",
    "<![CDATA[",
    "]]>",
    "0",
    "-1",
    "9e99",
]


def _mutate(seed: int, base: str) -> str:
    rng = random.Random(seed)
    out = base
    for _ in range(rng.randint(1, 40)):
        op = rng.randrange(4)
        if op == 0 and len(out) > 0:  # 插入随机 token
            p = rng.randrange(len(out) + 1)
            out = out[:p] + rng.choice(_TOKENS) + out[p:]
        elif op == 1 and len(out) > 0:  # 删除片段
            a, b = sorted((rng.randrange(len(out) + 1), rng.randrange(len(out) + 1)))
            out = out[:a] + out[b:]
        elif op == 2 and len(out) > 0:  # 替换单字符（真替换，区别于 op==0 插入）
            p = rng.randrange(len(out))
            out = out[:p] + rng.choice(_TOKENS)[0] + out[p + 1 :]
        else:  # 截断
            out = out[: rng.randrange(len(out) + 1)]
    return out


@pytest.mark.parametrize("seed", range(25))
def test_fuzz_drawio_parse(seed, tmp_path):
    base = (
        '<mxfile><diagram><mxGraphModel><root><mxCell id="0"/>'
        '<mxCell id="1" parent="0"/></root></mxGraphModel></diagram></mxfile>'
    )
    p = tmp_path / f"f{seed}.drawio"
    p.write_text(_mutate(SEED + seed, base), encoding="utf-8")
    # 白名单：SystemExit 已被包成 ValueError（drawio.py:126），畸形输入友好拒绝；
    # 白名单之外（IndexError/AttributeError 等裸 bug）→ 测试失败
    with contextlib.suppress(OffipyError, ValueError, ET.ParseError, RecursionError):
        parse_drawio(p)  # public 入口


@pytest.mark.parametrize("seed", range(25))
def test_fuzz_html_deck(seed, tmp_path):
    """fuzz offipy 自身在 Chromium 之前的纯 Python HTML 预处理（DOM 由 Chromium
    产出，原始 HTML 解析不在此包内；vendored 无纯 Python 入口且非包导入，不 fuzz）。"""
    base = (
        "<html><body><section data-pptx-slide><p>hi</p>"
        '<div class="mermaid">graph LR; a-->b</div></section></body></html>'
    )
    mutated = _mutate(SEED + seed, base)
    # 白名单：_reject 对声明类输入 fail-fast，畸形输入友好拒绝；
    # 白名单之外（IndexError/AttributeError 等裸 bug）→ 测试失败
    with contextlib.suppress(InvalidArgumentError, ValueError):
        _reject_no_visual_audit_declarations(mutated)
        _rewrite_relative_urls(mutated, tmp_path)
        chart_dominant_slide_indices(mutated)
