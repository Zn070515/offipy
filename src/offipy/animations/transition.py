"""<p:transition> OOXML 构建：fade/wipe/push/cover 4 种 classic 过渡 + speed。

只用标准 ECMA-376 过渡（PowerPoint/WPS/LibreOffice 兼容）；不做需要
mc:AlternateContent 的 p14 过渡（spec Out of scope）。
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

from offipy.exceptions import InvalidArgumentError

if TYPE_CHECKING:
    from lxml import etree

_SPEED_MAP = {"slow": "slow", "medium": "med", "fast": "fast"}


def build_transition(kind: str, speed: str = "medium") -> etree._Element:
    if speed not in _SPEED_MAP:
        raise InvalidArgumentError(f"未知过渡速度: {speed}（slow/medium/fast）")
    spd = _SPEED_MAP[speed]
    if kind == "fade":
        inner = "<p:fade/>"
    elif kind in {"wipe", "push", "cover"}:
        inner = f"<p:{kind} dir='l'/>"
    else:
        raise InvalidArgumentError(f"未知过渡类型: {kind}（fade/wipe/push/cover）")
    xml = f"<p:transition {nsdecls('p')} spd='{spd}'>{inner}</p:transition>"
    return parse_xml(xml)
