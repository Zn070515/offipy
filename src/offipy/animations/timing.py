"""<p:timing> OOXML 树构建：curated 入场效果 + click/after 触发模型。

效果模板来自 docs/development/animation-templates.md 捕获的 PowerPoint 原生输出
（Task 1）。结构对齐原生：每个动画单元三层 <p:par>，presetID/presetClass/
presetSubtype/nodeType='clickEffect' 落在最内层 <p:cTn>，<p:set> visibility→visible
作起始保持，效果体（<p:animEffect filter> / <p:anim> 位移）作同层 sibling。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

from offipy.exceptions import InvalidArgumentError

if TYPE_CHECKING:
    from lxml import etree


@dataclass
class AnimationUnit:
    """一个动画单元 = 一个 HTML 元素的所有形状（同 elem_id），同 trigger 同时触发。"""

    spids: list[int]
    effect: str
    direction: str
    trigger: str
    duration_ms: int
    delay_ms: int = 0


# 效果 → (presetID, 默认 presetSubtype)。数值来自 Task 1 捕获表。
_PRESET_META = {
    "fade": (10, 0),
    "float_up": (11, 0),
    "fly_in": (2, 4),  # 默认自底部 = 4（捕获）
    "wipe": (19, 10),
    "zoom_in": (20, 0),
    "grow": (13, 16),
}

# animEffect 滤镜效果（无手写 anim，行为由 filter 驱动）
_ANIM_EFFECT_FILTER = {
    "fade": "fade",
    "zoom_in": "wedge",
    "grow": "plus(in)",
}

# fly_in 方向 → (ppt_x 起始, ppt_y 起始)；目标恒为归位 (#ppt_x / #ppt_y)
_FLY_FROM = {
    "bottom": ("#ppt_x", "1+#ppt_h/2"),  # 捕获：y 从屏下
    "top": ("#ppt_x", "-#ppt_h/2"),  # y 从屏上
    "left": ("-#ppt_w/2", "#ppt_y"),  # x 从屏左
    "right": ("1+#ppt_w/2", "#ppt_y"),  # x 从屏右
}
# fly_in 方向 → presetSubtype（bottom=4 捕获验证；left/right/top 为 PowerPoint
# Fly In 方向编码 best-effort，只影响动画窗格标注，实际运动由手写 anim 决定）
_FLY_SUBTYPE = {"bottom": 4, "left": 5, "right": 6, "top": 7}

# wipe：v1 用捕获模板（ppt_w sin 0→1 水平生长）；vertical（bottom/top）镜像成
# ppt_h 生长。左右/上下内部不细分（同模板）。subtype 是 best-effort 标注。
_WIPE_IS_VERTICAL = {"bottom", "top"}
_WIPE_SUBTYPE = {"left": 10, "right": 10, "bottom": 11, "top": 11}


class _IdSeq:
    """timing 树内唯一 id 分配（自 1 起）。"""

    def __init__(self) -> None:
        self.n = 0

    def next(self) -> int:
        self.n += 1
        return self.n


def _cond(delay: str) -> str:
    return f"<p:cond delay='{delay}'/>"


def _set_visible(ids: _IdSeq, spid: int) -> str:
    """原生起始保持：visibility → visible（入场隐藏由 presetClass='entr' 语义保证）。"""
    i = ids.next()
    return (
        "<p:set><p:cBhvr>"
        f"<p:cTn id='{i}' dur='1' fill='hold'><p:stCondLst>{_cond('0')}</p:stCondLst></p:cTn>"
        f"<p:tgtEl><p:spTgt spid='{spid}'/></p:tgtEl>"
        "<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>"
        "</p:cBhvr><p:to><p:strVal val='visible'/></p:to>"
        "</p:set>"
    )


def _num_anim(
    ids: _IdSeq, spid: int, attr: str, from_val: str, to_val: str, duration_ms: int
) -> str:
    """原生 <p:anim calcmode='lin' valueType='num'> 位移（ppt_x/ppt_y/常量维度）。"""
    i = ids.next()
    return (
        f"<p:anim calcmode='lin' valueType='num'>"
        f"<p:cBhvr additive='base'><p:cTn id='{i}' dur='{duration_ms}' fill='hold'/>"
        f"<p:tgtEl><p:spTgt spid='{spid}'/></p:tgtEl>"
        f"<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr>"
        f"<p:tavLst><p:tav tm='0'><p:val><p:strVal val='{from_val}'/></p:val></p:tav>"
        f"<p:tav tm='100000'><p:val><p:strVal val='{to_val}'/></p:val></p:tav></p:tavLst>"
        "</p:anim>"
    )


def _flt_anim(ids: _IdSeq, spid: int, attr: str, from_fmla: str, duration_ms: int) -> str:
    """原生 <p:anim> 浮点缓动动画（wipe 生长维度：fmla sin 曲线 0→1）。"""
    i = ids.next()
    return (
        f"<p:anim calcmode='lin' valueType='num'>"
        f"<p:cBhvr><p:cTn id='{i}' dur='{duration_ms}' fill='hold'/>"
        f"<p:tgtEl><p:spTgt spid='{spid}'/></p:tgtEl>"
        f"<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr>"
        f"<p:tavLst><p:tav tm='0' fmla='{from_fmla}'><p:val><p:fltVal val='0'/></p:val></p:tav>"
        f"<p:tav tm='100000'><p:val><p:fltVal val='1'/></p:val></p:tav></p:tavLst>"
        "</p:anim>"
    )


def _validate_unit(unit: AnimationUnit) -> None:
    """build_timing 前置校验：效果/方向非法抛 InvalidArgumentError（AnimationUnit
    可被直接构造，绕过 spec.py 归一化）。"""
    if unit.effect not in _PRESET_META:
        raise InvalidArgumentError(f"未知动画效果 {unit.effect!r}（目录 {sorted(_PRESET_META)}）")
    if unit.effect == "fly_in" and unit.direction not in _FLY_SUBTYPE:
        raise InvalidArgumentError(
            f"fly_in direction {unit.direction!r} 非法（{sorted(_FLY_SUBTYPE)}）"
        )
    if unit.effect == "wipe" and unit.direction not in _WIPE_SUBTYPE:
        raise InvalidArgumentError(
            f"wipe direction {unit.direction!r} 非法（{sorted(_WIPE_SUBTYPE)}）"
        )


def _effect_body(ids: _IdSeq, spid: int, unit: AnimationUnit) -> str:
    if unit.effect in _ANIM_EFFECT_FILTER:
        i = ids.next()
        flt = _ANIM_EFFECT_FILTER[unit.effect]
        return (
            f"<p:animEffect transition='in' filter='{flt}'>"
            f"<p:cBhvr><p:cTn id='{i}' dur='{unit.duration_ms}'/>"
            f"<p:tgtEl><p:spTgt spid='{spid}'/></p:tgtEl></p:cBhvr>"
            "</p:animEffect>"
        )
    if unit.effect == "float_up":
        return ""  # presetID=11 运行时解释上浮，XML 无 anim 元素（捕获）
    if unit.effect == "fly_in":
        x0, y0 = _FLY_FROM[unit.direction]
        return _num_anim(ids, spid, "ppt_x", x0, "#ppt_x", unit.duration_ms) + _num_anim(
            ids, spid, "ppt_y", y0, "#ppt_y", unit.duration_ms
        )
    if unit.effect == "wipe":
        if unit.direction in _WIPE_IS_VERTICAL:
            return _num_anim(ids, spid, "ppt_w", "#ppt_w", "#ppt_w", unit.duration_ms) + _flt_anim(
                ids, spid, "ppt_h", "#ppt_h*sin(2.5*pi*$)", unit.duration_ms
            )
        return _flt_anim(ids, spid, "ppt_w", "#ppt_w*sin(2.5*pi*$)", unit.duration_ms) + _num_anim(
            ids, spid, "ppt_h", "#ppt_h", "#ppt_h", unit.duration_ms
        )
    raise InvalidArgumentError(f"未实现的动画效果: {unit.effect}")


def _unit_xml(ids: _IdSeq, unit: AnimationUnit, start_delay_ms: int) -> str:
    """一个动画单元：三层 <p:par>，click 用 delay='indefinite'，after 用累进毫秒。"""
    start = _cond("indefinite") if unit.trigger == "click" else _cond(str(max(0, start_delay_ms)))
    preset_id, default_subtype = _PRESET_META[unit.effect]
    if unit.effect == "fly_in":
        subtype = _FLY_SUBTYPE[unit.direction]
    elif unit.effect == "wipe":
        subtype = _WIPE_SUBTYPE[unit.direction]
    else:
        subtype = default_subtype
    bodies = "".join(_set_visible(ids, spid) + _effect_body(ids, spid, unit) for spid in unit.spids)
    outer = ids.next()
    inner = ids.next()
    eff = ids.next()
    return (
        f"<p:par><p:cTn id='{outer}' fill='hold'><p:stCondLst>{start}</p:stCondLst>"
        "<p:childTnLst><p:par>"
        f"<p:cTn id='{inner}' fill='hold'><p:stCondLst>{_cond('0')}</p:stCondLst>"
        "<p:childTnLst><p:par>"
        f"<p:cTn id='{eff}' presetID='{preset_id}' presetClass='entr' "
        f"presetSubtype='{subtype}' fill='hold' grpId='0' nodeType='clickEffect' nodePh='1'>"
        f"<p:stCondLst>{_cond('0')}</p:stCondLst>"
        f"<p:endCondLst><p:cond evt='begin' delay='0'><p:tn val='{eff}'/></p:cond></p:endCondLst>"
        f"<p:childTnLst>{bodies}</p:childTnLst>"
        "</p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>"
    )


def build_timing(units: list[AnimationUnit]) -> etree._Element | None:
    """构建 <p:timing>。units 为空 → None。"""
    if not units:
        return None
    ids = _IdSeq()
    root_id = ids.next()
    main_id = ids.next()

    # after 触发：delay 累进 = 前一个单元 (duration + delay)，再叠当前 delay
    cumulative = 0
    parts: list[str] = []
    blds: list[str] = []
    for u in units:
        _validate_unit(u)
        start_ms = cumulative + u.delay_ms if u.trigger == "after" else 0
        parts.append(_unit_xml(ids, u, start_ms))
        blds.extend(f"<p:bldP spid='{spid}' grpId='0'/>" for spid in u.spids)
        if u.trigger == "after":
            cumulative = start_ms + u.duration_ms

    xml = (
        f"<p:timing {nsdecls('p')}>"
        "<p:tnLst><p:par>"
        f"<p:cTn id='{root_id}' dur='indefinite' restart='never' nodeType='tmRoot'>"
        "<p:childTnLst>"
        "<p:seq concurrent='1' nextAc='seek'>"
        f"<p:cTn id='{main_id}' dur='indefinite' nodeType='mainSeq'>"
        f"<p:childTnLst>{''.join(parts)}</p:childTnLst>"
        "</p:cTn>"
        "<p:prevCondLst><p:cond evt='onPrev' delay='0'>"
        "<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>"
        "<p:nextCondLst><p:cond evt='onNext' delay='0'>"
        "<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>"
        "</p:seq>"
        "</p:childTnLst>"
        "</p:cTn>"
        "</p:par></p:tnLst>"
        f"<p:bldLst>{''.join(blds)}</p:bldLst>"
        "</p:timing>"
    )
    return parse_xml(xml)
