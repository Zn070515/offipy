"""animations/timing.py：<p:timing> OOXML 树构建。"""

import pytest

from offipy.animations.timing import AnimationUnit, build_timing
from offipy.exceptions import InvalidArgumentError


def _qn(tag):
    return tag.split("}")[-1]


def _children_names(el):
    return [_qn(c.tag) for c in el]


def test_build_timing_root_structure():
    unit = AnimationUnit(
        spids=[2], effect="fade", direction="bottom", trigger="click", duration_ms=500, delay_ms=0
    )
    timing = build_timing([unit])
    assert _qn(timing.tag) == "timing"
    seq = timing.findall(".//{*}seq")
    assert len(seq) == 1
    assert seq[0].get("concurrent") == "1"


def test_build_timing_click_uses_indefinite_cond():
    unit = AnimationUnit(
        spids=[2], effect="fade", direction="bottom", trigger="click", duration_ms=500, delay_ms=0
    )
    timing = build_timing([unit])
    conds = timing.findall(".//{*}cond")
    # click：起始 stCondLst 的 cond delay='indefinite'
    assert any(c.get("delay") == "indefinite" for c in conds)


def test_build_timing_after_uses_cumulative_delay():
    units = [
        AnimationUnit(
            spids=[2],
            effect="fade",
            direction="bottom",
            trigger="after",
            duration_ms=500,
            delay_ms=0,
        ),
        AnimationUnit(
            spids=[3],
            effect="fade",
            direction="bottom",
            trigger="after",
            duration_ms=400,
            delay_ms=200,
        ),
    ]
    timing = build_timing(units)
    conds = [c.get("delay") for c in timing.findall(".//{*}cond")]
    # after 链：第一单元 start = 0（无前驱）+ 自身 delay 0 = 0；
    # 第二单元 start = 前单元累计(0+500) + 本 delay 200 = 700
    assert "0" in conds
    assert "700" in conds


def test_build_timing_after_delay_accumulates():
    unit = AnimationUnit(
        spids=[2], effect="fade", direction="bottom", trigger="after", duration_ms=500, delay_ms=200
    )
    timing = build_timing([unit])
    conds = [c.get("delay") for c in timing.findall(".//{*}cond")]
    # 第一个 after 单元 start = 自身 delay（无前驱累计），非 500+200
    assert "200" in conds


def test_build_timing_fade_effect_body():
    unit = AnimationUnit(
        spids=[2], effect="fade", direction="bottom", trigger="click", duration_ms=500, delay_ms=0
    )
    timing = build_timing([unit])
    # fade 效果体：p:animEffect filter='fade' + 起始 p:set visibility→visible
    # （Task 1 捕获校准：真实 PowerPoint 写 visible 而非 spec 简版的 hidden）
    anim_effects = timing.findall(".//{*}animEffect")
    assert len(anim_effects) == 1
    assert anim_effects[0].get("filter") == "fade"
    assert anim_effects[0].get("transition") == "in"
    sets = timing.findall(".//{*}set")
    assert len(sets) == 1
    str_val = sets[0].findall(".//{*}strVal")
    assert any(v.get("val") == "visible" for v in str_val)


def test_build_timing_spid_targets():
    unit = AnimationUnit(
        spids=[2], effect="fade", direction="bottom", trigger="click", duration_ms=500, delay_ms=0
    )
    timing = build_timing([unit])
    spids = [t.get("spid") for t in timing.findall(".//{*}spTgt")]
    assert spids == ["2", "2"]  # set + animEffect 各一


def test_build_timing_multi_shape_unit_same_group():
    # 同一 elem_id 的多个形状（fill + 边框线）同 trigger 同时触发：
    # 单元内每个形状一组 set+effect
    unit = AnimationUnit(
        spids=[2, 5],
        effect="fade",
        direction="bottom",
        trigger="click",
        duration_ms=500,
        delay_ms=0,
    )
    timing = build_timing([unit])
    spids = [t.get("spid") for t in timing.findall(".//{*}spTgt")]
    assert sorted(spids) == ["2", "2", "5", "5"]
    # 两组效果体，各带一个起始 hidden
    sets = timing.findall(".//{*}set")
    assert len(sets) == 2


def test_build_timing_ids_unique():
    units = [
        AnimationUnit(
            spids=[2],
            effect="fade",
            direction="bottom",
            trigger="click",
            duration_ms=500,
            delay_ms=0,
        ),
        AnimationUnit(
            spids=[3],
            effect="fade",
            direction="bottom",
            trigger="after",
            duration_ms=400,
            delay_ms=100,
        ),
    ]
    timing = build_timing(units)
    ids = [c.get("id") for c in timing.findall(".//*[@id]")]
    assert len(ids) == len(set(ids))


def test_build_timing_duration_ms():
    unit = AnimationUnit(
        spids=[2], effect="fade", direction="bottom", trigger="click", duration_ms=800, delay_ms=0
    )
    timing = build_timing([unit])
    ctn = timing.find(".//{*}animEffect/{*}cBhvr/{*}cTn")
    assert ctn.get("dur") == "800"


def test_build_timing_fly_in_native_structure():
    # Task 1 捕获校准：presetID/presetClass/presetSubtype 落在 clickEffect cTn 上，
    # 不在 <p:anim> 上；运动由两个 <p:anim>（ppt_x/ppt_y）手写位移驱动。
    unit = AnimationUnit(
        spids=[2], effect="fly_in", direction="left", trigger="click", duration_ms=500, delay_ms=0
    )
    timing = build_timing([unit])
    eff = timing.find(".//{*}cTn[@nodeType='clickEffect']")
    assert eff is not None
    assert eff.get("presetID") == "2"
    assert eff.get("presetClass") == "entr"
    assert eff.get("presetSubtype") == "5"  # left（bottom=4 已捕获，left/right/top 按约定）
    anims = timing.findall(".//{*}anim")
    assert len(anims) == 2
    attrs = [a.findall(".//{*}attrName")[0].text for a in anims]
    assert set(attrs) == {"ppt_x", "ppt_y"}


def test_build_timing_zoom_grow_use_anim_effect_filters():
    # Task 1 捕获校准：zoom_in/grow 是 animEffect 滤镜（wedge / plus(in)），不是 animScale
    zoom = build_timing(
        [
            AnimationUnit(
                spids=[2],
                effect="zoom_in",
                direction="bottom",
                trigger="click",
                duration_ms=500,
                delay_ms=0,
            )
        ]
    )
    grow = build_timing(
        [
            AnimationUnit(
                spids=[2],
                effect="grow",
                direction="bottom",
                trigger="click",
                duration_ms=500,
                delay_ms=0,
            )
        ]
    )
    assert zoom.findall(".//{*}animEffect")[0].get("filter") == "wedge"
    assert grow.findall(".//{*}animEffect")[0].get("filter") == "plus(in)"


def test_build_timing_float_up_has_no_anim_body():
    # Task 1 捕获校准：float_up 只有 set + presetID=11，无 <p:anim>/<p:animEffect>
    unit = AnimationUnit(
        spids=[2],
        effect="float_up",
        direction="bottom",
        trigger="click",
        duration_ms=500,
        delay_ms=0,
    )
    timing = build_timing([unit])
    assert not timing.findall(".//{*}anim")
    assert not timing.findall(".//{*}animEffect")
    assert len(timing.findall(".//{*}set")) == 1
    eff = timing.find(".//{*}cTn[@nodeType='clickEffect']")
    assert eff.get("presetID") == "11"


def test_build_timing_bld_list_per_shape():
    unit = AnimationUnit(
        spids=[2, 5],
        effect="fade",
        direction="bottom",
        trigger="click",
        duration_ms=500,
        delay_ms=0,
    )
    timing = build_timing([unit])
    bldp = [b.get("spid") for b in timing.findall(".//{*}bldP")]
    assert sorted(bldp) == ["2", "5"]


def test_build_timing_empty_returns_none():
    assert build_timing([]) is None


def test_build_timing_has_prev_next_conds():
    # PowerPoint 接受的 mainSeq 需要 onPrev/onNext 条件（sldTgt）才能正确前进/后退
    unit = AnimationUnit(
        spids=[2],
        effect="fade",
        direction="bottom",
        trigger="click",
        duration_ms=500,
        delay_ms=0,
    )
    timing = build_timing([unit])
    prev = timing.find(".//{*}prevCondLst")
    nxt = timing.find(".//{*}nextCondLst")
    assert prev is not None and prev.findall(".//{*}sldTgt")
    assert nxt is not None and nxt.findall(".//{*}sldTgt")


def test_build_timing_invalid_input_raises():
    # AnimationUnit 可被直接构造绕过 spec.py 归一化 → 非法效果/方向抛 InvalidArgumentError
    with pytest.raises(InvalidArgumentError):
        build_timing(
            [
                AnimationUnit(
                    spids=[2],
                    effect="wobble",
                    direction="bottom",
                    trigger="click",
                    duration_ms=500,
                    delay_ms=0,
                )
            ]
        )
    with pytest.raises(InvalidArgumentError):
        build_timing(
            [
                AnimationUnit(
                    spids=[2],
                    effect="fly_in",
                    direction="northwest",
                    trigger="click",
                    duration_ms=500,
                    delay_ms=0,
                )
            ]
        )
