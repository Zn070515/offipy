"""配对构建器：同 (rule_id, profile) 组内 fixed>accepted，ignored 排除。"""

from offipy.art.feedback import ArtFeedbackRecord
from offipy.art.profiles import RULE_LOW_CONTRAST, RULE_TITLE_TOO_SMALL
from offipy.audit import Severity
from offipy.feedback.pairs import PER_RULE_MIN_BASE, build_pairs, per_rule_diagnosis


def _rec(rule_id, action, profile="balanced", **kw):
    return ArtFeedbackRecord(
        ts="2026-01-01T00:00:00+00:00",
        profile=profile,
        rule_id=rule_id,
        dimension="hierarchy",
        severity=Severity.MID,
        action=action,
        features={"x": 1.0},
        feature_schema_version="1",
        **kw,
    )


def test_pairs_from_same_group():
    recs = [
        _rec(RULE_TITLE_TOO_SMALL, "fixed"),
        _rec(RULE_TITLE_TOO_SMALL, "fixed"),
        _rec(RULE_TITLE_TOO_SMALL, "accepted"),
        _rec(RULE_TITLE_TOO_SMALL, "accepted"),
    ]
    pairs = build_pairs(recs)
    assert len(pairs) == 4  # 2 fixed × 2 accepted
    for f, a in pairs:
        assert f.action == "fixed"
        assert a.action == "accepted"


def test_ignored_excluded():
    recs = [
        _rec(RULE_TITLE_TOO_SMALL, "fixed"),
        _rec(RULE_TITLE_TOO_SMALL, "ignored"),
        _rec(RULE_TITLE_TOO_SMALL, "ignored"),
    ]
    assert build_pairs(recs) == []


def test_profile_and_rule_isolated():
    recs = [
        _rec(RULE_TITLE_TOO_SMALL, "fixed", profile="balanced"),
        _rec(RULE_TITLE_TOO_SMALL, "accepted", profile="academic"),
        _rec(RULE_LOW_CONTRAST, "fixed", profile="balanced"),
        _rec(RULE_LOW_CONTRAST, "accepted", profile="balanced"),
    ]
    pairs = build_pairs(recs)
    # 只 (title_too_small, fixed, balanced) × (low_contrast, accepted, balanced) 跨 rule 不算
    assert len(pairs) == 1
    assert pairs[0][1].rule_id == RULE_LOW_CONTRAST


def test_balanced_group_no_pairs():
    # 单组只有 fixed 或只有 accepted → 无配对
    assert build_pairs([_rec(RULE_TITLE_TOO_SMALL, "fixed")]) == []
    assert build_pairs([_rec(RULE_TITLE_TOO_SMALL, "accepted")]) == []


# ---- per_rule_diagnosis (#117) ----


def test_single_direction_rule_reports_suggest():
    """固定单方向（fixed 有、accepted 无）→ pairs=0、single_direction、suggest>0。"""
    recs = [_rec(RULE_TITLE_TOO_SMALL, "fixed") for _ in range(3)]
    out = per_rule_diagnosis(recs, min_pairs=50)
    d = out[RULE_TITLE_TOO_SMALL]
    assert d["pairs"] == 0
    assert d["single_direction"] is True
    assert d["suggest"] > 0  # 补另一侧（accepted）


def test_bidirectional_but_short_suggests_more():
    """双方向都有但不足（2×2=4 < 单规则 per_rule_min=50）→ suggest>0、非单方向。"""
    recs = [_rec(RULE_TITLE_TOO_SMALL, "fixed") for _ in range(2)] + [
        _rec(RULE_TITLE_TOO_SMALL, "accepted") for _ in range(2)
    ]
    d = per_rule_diagnosis(recs, min_pairs=50)[RULE_TITLE_TOO_SMALL]
    assert d["pairs"] == 4
    assert d["single_direction"] is False
    assert d["suggest"] > 0
    assert d["fixed"] == 2
    assert d["accepted"] == 2


def test_multi_rule_per_rule_min_floor():
    """多规则均摊 per_rule_min=max(min_pairs//n_rules, 10)：25 pairs 恰好够 → suggest=0。"""
    recs = [
        _rec(rule, action)
        for rule in (RULE_TITLE_TOO_SMALL, RULE_LOW_CONTRAST)
        for action in ("fixed", "accepted")
        for _ in range(5)
    ]
    out = per_rule_diagnosis(recs, min_pairs=50)
    assert len(out) == 2
    for d in out.values():
        assert d["pairs"] == 25  # per_rule_min = max(50//2, 10) = 25，25 不 < 25
        assert d["suggest"] == 0


def test_per_rule_min_base_floor_applies():
    """min_pairs//n_rules 小于下限时取 10：9 pairs < 10 → suggest>0（若下限失效则为 0）。"""
    recs = [
        _rec(rule, action)
        for rule in (RULE_TITLE_TOO_SMALL, RULE_LOW_CONTRAST)
        for action in ("fixed", "accepted")
        for _ in range(3)
    ]
    out = per_rule_diagnosis(recs, min_pairs=4)
    assert len(out) == 2
    d = out[RULE_TITLE_TOO_SMALL]
    assert d["pairs"] == 9  # max(4//2, PER_RULE_MIN_BASE)=10，9 < 10
    assert PER_RULE_MIN_BASE == 10
    assert d["suggest"] > 0


def test_diagnosis_shape_and_ignored_excluded():
    """返回 key 完整（fixed/accepted/pairs/single_direction/suggest），ignored 不计数。"""
    recs = (
        [_rec(RULE_TITLE_TOO_SMALL, "fixed") for _ in range(2)]
        + [_rec(RULE_TITLE_TOO_SMALL, "accepted")]
        + [_rec(RULE_TITLE_TOO_SMALL, "ignored") for _ in range(7)]
    )
    out = per_rule_diagnosis(recs, min_pairs=50)
    assert set(out) == {RULE_TITLE_TOO_SMALL}
    d = out[RULE_TITLE_TOO_SMALL]
    assert set(d) == {"fixed", "accepted", "pairs", "single_direction", "suggest"}
    assert d["fixed"] == 2
    assert d["accepted"] == 1  # ignored 不计入
    assert d["pairs"] == 2
    assert isinstance(d["suggest"], int)
    assert isinstance(d["single_direction"], bool)


def test_cross_profile_fixed_accepted_do_not_pair():
    """同 rule 跨 profile 的 fixed/accepted 不配对（与 build_pairs 语义一致）→ pairs=0。

    5 fixed/balanced + 5 accepted/academic 实际 0 训练对，不得虚报 pairs=25。"""
    recs = [_rec(RULE_TITLE_TOO_SMALL, "fixed", profile="balanced") for _ in range(5)] + [
        _rec(RULE_TITLE_TOO_SMALL, "accepted", profile="academic") for _ in range(5)
    ]
    assert build_pairs(recs) == []  # 语义锚：build_pairs 确实不跨 profile 配对
    d = per_rule_diagnosis(recs, min_pairs=50)[RULE_TITLE_TOO_SMALL]
    assert d["fixed"] == 5
    assert d["accepted"] == 5
    assert d["pairs"] == 0
    assert d["single_direction"] is True
    assert d["suggest"] > 0
