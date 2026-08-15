"""配对构建器：同 (rule_id, profile) 组内 fixed>accepted，ignored 排除。"""

from offipy.art.feedback import ArtFeedbackRecord
from offipy.art.profiles import RULE_LOW_CONTRAST, RULE_TITLE_TOO_SMALL
from offipy.audit import Severity
from offipy.feedback.pairs import build_pairs


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
