"""规则级反馈学习（v2）测试：JSONL 落盘/加载、按 (profile, rule) 聚合、apply_feedback。"""

import json

import pytest

from offipy.art import feedback
from offipy.art.feedback import (
    ART_FEEDBACK_FILE,
    ArtFeedbackRecord,
    append,
    apply_feedback,
    load_records,
    recommend_adjustments,
    record_file,
)
from offipy.art.profiles import (
    RULE_DIMENSIONS,
    RULE_LOW_CONTRAST,
    RULE_TINY_TEXT,
    RULE_TITLE_TOO_SMALL,
    ArtProfile,
    get_profile,
)
from offipy.audit import Severity
from offipy.exceptions import InvalidArgumentError


def _write_line(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(data, ensure_ascii=False) + "\n")


# ---------------------------------------------------------------------------
# append / load roundtrip
# ---------------------------------------------------------------------------


def test_append_and_load_roundtrip(tmp_path):
    path = append(
        "balanced",
        RULE_TITLE_TOO_SMALL,
        "fixed",
        Severity.HIGH,
        slide_index=2,
        message="标题太小",
        source="deck.html",
        feedback_dir=tmp_path,
    )
    assert path == tmp_path / ART_FEEDBACK_FILE
    records = load_records(tmp_path)
    assert len(records) == 1
    rec = records[0]
    assert rec.profile == "balanced"
    assert rec.rule_id == RULE_TITLE_TOO_SMALL
    assert rec.dimension == RULE_DIMENSIONS[RULE_TITLE_TOO_SMALL] == "hierarchy"
    assert rec.severity is Severity.HIGH
    assert rec.action == "fixed"
    assert rec.slide_index == 2
    assert rec.message == "标题太小"
    assert rec.source == "deck.html"
    assert rec.ts  # 自动打时间戳


def test_append_creates_parent_dir(tmp_path):
    target = tmp_path / "nested" / "dir"
    append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=target)
    assert record_file(target).exists()


def test_append_loads_missing_file_empty(tmp_path):
    assert load_records(tmp_path / "nonexistent") == []


def test_dimension_derived_not_trusted(tmp_path):
    """dimension 由 rule_id 派生，append 不接受调用方传入。"""
    append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.LOW, feedback_dir=tmp_path)
    rec = load_records(tmp_path)[0]
    assert rec.dimension == "hierarchy"


# ---------------------------------------------------------------------------
# severity / slide_index serialization
# ---------------------------------------------------------------------------


def test_severity_name_and_int_both_load(tmp_path):
    f = tmp_path / ART_FEEDBACK_FILE
    base = {
        "ts": "2026-01-01T00:00:00+00:00",
        "profile": "balanced",
        "rule_id": RULE_TITLE_TOO_SMALL,
        "dimension": "hierarchy",
        "action": "fixed",
        "slide_index": None,
        "message": "",
        "source": "",
    }
    _write_line(f, {**base, "severity": "HIGH"})
    _write_line(f, {**base, "severity": 3})
    records = load_records(tmp_path)
    assert len(records) == 2
    assert records[0].severity is Severity.HIGH
    assert records[1].severity is Severity.HIGH


def test_slide_index_none_vs_int_persists(tmp_path):
    append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    append(
        "balanced",
        RULE_TITLE_TOO_SMALL,
        "fixed",
        Severity.MID,
        slide_index=5,
        feedback_dir=tmp_path,
    )
    lines = (tmp_path / ART_FEEDBACK_FILE).read_text(encoding="utf-8").splitlines()
    assert '"slide_index": null' in lines[0]
    assert '"slide_index": 5' in lines[1]
    records = load_records(tmp_path)
    assert records[0].slide_index is None
    assert records[1].slide_index == 5


def test_to_dict_key_order_and_always_present_fields(tmp_path):
    rec = ArtFeedbackRecord(
        ts="2026-01-01T00:00:00+00:00",
        profile="balanced",
        rule_id=RULE_TITLE_TOO_SMALL,
        dimension="hierarchy",
        severity=Severity.MID,
        action="ignored",
    )
    d = rec.to_dict()
    assert list(d) == [
        "ts",
        "profile",
        "rule_id",
        "dimension",
        "severity",
        "action",
        "slide_index",
        "message",
        "source",
    ]
    assert d["severity"] == "MID"
    assert d["slide_index"] is None
    assert d["message"] == ""
    assert d["source"] == ""


# ---------------------------------------------------------------------------
# append validation
# ---------------------------------------------------------------------------


def test_append_empty_profile_raises(tmp_path):
    with pytest.raises(InvalidArgumentError):
        append("", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)


def test_append_bad_action_raises(tmp_path):
    with pytest.raises(InvalidArgumentError):
        append("balanced", RULE_TITLE_TOO_SMALL, "wat", Severity.MID, feedback_dir=tmp_path)


def test_append_unknown_rule_raises(tmp_path):
    with pytest.raises(InvalidArgumentError):
        append("balanced", "art.nope", "fixed", Severity.MID, feedback_dir=tmp_path)


def test_append_non_severity_raises(tmp_path):
    with pytest.raises(InvalidArgumentError):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", "HIGH", feedback_dir=tmp_path)


@pytest.mark.parametrize("bad", [0, -1, 1.5, "2"])
def test_append_invalid_slide_index_raises(tmp_path, bad):
    with pytest.raises(InvalidArgumentError):
        append(
            "balanced",
            RULE_TITLE_TOO_SMALL,
            "fixed",
            Severity.MID,
            slide_index=bad,
            feedback_dir=tmp_path,
        )


# ---------------------------------------------------------------------------
# load_records skips bad lines
# ---------------------------------------------------------------------------


def test_load_skips_corrupt_json(tmp_path):
    _write_line(tmp_path / ART_FEEDBACK_FILE, "not-json")
    append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    records = load_records(tmp_path)
    assert len(records) == 1
    assert records[0].rule_id == RULE_TITLE_TOO_SMALL


def _good_line():
    return {
        "ts": "2026-01-01T00:00:00+00:00",
        "profile": "balanced",
        "rule_id": RULE_TITLE_TOO_SMALL,
        "dimension": "hierarchy",
        "severity": "MID",
        "action": "fixed",
        "slide_index": None,
        "message": "",
        "source": "",
    }


@pytest.mark.parametrize(
    "mutate",
    [
        lambda d: d.update(action="wat"),
        lambda d: d.update(rule_id="art.nope"),
        lambda d: d.update(severity="NOPE"),
        lambda d: d.update(severity=99),
        lambda d: d.update(slide_index=0),
        lambda d: d.update(slide_index=-3),
    ],
)
def test_load_skips_each_bad_line_keeps_good(tmp_path, mutate):
    f = tmp_path / ART_FEEDBACK_FILE
    bad = _good_line()
    mutate(bad)
    _write_line(f, bad)
    good = _good_line()
    _write_line(f, good)
    records = load_records(tmp_path)
    assert len(records) == 1
    assert records[0].rule_id == RULE_TITLE_TOO_SMALL


# ---------------------------------------------------------------------------
# recommend_adjustments aggregation math
# ---------------------------------------------------------------------------


def test_den_less_than_3_no_recommendation(tmp_path):
    append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    append("balanced", RULE_TITLE_TOO_SMALL, "accepted", Severity.MID, feedback_dir=tmp_path)
    assert recommend_adjustments("balanced", feedback_dir=tmp_path) == {}


def test_three_fixed_plus_one(tmp_path):
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    assert recommend_adjustments("balanced", feedback_dir=tmp_path) == {RULE_TITLE_TOO_SMALL: 1}


def test_three_accepted_minus_one(tmp_path):
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "accepted", Severity.MID, feedback_dir=tmp_path)
    assert recommend_adjustments("balanced", feedback_dir=tmp_path) == {RULE_TITLE_TOO_SMALL: -1}


def test_split_even_no_recommendation(tmp_path):
    for _ in range(2):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
        append("balanced", RULE_TITLE_TOO_SMALL, "accepted", Severity.MID, feedback_dir=tmp_path)
    assert recommend_adjustments("balanced", feedback_dir=tmp_path) == {}


def test_three_fixed_two_accepted_boundary_plus_one(tmp_path):
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    for _ in range(2):
        append("balanced", RULE_TITLE_TOO_SMALL, "accepted", Severity.MID, feedback_dir=tmp_path)
    assert recommend_adjustments("balanced", feedback_dir=tmp_path) == {RULE_TITLE_TOO_SMALL: 1}


def test_two_fixed_three_accepted_boundary_minus_one(tmp_path):
    for _ in range(2):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "accepted", Severity.MID, feedback_dir=tmp_path)
    assert recommend_adjustments("balanced", feedback_dir=tmp_path) == {RULE_TITLE_TOO_SMALL: -1}


def test_ignored_neutral_below_denominator(tmp_path):
    append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    append("balanced", RULE_TITLE_TOO_SMALL, "ignored", Severity.MID, feedback_dir=tmp_path)
    append("balanced", RULE_TITLE_TOO_SMALL, "ignored", Severity.MID, feedback_dir=tmp_path)
    # den = 2（ignored 不进分母）< 3 → 无建议
    assert recommend_adjustments("balanced", feedback_dir=tmp_path) == {}


def test_ignored_excluded_from_denominator(tmp_path):
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "ignored", Severity.MID, feedback_dir=tmp_path)
    # den = 3，fixed/den = 1.0 → +1
    assert recommend_adjustments("balanced", feedback_dir=tmp_path) == {RULE_TITLE_TOO_SMALL: 1}


def test_only_minus_one_plus_one_entries_returned(tmp_path):
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    for _ in range(3):
        append("balanced", RULE_TINY_TEXT, "accepted", Severity.MID, feedback_dir=tmp_path)
    # 4 fixed + 4 accepted → 无建议，不出现在结果里
    for _ in range(4):
        append("balanced", RULE_LOW_CONTRAST, "fixed", Severity.MID, feedback_dir=tmp_path)
        append("balanced", RULE_LOW_CONTRAST, "accepted", Severity.MID, feedback_dir=tmp_path)
    result = recommend_adjustments("balanced", feedback_dir=tmp_path)
    assert result == {RULE_TITLE_TOO_SMALL: 1, RULE_TINY_TEXT: -1}
    assert set(result.values()) <= {-1, 1}
    assert 0 not in result.values()


def test_profile_isolation(tmp_path):
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    for _ in range(3):
        append("academic", RULE_TITLE_TOO_SMALL, "accepted", Severity.MID, feedback_dir=tmp_path)
    assert recommend_adjustments("balanced", feedback_dir=tmp_path) == {RULE_TITLE_TOO_SMALL: 1}
    assert recommend_adjustments("academic", feedback_dir=tmp_path) == {RULE_TITLE_TOO_SMALL: -1}


# ---------------------------------------------------------------------------
# apply_feedback
# ---------------------------------------------------------------------------


def test_apply_feedback_string_profile(tmp_path):
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    p = apply_feedback("balanced", feedback_dir=tmp_path)
    assert p.name == "balanced"
    assert p.feedback_severity_adjustments == {RULE_TITLE_TOO_SMALL: 1}
    assert p.feedback_severity_adjustments == recommend_adjustments(
        "balanced", feedback_dir=tmp_path
    )
    assert p.severity_overrides == {}
    assert p.confidence_overrides == {}
    assert p is not get_profile("balanced")


def test_apply_feedback_custom_profile_preserves_overrides(tmp_path):
    for _ in range(3):
        append("custom", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    base = ArtProfile(
        name="custom",
        severity_overrides={RULE_TITLE_TOO_SMALL: Severity.HIGH},
        confidence_overrides={RULE_TITLE_TOO_SMALL: 0.9},
    )
    p = apply_feedback(base, feedback_dir=tmp_path)
    assert p.name == "custom"
    assert p.feedback_severity_adjustments == {RULE_TITLE_TOO_SMALL: 1}
    assert p.severity_overrides == {RULE_TITLE_TOO_SMALL: Severity.HIGH}
    assert p.confidence_overrides == {RULE_TITLE_TOO_SMALL: 0.9}
    # 输入 profile 未被修改
    assert base.feedback_severity_adjustments == {}
    assert p is not base


def test_apply_feedback_does_not_mutate_builtin(tmp_path):
    for _ in range(3):
        append("balanced", RULE_TITLE_TOO_SMALL, "fixed", Severity.MID, feedback_dir=tmp_path)
    apply_feedback("balanced", feedback_dir=tmp_path)
    # 内置 profile 是共享对象，必须保持原样
    original = get_profile("balanced")
    assert original.feedback_severity_adjustments == {}
    assert original.severity_overrides == {}


def test_record_file_default_location():
    assert record_file() == feedback.DEFAULT_DIR / ART_FEEDBACK_FILE
