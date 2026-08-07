import pytest

from offipy.art.profiles import (
    ALL_RULES,
    RULE_ACCENT_FLOOD,
    RULE_BACKGROUND_LIKE_AREA,
    RULE_CORNER_CLUSTER,
    RULE_DIMENSIONS,
    RULE_OFF_BALANCE,
    RULE_TITLE_TOO_SMALL,
    ArtProfile,
    get_profile,
    profile_names,
)
from offipy.exceptions import InvalidArgumentError


def test_all_rules_count():
    assert len(ALL_RULES) == 17


def test_rule_dimensions_covers_all_rules():
    assert set(RULE_DIMENSIONS) == set(ALL_RULES)


def test_experimental_rules():
    assert RULE_OFF_BALANCE in get_profile("balanced").experimental_rules
    assert RULE_CORNER_CLUSTER in get_profile("balanced").experimental_rules
    assert RULE_ACCENT_FLOOD in get_profile("balanced").experimental_rules


def test_balanced_defaults():
    p = get_profile("balanced")
    assert p.max_font_families == 3
    assert p.title_drift_tol == 0.15
    assert p.experimental_rules == {
        RULE_OFF_BALANCE,
        RULE_CORNER_CLUSTER,
        RULE_ACCENT_FLOOD,
        "art.hierarchy.no_focus",
        "art.color.no_accent",
        RULE_BACKGROUND_LIKE_AREA,
    }


def test_consulting_tight_typography():
    p = get_profile("consulting")
    assert p.min_contrast == 4.5
    assert p.max_font_families == 2
    assert p.title_drift_tol == 0.10
    assert p.spacing_drift_tol == 0.20


def test_academic_disables_off_balance():
    p = get_profile("academic")
    assert RULE_OFF_BALANCE in p.disabled_rules


def test_event_relaxes_rules():
    p = get_profile("event")
    assert RULE_OFF_BALANCE in p.disabled_rules
    assert RULE_CORNER_CLUSTER in p.disabled_rules
    assert p.max_accent_ratio == 0.7
    assert p.max_font_families == 4
    assert p.title_drift_tol == 0.25


def test_get_profile_unknown_raises():
    with pytest.raises(KeyError):
        get_profile("nope")


def test_background_like_area_is_experimental():
    assert RULE_BACKGROUND_LIKE_AREA in get_profile("balanced").experimental_rules


def test_max_background_like_ratio_per_profile():
    assert get_profile("balanced").max_background_like_ratio == 0.75
    assert get_profile("consulting").max_background_like_ratio == 0.85
    assert get_profile("academic").max_background_like_ratio == 0.75
    assert get_profile("technology").max_background_like_ratio == 0.75
    assert get_profile("event").max_background_like_ratio == 0.90
    assert RULE_BACKGROUND_LIKE_AREA in get_profile("balanced").experimental_rules


def test_profile_names():
    assert set(profile_names()) == {"balanced", "consulting", "academic", "technology", "event"}


# ---------------------------------------------------------------------------
# feedback_severity_adjustments validation
# ---------------------------------------------------------------------------


def test_feedback_severity_adjustments_accepts_minus_one_and_plus_one():
    p = ArtProfile(
        name="x",
        feedback_severity_adjustments={RULE_TITLE_TOO_SMALL: -1},
    )
    assert p.feedback_severity_adjustments == {RULE_TITLE_TOO_SMALL: -1}
    p2 = ArtProfile(
        name="y",
        feedback_severity_adjustments={RULE_TITLE_TOO_SMALL: +1},
    )
    assert p2.feedback_severity_adjustments == {RULE_TITLE_TOO_SMALL: +1}


@pytest.mark.parametrize("bad_delta", [2, 0, -2, 3])
def test_feedback_severity_adjustments_rejects_invalid(bad_delta):
    with pytest.raises(InvalidArgumentError):
        ArtProfile(
            name="x",
            feedback_severity_adjustments={RULE_TITLE_TOO_SMALL: bad_delta},
        )
