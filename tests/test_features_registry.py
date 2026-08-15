"""FEATURES 注册表 + encode_features：扁平标量快照，schema 版本。"""

from offipy.art import encode_features, feature_keys, feature_schema_version
from offipy.art.models import ArtFinding, ArtSlide
from offipy.art.profiles import RULE_NO_FOCUS, RULE_OFF_BALANCE
from offipy.audit import Severity

INPUT_DIM = 59  # 锁定的输入维度：17 rule + 6 dim + 18 measure + 4 scalar + 8 slide + 2 deck


def _finding(**kw):
    base = {
        "rule_id": RULE_NO_FOCUS,
        "dimension": "hierarchy",
        "severity": Severity.HIGH,
        "message": "m",
        "confidence": 0.9,
        "slide_index": 2,
        "details": {},
    }
    base.update(kw)
    return ArtFinding(**base)


def _slide(index=2):
    return ArtSlide(index=index, width=1920.0, height=1080.0, elements=[])


def test_feature_keys_dimension_is_59():
    keys = feature_keys()
    assert len(keys) == INPUT_DIM == 59
    assert len(set(keys)) == len(keys)  # 无重复


def test_encode_flat_scalar_dict_all_keys_present():
    f = _finding()
    enc = encode_features(f, _slide(), deck=None, profile="balanced")
    assert set(enc) == set(feature_keys())
    assert all(isinstance(v, float) for v in enc.values())


def test_rule_onehot_only_matched_rule_is_one():
    enc = encode_features(_finding(rule_id=RULE_NO_FOCUS), _slide())
    assert enc[f"finding.rule_id.{RULE_NO_FOCUS}"] == 1.0
    assert enc[f"finding.rule_id.{RULE_OFF_BALANCE}"] == 0.0


def test_measurement_extracted_from_details():
    enc = encode_features(
        _finding(rule_id=RULE_OFF_BALANCE, details={"balance_dist": 0.4, "ink": 0.8}),
        _slide(),
    )
    assert enc["measure.art.composition.off_balance.balance_dist"] == 0.4
    assert enc["measure.art.composition.off_balance.ink"] == 0.8


def test_measurement_missing_default_zero_for_other_rules():
    # 非本规则的 measurement 特征回退 missing_default(0.0)，维度固定
    enc = encode_features(_finding(rule_id=RULE_NO_FOCUS), _slide())
    assert enc["measure.art.composition.off_balance.balance_dist"] == 0.0


def test_severity_ordinal_and_confidence():
    enc = encode_features(_finding(severity=Severity.MID, confidence=0.4), _slide())
    assert enc["finding.severity_ordinal"] == 2.0
    assert enc["finding.confidence"] == 0.4


def test_profile_onehot():
    enc = encode_features(_finding(), _slide(), profile="academic")
    assert enc["deck.profile.academic"] == 1.0
    assert enc["deck.profile.balanced"] == 0.0


def test_missing_optional_fields_default():
    enc = encode_features(_finding(evidence_reliability=None), _slide())
    assert enc["finding.evidence_reliability"] == 0.0
    assert enc["finding.page_ratio"] == 0.0  # 无 deck → total_slides=0 → extract None → 0.0


def test_schema_version_constant():
    assert feature_schema_version() == "1"


def test_slide_feature_scalarizers():
    enc = encode_features(_finding(), _slide())
    assert "slide.alignment" in enc
    assert "slide.spacing" in enc
    assert "slide.mass" in enc
    assert "slide.density" in enc
    assert "slide.font_hierarchy" in enc
    assert "slide.palette" in enc
    assert "slide.focus" in enc
    assert "slide.page_signature" in enc
