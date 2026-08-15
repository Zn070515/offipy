"""FEATURES 注册表 + encode_features：扁平标量快照，schema 版本。"""

from offipy.art import encode_features, feature_keys, feature_schema_version
from offipy.art.models import ArtColor, ArtElement, ArtFinding, ArtScene, ArtSlide
from offipy.art.profiles import (
    RULE_CORNER_CLUSTER,
    RULE_MANY_FAMILIES,
    RULE_NO_FOCUS,
    RULE_OFF_BALANCE,
    RULE_SPACING_DRIFT,
)
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


def _real_slide():
    """带真实元素的一页：3 个左对齐 body（红色前景）+ 1 个 title，全部有字号。

    手工推导的期望值（见 test_slide_feature_scalarizers_values）：
    alignment.lines=3（x/cx/right 三条竖线）、mass.ink=0.46、density.sum_area=0.045、
    font_hierarchy.ratio=2.0、palette.accent_ratio=1.0、focus.ratio=2.0、
    page_signature.role=content（无显式 role 标记）。
    """
    title = ArtElement(
        element_id="t",
        kind="text",
        role="title",
        x=0.1,
        y=0.05,
        width=0.5,
        height=0.03,
        slide_index=2,
        text="Title",
        font_size_norm=0.04,
    )
    bodies = [
        ArtElement(
            element_id=f"b{i}",
            kind="text",
            role="body",
            x=0.1,
            y=y,
            width=0.2,
            height=0.05,
            slide_index=2,
            text=text,
            font_size_norm=0.02,
            foreground=ArtColor(200, 0, 0),
        )
        for i, (y, text) in enumerate(((0.1, "Hello"), (0.2, "World"), (0.3, "Foo")), start=1)
    ]
    return ArtSlide(index=2, width=1920.0, height=1080.0, elements=[title, *bodies])


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


def test_page_ratio_positive_path():
    # 5 页场景：idx/total 正分支（对照 test_missing_optional_fields_default 的缺失路径）
    deck = ArtScene(slides=[_slide(index=i) for i in range(1, 6)])
    enc = encode_features(_finding(slide_index=2), _slide(index=2), deck=deck)
    assert enc["finding.page_ratio"] == 2.0 / 5.0
    # 中间页：ratio 严格 < 1.0
    enc_mid = encode_features(_finding(slide_index=4), _slide(index=4), deck=deck)
    assert enc_mid["finding.page_ratio"] == 4.0 / 5.0 < 1.0
    # 最后一页：ratio == total/total == 1.0，有界 ≤ 1.0，不会超界
    enc_last = encode_features(_finding(slide_index=5), _slide(index=5), deck=deck)
    assert enc_last["finding.page_ratio"] == 1.0
    assert 0.0 < enc_last["finding.page_ratio"] <= 1.0


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


def test_slide_feature_scalarizers_values():
    # 带真实元素的一页：标量化器必须提取真实值，而不是全部走 missing_default
    enc = encode_features(_finding(), _real_slide())
    assert enc["slide.alignment"] == 3.0
    assert enc["slide.spacing"] == 0.0
    assert enc["slide.mass"] == 0.46
    assert enc["slide.density"] == 0.045
    assert enc["slide.font_hierarchy"] == 2.0
    assert enc["slide.palette"] == 1.0
    assert enc["slide.focus"] == 2.0
    assert enc["slide.page_signature"] == 1.0  # role=content


def test_measurement_len_many_families():
    enc = encode_features(
        _finding(rule_id=RULE_MANY_FAMILIES, details={"families": [f"F{i}" for i in range(3)]}),
        _slide(),
    )
    assert enc["measure.art.typography.many_families.families"] == 3.0


def test_measurement_dict_max_corner_cluster():
    enc = encode_features(
        _finding(
            rule_id=RULE_CORNER_CLUSTER,
            details={"quadrants": {"tl": 0.8, "tr": 0.1, "bl": 0.05, "br": 0.05}},
        ),
        _slide(),
    )
    assert enc["measure.art.composition.corner_cluster.quadrants"] == 0.8


def test_measurement_nested_spacing_drift():
    enc = encode_features(
        _finding(
            rule_id=RULE_SPACING_DRIFT,
            details={
                "horizontal": {"gaps": [0.1, 0.2], "drift_count": 1, "max_drift_ratio": 0.5},
                "vertical": {"gaps": [0.1, 0.1], "drift_count": 0, "max_drift_ratio": 0.0},
            },
        ),
        _slide(),
    )
    assert enc["measure.art.composition.spacing_drift.horizontal"] == 0.5
    assert enc["measure.art.composition.spacing_drift.vertical"] == 0.0


def test_page_signature_unknown_role_other_bucket():
    from offipy.art.features_registry import _slide_page_signature

    # 未知 role → 显式 catch-all 桶 len(roles)=6，不再坍缩成 cover 的 2.0
    assert (
        _slide_page_signature({"slide_features": {"page_signature": {"role": "appendix"}}}) == 6.0
    )
    # 已知 role → sorted 序号（closing=0, content=1, cover=2, data=3, gallery=4, section=5）
    assert _slide_page_signature({"slide_features": {"page_signature": {"role": "content"}}}) == 1.0
    # 无 slide 数据 → None → missing_default
    assert _slide_page_signature({"slide_features": {}}) is None


def test_page_signature_missing_default_is_other_bucket():
    # 行为变化（Fix 1）：slide=None → page_signature 回退到「其他」桶 6.0，
    # 不再是旧值 2.0（旧值恰好等于 cover 的序号，会让未知 role 坍缩成 cover）
    enc = encode_features(_finding(), slide=None)
    assert enc["slide.page_signature"] == 6.0


def test_feature_keys_layout_tripwire():
    from offipy.art.features_registry import _SLIDE_ROLE_ORDER, _SLIDE_ROLE_OTHER

    # page_signature 的序号映射也一并钉死：features.py 的 _KNOWN_SLIDE_ROLES 增删
    # role 时 59 键不变但数值含义漂移，这里必须显式换代（tests/** 的 SLF001 已放行）
    assert _SLIDE_ROLE_ORDER == ("closing", "content", "cover", "data", "gallery", "section")
    assert _SLIDE_ROLE_OTHER == 6.0
    # 布局黄金快照：任何 key 增删/重排/大小写变化都会触发失败，
    # 强制开发者 conscious bump feature_schema_version()（Task 7 模型有效性门禁依赖它）
    assert feature_keys() == (
        "deck.profile.academic",
        "deck.profile.balanced",
        "deck.profile.consulting",
        "deck.profile.event",
        "deck.profile.technology",
        "deck.total_slides",
        "finding.confidence",
        "finding.dimension.color",
        "finding.dimension.composition",
        "finding.dimension.consistency",
        "finding.dimension.hierarchy",
        "finding.dimension.media",
        "finding.dimension.typography",
        "finding.evidence_reliability",
        "finding.page_ratio",
        "finding.rule_id.art.color.accent_flood",
        "finding.rule_id.art.color.low_contrast",
        "finding.rule_id.art.color.no_accent",
        "finding.rule_id.art.composition.background_like_area",
        "finding.rule_id.art.composition.corner_cluster",
        "finding.rule_id.art.composition.off_balance",
        "finding.rule_id.art.composition.spacing_drift",
        "finding.rule_id.art.consistency.margin_drift",
        "finding.rule_id.art.consistency.title_drift",
        "finding.rule_id.art.hierarchy.no_focus",
        "finding.rule_id.art.hierarchy.title_too_small",
        "finding.rule_id.art.media.distorted_image",
        "finding.rule_id.art.media.mixed_image_sizes",
        "finding.rule_id.art.media.tiny_image",
        "finding.rule_id.art.typography.flat_scale",
        "finding.rule_id.art.typography.many_families",
        "finding.rule_id.art.typography.tiny_text",
        "finding.severity_ordinal",
        "measure.art.color.accent_flood.accent_ratio",
        "measure.art.color.low_contrast.foreground_match_ratio",
        "measure.art.color.low_contrast.ratio",
        "measure.art.color.no_accent.accent_ratio",
        "measure.art.composition.background_like_area.background_like_ratio",
        "measure.art.composition.corner_cluster.quadrants",
        "measure.art.composition.off_balance.balance_dist",
        "measure.art.composition.off_balance.ink",
        "measure.art.composition.spacing_drift.horizontal",
        "measure.art.composition.spacing_drift.vertical",
        "measure.art.consistency.margin_drift.median",
        "measure.art.consistency.title_drift.size_drift",
        "measure.art.consistency.title_drift.x_drift",
        "measure.art.media.distorted_image.natural_ratio",
        "measure.art.media.distorted_image.physical_ratio",
        "measure.art.media.mixed_image_sizes.spread",
        "measure.art.typography.flat_scale.ratio",
        "measure.art.typography.many_families.families",
        "slide.alignment",
        "slide.density",
        "slide.focus",
        "slide.font_hierarchy",
        "slide.mass",
        "slide.page_signature",
        "slide.palette",
        "slide.spacing",
    )
    assert feature_schema_version() == "1"
