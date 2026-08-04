"""反馈学习测试：记录落盘/加载、权重折算、合法性校验。"""

import json

import pytest

from offipy import feedback
from offipy.feedback import append, dimension_weights, load_records, record_file


def test_append_and_load_roundtrip(tmp_path):
    path = append(
        "palette",
        "fixed",
        severity="HIGH",
        page=2,
        message="色太多",
        source="deck.html",
        feedback_dir=tmp_path,
    )
    assert path == tmp_path / feedback.FEEDBACK_FILE
    records = load_records(tmp_path)
    assert len(records) == 1
    rec = records[0]
    assert rec.dimension == "palette"
    assert rec.action == "fixed"
    assert rec.severity == "HIGH"
    assert rec.page == 2
    assert rec.source == "deck.html"
    assert rec.ts  # 自动打时间戳


def test_append_creates_dir(tmp_path):
    target = tmp_path / "nested" / "dir"
    append("whitespace", "fixed", feedback_dir=target)
    assert record_file(target).exists()


def test_weights_default_when_no_records(tmp_path):
    w = dimension_weights(tmp_path)
    assert w == {d: 1.0 for d in feedback.ALL_DIMENSIONS}


def test_weights_missing_file_ok(tmp_path):
    assert dimension_weights(tmp_path / "nonexistent")[feedback.PALETTE] == 1.0


def test_fixed_boosts_weight(tmp_path):
    append("palette", "fixed", feedback_dir=tmp_path)
    w = dimension_weights(tmp_path)
    assert w["palette"] == pytest.approx(1.5)
    assert w["whitespace"] == 1.0  # 其他维度不受影响


def test_accepted_lowers_weight(tmp_path):
    append("contrast", "accepted", feedback_dir=tmp_path)
    assert dimension_weights(tmp_path)["contrast"] == pytest.approx(0.75)


def test_fixed_caps_at_max(tmp_path):
    for _ in range(10):
        append("palette", "fixed", feedback_dir=tmp_path)
    assert dimension_weights(tmp_path)["palette"] == feedback._WEIGHT_MAX


def test_accepted_floor_at_min(tmp_path):
    for _ in range(10):
        append("type-scale", "accepted", feedback_dir=tmp_path)
    assert dimension_weights(tmp_path)["type-scale"] == feedback._WEIGHT_MIN


def test_fixed_and_accepted_cancel(tmp_path):
    append("whitespace", "fixed", feedback_dir=tmp_path)
    append("whitespace", "fixed", feedback_dir=tmp_path)
    append("whitespace", "accepted", feedback_dir=tmp_path)
    # 2 * 0.5 - 0.25 = 0.75 增量 → 1.75
    assert dimension_weights(tmp_path)["whitespace"] == pytest.approx(1.75)


def test_ignored_does_not_change_weight(tmp_path):
    append("palette", "ignored", feedback_dir=tmp_path)
    assert dimension_weights(tmp_path)["palette"] == 1.0


def test_base_weights_respected(tmp_path):
    append("palette", "fixed", feedback_dir=tmp_path)
    w = dimension_weights(tmp_path, base={"palette": 2.0})
    assert w["palette"] == pytest.approx(2.5)


def test_invalid_dimension_raises(tmp_path):
    with pytest.raises(ValueError):
        append("nope", "fixed", feedback_dir=tmp_path)


def test_invalid_action_raises(tmp_path):
    with pytest.raises(ValueError):
        append("palette", "wat", feedback_dir=tmp_path)


def test_bad_line_skipped(tmp_path):
    f = tmp_path / feedback.FEEDBACK_FILE
    f.parent.mkdir(parents=True, exist_ok=True)
    good = {
        "ts": "2026-01-01T00:00:00+00:00",
        "dimension": "palette",
        "severity": "MID",
        "page": 1,
        "message": "ok",
        "action": "fixed",
    }
    f.write_text("not-json\n" + json.dumps(good, ensure_ascii=False) + "\n", encoding="utf-8")
    records = load_records(tmp_path)
    assert len(records) == 1
    assert records[0].dimension == "palette"


def test_weights_feed_into_audit(tmp_path):
    """端到端：反馈权重可直接传给审计。"""
    from offipy import aesthetic

    append("whitespace", "fixed", feedback_dir=tmp_path)
    weights = dimension_weights(tmp_path)
    # 构造一页留白充足但非零 finding 的页，验证权重真的改变了扣分
    rec = _text({"x": 0, "y": 0, "w": 1900, "h": 1000}, 24, "rgb(20, 20, 20)")
    measurement = {"slides": [{"slide": {"background": "rgb(255, 255, 255)"}, "records": [rec]}]}
    base_report = aesthetic.audit_measurement(measurement)
    weighted_report = aesthetic.audit_measurement(measurement, weights=weights)
    assert weighted_report.pages[0].score < base_report.pages[0].score


def _text(rect, size, color, text="x"):
    return {
        "id": 0,
        "kind": "text",
        "rect": rect,
        "runs": [{"text": text, "fontSize": size, "color": color, "fontFamily": "Arial"}],
    }
