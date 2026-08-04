"""审美审计测试：颜色工具、单页评分、一致性校验、报告输出。"""

import json

from offipy import aesthetic
from offipy.aesthetic import Finding, audit_measurement, load_measurement

# ---------------------------------------------------------------- 颜色工具


def test_parse_rgb_formats():
    assert aesthetic.parse_rgb("rgb(255, 255, 255)") == (255, 255, 255)
    assert aesthetic.parse_rgb("rgba(34, 81, 255, 0.5)") == (34, 81, 255)
    assert aesthetic.parse_rgb("#2251FF") == (34, 81, 255)
    assert aesthetic.parse_rgb("transparent") is None
    assert aesthetic.parse_rgb(None) is None


def test_contrast_ratio_white_black():
    assert aesthetic.contrast_ratio((0, 0, 0), (255, 255, 255)) > 20.0
    assert aesthetic.contrast_ratio((255, 255, 255), (255, 255, 255)) == 1.0


def test_neutral_detection():
    assert aesthetic._is_neutral((0, 0, 0))  # 黑
    assert aesthetic._is_neutral((128, 128, 128))  # 灰
    assert not aesthetic._is_neutral((34, 81, 255))  # 电光蓝
    assert not aesthetic._is_neutral((56, 189, 248))  # 亮青


# ---------------------------------------------------------------- 样本构造


def _text(rect, size, color, text="x"):
    return {
        "id": 0,
        "kind": "text",
        "rect": rect,
        "runs": [{"text": text, "fontSize": size, "color": color, "fontFamily": "Arial"}],
    }


def _shape(rect, bg):
    return {
        "id": 1,
        "kind": "shape",
        "rect": rect,
        "deco": {"bg": bg, "borderTop": False},
    }


def _slide(records, background="rgb(255, 255, 255)"):
    return {
        "slide": {"width": 1920, "height": 1080, "theme": "light", "background": background},
        "records": records,
    }


def _measurement(slides):
    return {"slides": slides}


def _clean_page():
    """留白充足、2 种字号、1 个强调色、高对比的干净页。"""
    return _slide(
        [
            _text({"x": 96, "y": 96, "w": 1200, "h": 90}, 52, "rgb(34, 34, 34)", "标题"),
            _text({"x": 96, "y": 240, "w": 1100, "h": 60}, 24, "rgb(68, 68, 68)", "正文"),
            _text({"x": 96, "y": 360, "w": 200, "h": 80}, 28, "rgb(34, 81, 255)", "+18%"),
            _shape({"x": 96, "y": 520, "w": 400, "h": 200}, "rgb(242, 244, 247)"),
        ]
    )


def _crowded_page():
    """内容过满 + 5 种字号 + 多个非中性色 + 低对比文本。"""
    records = []
    colors = [
        "rgb(255, 0, 0)",
        "rgb(0, 255, 0)",
        "rgb(0, 0, 255)",
        "rgb(255, 255, 0)",
        "rgb(255, 0, 255)",
        "rgb(0, 255, 255)",
    ]
    for i, c in enumerate(colors):
        records.append(_shape({"x": 0, "y": i * 180, "w": 1920, "h": 170}, c))
    for i, size in enumerate([10, 14, 18, 26, 40, 60, 90, 120]):
        records.append(
            _text({"x": 0, "y": i * 130, "w": 1700, "h": 120}, size, "rgb(204, 204, 204)")
        )
    return _slide(records)


# ---------------------------------------------------------------- 单页审计


def test_clean_page_scores_high():
    report = audit_measurement(_measurement([_clean_page()]))
    assert report.pages[0].score >= 85
    assert report.pages[0].findings == []


def test_crowded_page_has_findings():
    report = audit_measurement(_measurement([_crowded_page()]))
    page = report.pages[0]
    dims = {f.dimension for f in page.findings}
    assert "whitespace" in dims
    assert "type-scale" in dims
    assert "palette" in dims
    assert "contrast" in dims
    assert page.score <= 40


def test_whitespace_thresholds():
    # 大面积内容 → 命中留白检查
    big = _slide([_text({"x": 0, "y": 0, "w": 1920, "h": 1080}, 24, "rgb(20, 20, 20)")])
    findings, ratio = aesthetic._audit_whitespace(big["records"], 1)
    assert ratio > 0.75
    assert findings and findings[0].severity == "HIGH"


def test_background_shape_not_counted_as_content():
    # convert 给每页记一个整页背景 shape，面积≠内容，不应把留白判成 0%
    records = [
        {
            "id": 0,
            "kind": "shape",
            "rect": {"x": 0, "y": 0, "w": 1920, "h": 1080},
            "deco": {"bg": "rgb(255, 255, 255)"},
        },
        _text({"x": 96, "y": 96, "w": 1200, "h": 90}, 52, "rgb(34, 34, 34)", "标题"),
    ]
    findings, ratio = aesthetic._audit_whitespace(records, 1)
    assert findings == []
    assert ratio < 0.75


def test_type_scale_clusters():
    assert aesthetic._cluster_font_sizes([52, 24, 24, 18]) == 3
    assert aesthetic._cluster_font_sizes([52, 50, 51, 24]) == 2  # 52/50/51 合并


# ---------------------------------------------------------------- 一致性


def _multi_page_measurement():
    p1 = _clean_page()
    p2 = _clean_page()
    p3 = _clean_page()
    # 第 2 页换大字号标题 + 深色背景，第 3 页换第三套背景 → 触发漂移
    p2["records"][0]["runs"][0]["fontSize"] = 120
    p2["slide"]["background"] = "rgb(5, 28, 44)"
    p3["slide"]["background"] = "rgb(220, 220, 235)"
    return _measurement([p1, p2, p3])


def test_consistency_detects_drift():
    report = audit_measurement(_multi_page_measurement())
    dims = {f.dimension for f in report.deck_findings}
    assert "consistency" in dims


def test_consistency_ok_for_same_style_pages():
    report = audit_measurement(_measurement([_clean_page(), _clean_page()]))
    assert report.deck_findings == []


# ---------------------------------------------------------------- 入口与输出


def test_audit_auto_finds_measurement(tmp_path):
    html = tmp_path / "deck.html"
    html.write_text("<html></html>", encoding="utf-8")
    audit_dir = tmp_path / "deck_audit" / "_cache"
    audit_dir.mkdir(parents=True)
    (audit_dir / "measurements.json").write_text(
        json.dumps(_measurement([_clean_page()])), encoding="utf-8"
    )
    report = aesthetic.audit(str(html))
    assert report.pages and report.pages[0].score >= 85


def test_load_measurement_missing_raises(tmp_path):
    import pytest

    with pytest.raises(FileNotFoundError):
        load_measurement(tmp_path / "nope.json")


def test_report_markdown(tmp_path):
    report = audit_measurement(_measurement([_clean_page(), _crowded_page()]))
    md = report.markdown()
    assert "审美审计报告" in md
    assert "总分" in md
    assert "第 1 页" in md and "第 2 页" in md


def test_finding_to_dict_roundtrip():
    f = Finding("palette", "MID", 2, "消息", {"n": 3})
    d = f.to_dict()
    assert d["dimension"] == "palette"
    assert d["detail"] == {"n": 3}
