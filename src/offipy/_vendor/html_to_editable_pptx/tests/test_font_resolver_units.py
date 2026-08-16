"""font_resolver 拉黑语义回归测试（纯 Python，monkeypatch 网络层）。

- 网络类临时失败（断网 / 超时 / 5xx）不得把家族持久化标记为 not-in-google-fonts
- 真 404 才写入拉黑索引
"""
import json
import urllib.error

import font_resolver as fr


def _read_idx(path):
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def test_offline_does_not_blacklist(monkeypatch, tmp_path):
    idx_path = tmp_path / "_resolved.json"
    monkeypatch.setattr(fr, "RESOLVED_INDEX", idx_path)

    def offline(*a, **k):
        raise urllib.error.URLError("simulated offline")

    monkeypatch.setattr(fr.urllib.request, "urlopen", offline)
    report = fr.resolve_fonts({"Zz Test Family": {(400, False)}}, set())
    assert "Zz Test Family" in report["unavailable"]
    assert "zz test family" not in _read_idx(idx_path), \
        "一次离线运行不得把字体永久拉黑"


def test_http_404_blacklists(monkeypatch, tmp_path):
    idx_path = tmp_path / "_resolved.json"
    monkeypatch.setattr(fr, "RESOLVED_INDEX", idx_path)

    def gone(req, *a, **k):
        raise urllib.error.HTTPError("http://x", 404, "not found", {}, None)

    monkeypatch.setattr(fr.urllib.request, "urlopen", gone)
    report = fr.resolve_fonts({"Zz Test Family": {(400, False)}}, set())
    assert "Zz Test Family" in report["unavailable"]
    assert _read_idx(idx_path).get("zz test family") == "not-in-google-fonts"


def test_http_5xx_does_not_blacklist(monkeypatch, tmp_path):
    idx_path = tmp_path / "_resolved.json"
    monkeypatch.setattr(fr, "RESOLVED_INDEX", idx_path)

    def boom(req, *a, **k):
        raise urllib.error.HTTPError("http://x", 503, "unavailable", {}, None)

    monkeypatch.setattr(fr.urllib.request, "urlopen", boom)
    report = fr.resolve_fonts({"Zz Test Family": {(400, False)}}, set())
    assert "Zz Test Family" in report["unavailable"]
    assert "zz test family" not in _read_idx(idx_path)


def test_chars_from_measurement_coerces_non_string_text():
    # F2：不可信测量可能注入非字符串 text（DOM 覆盖 / 手写 JSON）——str() 化后不崩，
    # 且数字字符正确进集合
    from embed_fonts import chars_from_measurement
    meas = {
        "slides": [{
            "records": [
                {"runs": [{"text": 42}], "text": 7},
                {"runs": [{"text": "abc"}], "text": None},
            ]
        }]
    }
    chars = chars_from_measurement(meas)
    assert "a" in chars and "b" in chars and "c" in chars
    assert "4" in chars and "2" in chars and "7" in chars


def test_embed_missing_font_skips_and_warns(tmp_path, monkeypatch):
    """#141：字体未缓存 → skip 该 slot（不再 raise FileNotFoundError），并产出降级警告。

    真实 pptx + 非空 FONT_PLAN 指向缺失文件：空 FONT_PLAN 时字体循环根本不进，
    空 zip 会在 rels 重写时 KeyError——必须让字体循环实际走到缺缓存分支。
    """
    import embed_fonts
    from embed_fonts import embed
    from pptx import Presentation

    in_pptx = tmp_path / "in.pptx"
    out_pptx = tmp_path / "out.pptx"
    Presentation().save(str(in_pptx))  # 真实 pptx：presentation.xml/rels/content-types 齐全
    cache = tmp_path / "cache"
    cache.mkdir()
    monkeypatch.setattr(embed_fonts, "CACHE_DIR", cache)  # 空缓存目录 → 字体必缺
    monkeypatch.setattr(embed_fonts, "FONT_PLAN", [
        {"typeface": "Zz Face", "charset": "00", "pitchFamily": "2", "cjk": False,
         "style": "sans", "slots": {"regular": "Zz-Face.ttf"}},
    ])
    warnings: list[dict] = []
    embed(in_pptx, {"slides": [{"records": []}]}, out_pptx, warnings=warnings)
    assert warnings, "缺缓存字体应产生降级警告"
    assert all(w["kind"] == "font" for w in warnings)
    assert out_pptx.exists()  # skip 后整体流程正常产出，不再中途崩


def test_append_measure_warnings_merges_preserving_existing(tmp_path):
    """#141：convert 端把字体降级警告幂等合并进 measurements.json 的 _warnings（保留既有条目）。"""
    from convert import _append_measure_warnings

    anchor = tmp_path / "measurements.json"
    anchor.write_text(json.dumps({"slides": [], "_warnings": [{"kind": "img"}]}),
                      encoding="utf-8")
    new_warnings = [{"kind": "font", "message": "字体未缓存，已跳过子集嵌入"}]
    _append_measure_warnings(anchor, new_warnings)
    data = json.loads(anchor.read_text(encoding="utf-8"))
    assert data["_warnings"] == [
        {"kind": "img"},
        {"kind": "font", "message": "字体未缓存，已跳过子集嵌入"},
    ]


def test_append_measure_warnings_ignores_missing_corrupt_and_nondict(tmp_path):
    """损坏/缺失/非 dict 的 measurements.json 不得阻断转换成功路径。"""
    from convert import _append_measure_warnings

    _append_measure_warnings(tmp_path / "missing.json", [{"kind": "font"}])  # 缺失 → 静默

    corrupt = tmp_path / "corrupt.json"
    corrupt.write_text("{ not json", encoding="utf-8")
    _append_measure_warnings(corrupt, [{"kind": "font"}])  # 损坏 → 静默

    not_dict = tmp_path / "not_dict.json"
    not_dict.write_text("[1, 2]", encoding="utf-8")
    _append_measure_warnings(not_dict, [{"kind": "font"}])  # 非 dict 根 → 静默
