"""asset 声明的 kind='asset' 测量协议（A3 Task 4）。

需要 playwright + chromium；不可用时整模块 skip。用 offipy.assets.declarations
预处理器把 data-asset / data-icon / data-primitive 声明注入确定性 ID 后 measure，
断言：每个声明恰一个 asset record、rect 正确、主题 token（含 dark variant）与
computed color 正确捕获，且 legacy <svg data-icon> 被量成 asset 而非 svg。
"""
import pytest

from offipy.assets.declarations import preprocess_asset_declarations


def _measure_html(html_text, tmp_path_factory):
    pytest.importorskip("playwright.sync_api")
    from measure import measure

    target = tmp_path_factory.mktemp("asset_html") / "assets.html"
    target.write_text(html_text, encoding="utf-8")
    out = tmp_path_factory.mktemp("meas") / "measurements.json"
    try:
        return measure(target, out, no_screenshots=True, verbose=False)
    except Exception as e:  # chromium 未安装等
        pytest.skip(f"Playwright/Chromium 不可用: {e}")


_ASSET_HTML = """<!doctype html>
<html><head><meta charset="utf-8">
<style>
  * { margin: 0; box-sizing: border-box; }
  :root {
    --bg: #ffffff; --surface: #f3f4f6; --ink: #111827; --muted: #6b7280;
    --accent: #2251ff; --accent-soft: #e0e7ff; --divider: #e5e7eb;
  }
  .slide { width: 1920px; height: 1080px; position: relative; overflow: hidden; }
  .dark {
    --bg: #0f172a; --surface: #1e293b; --ink: #f8fafc; --muted: #94a3b8;
    --accent: #38bdf8; --accent-soft: #0c4a6e; --divider: #334155;
  }
  .a { position: absolute; }
</style>
</head><body>
<section class="slide" data-pptx-slide>
  <div class="a" data-primitive="quote-mark" style="left:20px;top:30px;width:120px;height:60px"></div>
  <svg class="a" data-icon="ph:check" viewBox="0 0 256 256"
       style="left:200px;top:40px;width:48px;height:48px;color:#ff0000"></svg>
  <div class="a" data-asset="asset://procedural/pattern/topo?seed=42"
       style="left:320px;top:20px;width:80px;height:80px"></div>
</section>
<section class="slide dark" data-pptx-slide>
  <div class="a" data-asset="asset://procedural/pattern/grid"
       style="left:20px;top:20px;width:60px;height:60px"></div>
</section>
</body></html>
"""


@pytest.fixture(scope="module")
def deck(tmp_path_factory):
    rewritten, decls = preprocess_asset_declarations(_ASSET_HTML)
    assert [d.declaration_id for d in decls] == [
        "asset-s01-001",
        "asset-s01-002",
        "asset-s01-003",
        "asset-s02-001",
    ]
    return _measure_html(rewritten, tmp_path_factory)


def _records(meas, page):
    return meas["slides"][page]["records"]


def _assets(meas):
    return [r for s in meas["slides"] for r in s["records"] if r["kind"] == "asset"]


def _by_id(meas, asset_id):
    return next(r for r in _assets(meas) if r["assetId"] == asset_id)


def test_one_asset_record_per_declaration(deck):
    records = _assets(deck)
    assert [r["assetId"] for r in records] == [
        "asset-s01-001",
        "asset-s01-002",
        "asset-s01-003",
        "asset-s02-001",
    ]


def test_legacy_svg_icon_measured_as_asset_not_svg(deck):
    # 注入副本里 legacy <svg data-icon> 带 data-offipy-asset-id → kind='asset'
    assert not [r for r in deck["slides"][0]["records"] if r["kind"] == "svg"]
    icon = _by_id(deck, "asset-s01-002")
    assert icon["tag"] == "svg"


def test_asset_slides_have_no_other_records(deck):
    # 空容器声明页不产生任何 shape/text 记录（asset 分支 return，不下钻子节点）
    assert [r["kind"] for r in _records(deck, 0)] == ["asset", "asset", "asset"]
    assert [r["kind"] for r in _records(deck, 1)] == ["asset"]


def test_rects_measured_from_dom(deck):
    assert _by_id(deck, "asset-s01-001")["rect"] == pytest.approx(
        {"x": 20, "y": 30, "w": 120, "h": 60}
    )
    assert _by_id(deck, "asset-s01-002")["rect"] == pytest.approx(
        {"x": 200, "y": 40, "w": 48, "h": 48}
    )
    assert _by_id(deck, "asset-s01-003")["rect"] == pytest.approx(
        {"x": 320, "y": 20, "w": 80, "h": 80}
    )
    assert _by_id(deck, "asset-s02-001")["rect"] == pytest.approx(
        {"x": 20, "y": 20, "w": 60, "h": 60}
    )


def test_theme_vars_captured_including_dark_variant(deck):
    light = _by_id(deck, "asset-s01-001")["themeVars"]
    dark = _by_id(deck, "asset-s02-001")["themeVars"]
    for key in ("bg", "surface", "ink", "muted", "accent", "accent-soft", "divider"):
        assert light[key], f"light theme var {key} 未捕获"
    assert light["accent"].lower() == "#2251ff"
    assert light["surface"].lower() == "#f3f4f6"
    assert dark["accent"].lower() == "#38bdf8"  # dark variant 覆盖生效
    assert dark["surface"].lower() == "#1e293b"
    assert light["accent"] != dark["accent"]


def test_computed_color_captured(deck):
    assert _by_id(deck, "asset-s01-002")["color"] == "rgb(255, 0, 0)"
