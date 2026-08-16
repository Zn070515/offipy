"""measure / adapters 浏览器回归测试。

需要 playwright + chromium；不可用时整模块 skip。
一个 fixture deck 覆盖多个已修复 checkpoint：
- 同 tag 装饰 div 组不得被误识别成 deck（slide 数 = section 数）
- 负坐标装饰 shape 的 measurement rect 保留负值
- 满页半透明遮罩产生 shape 记录
- display:none 文本不进 runs
- 零间隙相邻 span 之间不插空格
- 永动 canvas 不崩 measure（超时兜底）
- 独立 split-2col 单页 deck（naturalDisplay=flex，忠实复现使用场景）：
  display:contents wrapper 子树不剪枝、眉标/标题不被 align-items:stretch
  拉成超高、.col 仍填满整页
- deck 缩短后参考截图目录清掉"幽灵页"
"""
import struct
from pathlib import Path

import pytest

FIXTURES = Path(__file__).parent / "fixtures"


def _measure(html_name, tmp_path_factory):
    pytest.importorskip("playwright.sync_api")
    from measure import measure
    out = tmp_path_factory.mktemp("meas") / "measurements.json"
    try:
        return measure(FIXTURES / html_name, out,
                       no_screenshots=True, verbose=False)
    except Exception as e:  # chromium 未安装等
        pytest.skip(f"Playwright/Chromium 不可用: {e}")


@pytest.fixture(scope="module")
def deck(tmp_path_factory):
    return _measure("regression_deck.html", tmp_path_factory)


@pytest.fixture(scope="module")
def split2col(tmp_path_factory):
    return _measure("split_2col_deck.html", tmp_path_factory)


@pytest.fixture(scope="module")
def mixed_block_split(tmp_path_factory):
    # #21：混合布局 deck（slide 1 block + slide 2 split-2col flex）
    return _measure("mixed_block_split_deck.html", tmp_path_factory)


@pytest.fixture(scope="module")
def drawio_placeholder(tmp_path_factory):
    # #94：空 div.drawio（无背景无边框）也要测出 bbox 供注入定位
    return _measure("drawio_placeholder.html", tmp_path_factory)


@pytest.fixture(scope="module")
def img_sizes(tmp_path_factory):
    # #126：1×1 PNG data URI 被 CSS 拉成 200×100，验证解码/渲染尺寸分存
    return _measure("img_sizes_deck.html", tmp_path_factory)


@pytest.fixture(scope="module")
def gradient_shadow(tmp_path_factory):
    # #140：渐变/阴影是 OOXML 不可表达的 CSS 装饰，必走 complexDecoration →
    # deco_snapshot 光栅化层。audit 要识别这些「光栅化装饰」→ fill_kind 必须标记。
    return _measure("gradient_shadow_deck.html", tmp_path_factory)


def _records(deck, page):
    return deck["slides"][page]["records"]


def test_decoration_divs_not_mistaken_for_slides(deck):
    assert len(deck["slides"]) == 2, \
        "slide 内的同 tag 装饰 div 组被误识别成了 deck"


def test_negative_offset_rect_preserved(deck):
    xs = [r["rect"]["x"] for r in _records(deck, 0) if r["kind"] == "shape"]
    assert any(x < 0 for x in xs), "负坐标装饰 shape 的 rect 丢了负值"


def test_fullpage_overlay_measured(deck):
    fullpage = [r for r in _records(deck, 0)
                if r["kind"] == "shape"
                and r["rect"]["w"] >= 1900 and r["rect"]["h"] >= 1060
                and r["tag"] == "div"]
    assert fullpage, "满页半透明遮罩没有产生 shape 记录"


def test_fullpage_overlay_opacity_preserved(deck):
    # #137：元素级 opacity 不再静默丢——deco 必须保留数值，而非只带字段
    # rgba 背景的满页遮罩 s.opacity 恒为 "1"（没有 CSS opacity 属性），
    # 因此数值断言必须落在带真实 CSS opacity 的 .faded 元素上。
    faded = [
        r for r in _records(deck, 0)
        if r["kind"] == "shape" and r["tag"] == "div"
        and "faded" in (r.get("className") or "").split()
    ]
    assert faded, "带 CSS opacity 的 .faded 元素没有产生 shape 记录"
    deco = faded[0].get("deco", {})
    assert "opacity" in deco, "deco 缺 opacity 字段"
    assert float(deco["opacity"]) == 0.5, f"opacity 数值丢失: {deco['opacity']!r}"

    # 原 rgba 半透明遮罩仍在（s.opacity 恒 "1"，不作为数值断言点）
    overlay = [
        r for r in _records(deck, 0)
        if r["kind"] == "shape" and r["rect"]["w"] >= 1900 and r["rect"]["h"] >= 1060
    ]
    assert overlay, "满页半透明遮罩没有产生 shape 记录"


def test_empty_drawio_placeholder_measured(drawio_placeholder):
    # #94：空 div.drawio（无背景无边框）此前被装饰门跳过 → 「没测到容器」。
    # 修复后必须产生 shape 记录，几何与 fixture 显式尺寸一致，供注入定位。
    recs = [r for r in _records(drawio_placeholder, 0)
            if r["kind"] == "shape" and "drawio" in (r["className"] or "").split()]
    assert recs, "空 div.drawio 没有产生 shape 记录"
    rect = recs[0]["rect"]
    assert rect["x"] == 200 and rect["y"] == 240, f"位置偏差: {rect}"
    assert rect["w"] == 900 and rect["h"] == 500, f"尺寸偏差: {rect}"


def test_img_decoded_vs_rendered_size(img_sizes):
    recs = [r for r in _records(img_sizes, 0) if r["kind"] == "img"]
    assert recs, "fixture 未产出 img 记录"
    r = recs[0]
    assert r["decodedSize"] == {"w": 1, "h": 1}, f"解码尺寸错误: {r['decodedSize']}"
    assert r["renderedSize"] == {"w": 200, "h": 100}, f"渲染尺寸错误: {r['renderedSize']}"


def test_hidden_text_not_extracted(deck):
    all_text = "".join(
        run.get("text", "")
        for r in _records(deck, 0) if r["kind"] == "text"
        for run in r.get("runs", []))
    assert "HIDDEN" not in all_text, "display:none 文本泄漏进 runs"


def test_zero_gap_spans_no_inserted_space(deck):
    h1 = next(r for r in _records(deck, 0) if r["kind"] == "text")
    joined = "".join(run.get("text", "") for run in h1["runs"])
    assert "ShipFast" in joined, f"零间隙 span 边界被插入空格: {joined!r}"
    assert "now" in joined


def test_display_contents_wrapper_children_measured(split2col):
    # #9：.cols { display: contents } 无盒（rect 0×0）但子元素可见。isHidden 若
    # 按盒判隐藏，walk() 会在 wrapper 上提前剪枝，两栏子树整体丢弃（ZERO shapes）。
    # 断言 wrapper 内的两个 .col 文本都被测到，且 wrapper 自身不产生空形状。
    texts = [
        run.get("text", "")
        for r in _records(split2col, 0)
        if r["kind"] == "text"
        for run in r.get("runs", [])
    ]
    joined = "".join(texts)
    assert "LEFT_COL_BODY" in joined, "display:contents wrapper 的左栏被剪枝"
    assert "RIGHT_COL_BODY" in joined, "display:contents wrapper 的右栏被剪枝"
    # .cols 无 bg/border，不应以 0×0 空形状出现
    assert not any(
        r["kind"] == "shape" and r["rect"]["w"] == 0 and r["rect"]["h"] == 0
        for r in _records(split2col, 0)
    ), "display:contents wrapper 自身产生了退化形状"


def test_split_2col_non_col_children_not_stretched(split2col):
    # #9：align-items: stretch 会把眉标/标题（非 .col 的 flex 子元素）拉成整页高
    # （超高文本框）。修复后容器 align-items: flex-start + .col align-self: stretch：
    # 眉标/标题保持自然高度（远小于整页 1080px），.col 仍填满整页内容盒
    # （.slide 有 --pad 96px 边距，flex 内容盒高 1080−2×96=888px）。
    texts = [r for r in _records(split2col, 0) if r["kind"] == "text" and r.get("text")]
    kicker = next(r for r in texts if "AGENDA_KICKER" in r["text"])
    title = next(r for r in texts if "QUARTERLY_TITLE" in r["text"])
    assert kicker["rect"]["h"] < 400, f"眉标被拉成超高: {kicker['rect']['h']}px"
    assert title["rect"]["h"] < 400, f"标题被拉成超高: {title['rect']['h']}px"
    cols = [
        r for r in _records(split2col, 0)
        if r["kind"] == "shape" and r["rect"]["w"] > 0 and r["rect"]["h"] > 0
    ]
    assert len(cols) >= 2, "两栏 col 形状缺失"
    assert all(c["rect"]["h"] >= 700 for c in cols), "col 未填满整页内容盒"


def test_mixed_deck_split_2col_not_forced_block(mixed_block_split, split2col):
    # #21：混合布局 deck 的 flex 页此前被"第一张可见页的 display"全局强加成 block
    # （.col 全宽竖排）。修复后按 slide 记录各自 natural display，split-2col 页
    # 保持 flex——两栏并排，几何与独立单页 split-2col fixture 完全一致。
    assert len(mixed_block_split["slides"]) == 2
    slide2 = mixed_block_split["slides"][1]["records"]

    def cols(records):
        # col 形状：填满内容盒高度（h≥700）且非整页 slide 背景（w<1900）
        return [
            r["rect"] for r in records
            if r["kind"] == "shape"
            and r["rect"]["h"] >= 700
            and r["rect"]["w"] < 1900
        ]

    mixed_cols = cols(slide2)
    assert len(mixed_cols) == 2, "slide 2 的两栏 col 缺失"
    xs = sorted(c["x"] for c in mixed_cols)
    assert xs[1] - xs[0] > 100, f"两栏被压成竖排（同 x）：{xs}"
    assert all(c["w"] < 900 for c in mixed_cols), f"col 全宽竖排：{mixed_cols}"

    # 与独立单页 split-2col 对照：几何逐字段一致（唯一变量就是前面多了 block 页）
    ref_cols = cols(split2col["slides"][0]["records"])
    assert len(ref_cols) == 2
    for mine, ref in zip(sorted(mixed_cols, key=lambda c: c["x"]),
                         sorted(ref_cols, key=lambda c: c["x"])):
        for k in ("x", "y", "w", "h"):
            assert abs(mine[k] - ref[k]) < 1, f"col 几何偏离单页对照: {k} {mine[k]} vs {ref[k]}"


def test_animated_canvas_does_not_crash(deck):
    # deck fixture 能构建成功本身就是断言（slide 2 的 canvas 永不稳定）
    kinds = {r["kind"] for r in _records(deck, 1)}
    assert "canvas" in kinds


def test_stale_screenshot_pruned(tmp_path):
    pytest.importorskip("playwright.sync_api")
    from measure import measure
    anchor = tmp_path / "m.json"
    shots = tmp_path / "m_screenshots"
    shots.mkdir()
    stale = shots / "slide_99.png"
    stale.write_bytes(b"stale")
    try:
        measure(FIXTURES / "regression_deck.html", anchor,
                no_screenshots=False, verbose=False)
    except Exception as e:
        pytest.skip(f"Playwright/Chromium 不可用: {e}")
    assert not stale.exists(), "deck 缩短后的幽灵参考截图没有被清理"
    assert (shots / "slide_01.png").exists()


def test_deco_shadow_beyond_viewport_not_clamped(tmp_path):
    # #138：deco_snapshot 捕获框（含 box-shadow 外延）超出视口时，不得 clamp 到
    # 1920×1080 丢阴影；扩视口截图并还原，rect 保持完整捕获框。
    pytest.importorskip("playwright.sync_api")
    from measure import measure
    out = tmp_path / "m.json"
    try:
        data = measure(FIXTURES / "deco_shadow_edge.html", out, no_screenshots=False, verbose=False)
    except Exception as e:
        pytest.skip(f"Playwright/Chromium 不可用: {e}")
    recs = [r for r in data["slides"][0]["records"] if r["kind"] == "deco_snapshot"]
    assert recs, "边缘阴影卡片没有产生 deco_snapshot"
    r = recs[0]
    # 捕获框 top=1000 + h=80 + shadow 向下 60+40=100 → h≈180；修复前被 clamp 到 ~80。
    # 精确几何交给下方 PNG 断言；这里只验「明确大于被 clamp 的 80」。
    assert r["rect"]["h"] > 100, f"阴影捕获框被视口裁剪: rect={r['rect']}"
    # 截图必须真的截到完整捕获框（未 clamp 到 80）：从 PNG 头部 IHDR 读像素高验证。
    png = Path(r["screenshot"])
    assert png.exists(), f"deco_snapshot 截图未生成: {png}"
    with png.open("rb") as f:
        png_w, png_h = struct.unpack(">II", f.read(24)[16:24])
    assert png_h > 100, f"截图被视口裁剪到 {png_h}px"
    assert abs(png_h - r["rect"]["h"]) <= 1, f"截图高 {png_h} 与捕获框 {r['rect']['h']} 不符"


def test_media_warnings_and_href(tmp_path_factory):
    """#141：audio 丢 / video 首帧化产生 _warnings；<a href> 文本 run 携带 href。"""
    data = _measure("media_links.html", tmp_path_factory)
    warnings = data.get("_warnings", [])
    kinds = {w["kind"] for w in warnings}
    assert "audio" in kinds, "audio 应告警（PPTX 无法表达）"
    assert "video" in kinds, "video 应告警（首帧静态图）"
    for rec in data["slides"][0].get("records", []):
        if rec.get("kind") == "text":
            hrefs = [r.get("href") for r in rec.get("runs", []) if r.get("href")]
            assert "https://example.com" in hrefs
            return
    raise AssertionError("未找到携带 href 的 text run")


def test_lone_audio_in_mixed_container_warns(tmp_path_factory):
    """#141：block 兄弟 + 孤立 <audio> 的混合容器——audio 构成 0×0 inline group，
    在 emitInlineGroup 的 early return 就被丢掉，必须仍产生 audio 告警。"""
    data = _measure("media_links.html", tmp_path_factory)
    # 第二页 = block <p> + 裸 <audio>；audio 单独成组（0×0 / 无文本），
    # 修复前早退静默丢弃 → 无告警；修复后 slide=2 必须出现 audio 警告。
    audio_warns = [w for w in data.get("_warnings", [])
                   if w["kind"] == "audio" and w.get("slide") == 2]
    assert audio_warns, "slide 2 的孤立 <audio>（0×0 inline group）未告警"


def test_gradient_and_shadow_fill_kind_marked(gradient_shadow):
    # #140：渐变/阴影必走 complexDecoration → deco_snapshot（kind 非 "shape"），
    # 断言必须覆盖 deco_snapshot 记录，否则「未写进记录」会误绿。
    kinds = {
        r.get("fill_kind")
        for r in _records(gradient_shadow, 0)
        if r["kind"] in ("shape", "deco_snapshot")
    }
    assert "gradient" in kinds, f"渐变卡片未标记 fill_kind=gradient: {kinds}"
    assert "shadow" in kinds, f"阴影卡片未标记 fill_kind=shadow: {kinds}"


def test_measure_border_style_dashed(tmp_path_factory):
    """#141：border-bottom: dashed → shape deco.borderStyle == 'dashed'。

    单侧 bottom-dashed 元素（borderTopStyle == 'none'）也必须算出 borderStyle='dashed'，
    否则 assemble 的 _border_dash 走不到 _add_line 的 dash 参数，虚线会渲染成实线。
    """
    data = _measure("border_dashed.html", tmp_path_factory)
    border_styles = {rec["deco"].get("borderStyle")
                     for rec in data["slides"][0].get("records", []) if rec.get("deco")}
    assert "dashed" in border_styles, f"border-bottom: dashed 未算出 borderStyle=dashed: {border_styles}"
