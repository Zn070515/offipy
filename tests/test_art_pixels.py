"""slides_dir 像素增强测试。Pillow 缺失时整体跳过（pixels.py 是惰性依赖）。"""

import json

import pytest

pytest.importorskip("PIL")

from PIL import Image

from art_helpers import make_element
from offipy.art.models import ArtColor, ArtScene, ArtSlide
from offipy.art.pixels import PixelEnricher, empty_scene_from_slides
from offipy.exceptions import InvalidArgumentError


def _img(size=(100, 100), color=(255, 255, 255)):
    return Image.new("RGB", size, color)


def _write(d, name, im):
    p = d / name
    im.save(p)
    return p


def _png_bytes(w, h):
    """最小 PNG（IHDR+空 IDAT）：头部声明大尺寸但不含像素数据（零分配）。"""
    import struct
    import zlib

    def chunk(typ, data):
        c = typ + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
    idat = chunk(b"IDAT", zlib.compress(b""))
    return sig + ihdr + idat


def test_scan_maps_pages_and_conflicts(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", _img())
    _write(d, "slide_01.png", _img())
    _write(d, "slide_2.png", _img())
    en = PixelEnricher(d)
    pages, warnings = en.scan()
    assert 1 not in pages  # 同整数页冲突 → 该页排除
    assert pages[2].name == "slide_2.png"
    assert any(w.code == "art.pixel.page_conflict" for w in warnings)


def test_empty_scene_from_slides_dimensions(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", _img(size=(100, 50)))
    _write(d, "slide_3.png", _img(size=(200, 100)))
    scene = empty_scene_from_slides(d)
    assert [s.index for s in scene.slides] == [1, 3]  # 保留真实页号，不重排
    assert scene.slides[0].width == 100.0 and scene.slides[0].height == 50.0
    assert scene.width_unit == "px"


def test_empty_scene_from_slides_corrupt_png_warns(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    (d / "slide_1.png").write_bytes(b"not-a-png")
    scene = empty_scene_from_slides(d)
    assert [s.index for s in scene.slides] == []  # 损坏页被跳过，不硬崩
    assert any(w.code == "art.pixel.decode_failed" for w in scene.warnings)


def test_empty_scene_from_slides_page_gap_warns(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", _img())
    _write(d, "slide_3.png", _img())
    scene = empty_scene_from_slides(d)
    assert [s.index for s in scene.slides] == [1, 3]  # 保留真实页号
    assert any(w.code == "art.pixel.page_gap" for w in scene.warnings)


def test_empty_scene_from_slides_consecutive_no_gap(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", _img())
    _write(d, "slide_2.png", _img())
    scene = empty_scene_from_slides(d)
    assert [s.index for s in scene.slides] == [1, 2]
    assert not any(w.code == "art.pixel.page_gap" for w in scene.warnings)


def test_enrich_declared_verified(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    im = _img()
    # 元素区域 (0.2,0.2,0.4,0.3) → (20,20,40,30)，黑色块占多数
    for y in range(22, 28):
        for x in range(22, 38):
            im.putpixel((x, y), (0, 0, 0))
    _write(d, "slide_1.png", im)
    el = make_element(
        "t",
        kind="text",
        role="body",
        x=0.2,
        y=0.2,
        w=0.2,
        h=0.1,
        foreground=ArtColor(0, 0, 0),
    )
    scene = ArtScene(
        slides=[ArtSlide(index=1, width=100.0, height=100.0, elements=[el])],
        width_unit="px",
    )
    en = PixelEnricher(d)
    en.enrich(scene)
    assert scene.slides[0].pixel_evidence is not None
    pe = scene.slides[0].elements[0].pixel_evidence
    assert pe.method == "declared_verified"
    assert pe.foreground_match_ratio >= 0.3
    assert "pixel" in scene.sources
    assert scene.metadata["pixel_pages_covered"] == 1


def test_enrich_text_declared_not_found(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", _img())  # 全白页
    el = make_element(
        "t",
        kind="text",
        role="body",
        x=0.2,
        y=0.2,
        w=0.2,
        h=0.1,
        foreground=ArtColor(0, 0, 0),
    )
    scene = ArtScene(
        slides=[ArtSlide(index=1, width=100.0, height=100.0, elements=[el])],
        width_unit="px",
    )
    PixelEnricher(d).enrich(scene)
    pe = scene.slides[0].elements[0].pixel_evidence
    assert pe.method == "declared_not_found"


def test_enrich_shape_center_fill_verified(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    im = _img()
    # shape 区域 (0.2,0.2,0.6,0.6)，内缩 15% 内填蓝
    for y in range(28, 52):
        for x in range(28, 52):
            im.putpixel((x, y), (0, 0, 255))
    _write(d, "slide_1.png", im)
    el = make_element(
        "s",
        kind="shape",
        role="shape",
        x=0.2,
        y=0.2,
        w=0.4,
        h=0.4,
        background=ArtColor(0, 0, 255),
    )
    scene = ArtScene(
        slides=[ArtSlide(index=1, width=100.0, height=100.0, elements=[el])],
        width_unit="px",
    )
    PixelEnricher(d).enrich(scene)
    pe = scene.slides[0].elements[0].pixel_evidence
    assert pe.method == "center_fill_verified"


def test_enrich_missing_deck_info_warns(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", _img())
    scene = ArtScene(slides=[ArtSlide(index=1, width=100.0, height=100.0)])
    PixelEnricher(d).enrich(scene)
    assert any(w.code == "art.pixel.source_unverified" for w in scene.warnings)
    assert "pixel" in scene.sources  # 分析继续，只标记未验证


def test_enrich_fingerprint_mismatch_raises(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", _img())
    (d / "_deck_info.json").write_text(
        json.dumps({"schema": 1, "pptx_sha256": "abc123"}), encoding="utf-8"
    )
    scene = ArtScene(slides=[ArtSlide(index=1, width=100.0, height=100.0)])
    en = PixelEnricher(d)
    with pytest.raises(InvalidArgumentError):
        en.enrich(scene, expected_sha256="def456")


def test_enrich_no_pixel_source_when_no_valid_page(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", _img(size=(100, 100)))  # 1:1，场景 16:9 → 纵横比不符
    scene = ArtScene(slides=[ArtSlide(index=1, width=1920.0, height=1080.0)])
    PixelEnricher(d).enrich(scene)
    assert "pixel" not in scene.sources
    assert scene.metadata["pixel_pages_covered"] == 0
    assert any(w.code == "art.pixel.aspect_mismatch" for w in scene.warnings)


def test_enrich_page_extra_and_missing(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", _img())
    _write(d, "slide_9.png", _img())
    scene = ArtScene(slides=[ArtSlide(index=1, width=100.0, height=100.0)])
    PixelEnricher(d).enrich(scene)
    assert any(w.code == "art.pixel.page_extra" for w in scene.warnings)
    # 场景有页 2，slides_dir 无 → 页面无对应 PNG
    scene2 = ArtScene(
        slides=[
            ArtSlide(index=1, width=100.0, height=100.0),
            ArtSlide(index=2, width=100.0, height=100.0),
        ]
    )
    PixelEnricher(d).enrich(scene2)
    assert any(w.code == "art.pixel.page_missing" for w in scene2.warnings)


def test_enrich_rgba_transparent_excluded(tmp_path):
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", Image.new("RGBA", (100, 100), (0, 0, 0, 0)))  # 全透明
    el = make_element(
        "t",
        kind="text",
        role="body",
        x=0.2,
        y=0.2,
        w=0.2,
        h=0.1,
        foreground=ArtColor(0, 0, 0),
    )
    scene = ArtScene(
        slides=[ArtSlide(index=1, width=100.0, height=100.0, elements=[el])],
        width_unit="px",
    )
    PixelEnricher(d).enrich(scene)
    pe = scene.slides[0].elements[0].pixel_evidence
    assert pe.background is None  # 透明像素不进背景估计
    assert pe.method != "declared_verified"  # 透明像素不当成黑/白验证成功


def test_enrich_decompression_bomb_hard_reject(tmp_path):
    # 解压炸弹：头部声明 72M 像素（> 60M 上限）。必须硬拒绝 InvalidArgumentError，
    # 不允许降级为 decode_failed warning 继续分析——输入是敌意的，不是普通坏文件。
    d = tmp_path / "slides"
    d.mkdir()
    (d / "slide_1.png").write_bytes(_png_bytes(12000, 6000))
    scene = ArtScene(slides=[ArtSlide(index=1, width=100.0, height=100.0)])
    with pytest.raises(InvalidArgumentError):
        PixelEnricher(d).enrich(scene)


def test_enrich_large_page_downsampled_background(tmp_path):
    # >256px 大页：背景估计在降采样图上进行（不整图直算），结果仍正确、不崩。
    d = tmp_path / "slides"
    d.mkdir()
    _write(d, "slide_1.png", Image.new("RGB", (1024, 576), (240, 240, 240)))
    el = make_element(
        "t",
        kind="text",
        role="body",
        x=0.2,
        y=0.2,
        w=0.2,
        h=0.1,
        foreground=ArtColor(0, 0, 0),
    )
    scene = ArtScene(
        slides=[ArtSlide(index=1, width=1024.0, height=576.0, elements=[el])],
        width_unit="px",
    )
    PixelEnricher(d).enrich(scene)
    pe = scene.slides[0].pixel_evidence
    assert pe.background is not None
    assert (pe.background.r, pe.background.g, pe.background.b) == (240, 240, 240)
