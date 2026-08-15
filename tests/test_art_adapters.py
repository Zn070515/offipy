import errno
import json
from pathlib import Path

import pytest

from art_helpers import make_element, make_scene, make_slide, make_text_element
from offipy.art.adapters import (
    MeasurementAdapter,
    PptxAuditAdapter,
    _rgb_string_to_color,
    build_scene,
)
from offipy.art.merge import _merge_element, merge_scenes
from offipy.art.models import ArtColor, _element_from_dict
from offipy.exceptions import InvalidArgumentError

FIXTURES = __import__("pathlib").Path(__file__).parent / "fixtures" / "art"


def test_measurement_adapter_real_fixture():
    raw = (FIXTURES / "real_measurements.json").read_text(encoding="utf-8")
    scene = MeasurementAdapter(json.loads(raw)).build()
    assert scene.width_unit == "px"
    assert len(scene.slides) == 1
    slide = scene.slides[0]
    assert slide.index == 1  # 位置索引 i+1
    assert len(slide.elements) == 3
    # h1 → title；真实 rect 归一化（192/1920=0.1, 54/1080=0.05, 1152/1920=0.6）
    title = slide.elements[0]
    assert title.kind == "text" and title.role == "title"
    assert abs(title.x - 192.0 / 1920.0) < 1e-6
    assert abs(title.width - 1152.0 / 1920.0) < 1e-6
    # style.color → foreground
    assert title.foreground == ArtColor(30, 30, 30)
    # font_size_norm 期望值 52/1080
    assert abs(title.font_size_norm - 52.0 / 1080.0) < 1e-9
    # canvas（有 naturalSize 的 div）→ image，自然尺寸带过来
    canvas = slide.elements[2]
    assert canvas.kind == "image"
    assert canvas.natural_width == 480.0


def test_measurement_adapter_background_and_font_family():
    raw = (FIXTURES / "real_measurements.json").read_text(encoding="utf-8")
    scene = MeasurementAdapter(json.loads(raw)).build()
    slide = scene.slides[0]
    # 元素 deco.hasBg=false → background=None；slide 背景来自 slide.background
    assert slide.background_color == ArtColor(255, 255, 255)
    assert all(e.background is None for e in slide.elements)
    assert slide.elements[0].runs[0].font_family == "Microsoft YaHei"


def test_measurement_adapter_malformed_rect_coords_graceful():
    # 未受信 measurement JSON：rect 数值损坏 → 回退默认 0，不抛 ValueError 崩库
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080, "background": "rgb(255,255,255)"},
                "records": [
                    {
                        "id": 1,
                        "kind": "text",
                        "className": "title",
                        "tag": "h1",
                        "rect": {"x": "abc", "y": None, "w": "600px", "h": 60},
                        "style": {"fontSize": "52px", "color": "rgb(0,0,0)"},
                        "runs": [],
                    }
                ],
            }
        ]
    }
    scene = MeasurementAdapter(raw).build()
    el = scene.slides[0].elements[0]
    assert el.x == 0.0  # "abc" → 0
    assert el.y == 0.0  # None → 0
    assert el.width == 0.0  # "600px" → 0
    assert el.height == pytest.approx(60.0 / 1080.0)


def test_measurement_adapter_malformed_font_size_graceful():
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080},
                "records": [
                    {
                        "id": 1,
                        "kind": "text",
                        "className": "title",
                        "tag": "h1",
                        "rect": {"x": 0, "y": 0, "w": 100, "h": 60},
                        "style": {"fontSize": "abcpx", "color": "rgb(0,0,0)"},
                        "runs": [{"text": "T", "fontSize": "notnum", "color": "rgb(0,0,0)"}],
                    }
                ],
            }
        ]
    }
    scene = MeasurementAdapter(raw).build()
    el = scene.slides[0].elements[0]
    assert el.font_size is None  # "abcpx" → 未知，不崩
    assert el.font_size_norm is None
    assert el.runs[0].font_size is None  # "notnum" → 未知


def test_measurement_adapter_malformed_slide_size_default():
    raw = {"slides": [{"slide": {"width": "wide", "height": None}, "records": []}]}
    scene = MeasurementAdapter(raw).build()
    assert scene.slides[0].width == 1920.0
    assert scene.slides[0].height == 1080.0


def test_measurement_adapter_malformed_color_dict_graceful():
    # deco.bg 是损坏 dict（r 非数字）→ background None；正常 dict color 仍解析
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080},
                "records": [
                    {
                        "id": 1,
                        "kind": "shape",
                        "className": "",
                        "tag": "div",
                        "rect": {"x": 0, "y": 0, "w": 100, "h": 100},
                        "deco": {"hasBg": True, "bg": {"r": "red", "g": 0, "b": 0}},
                        "style": {"color": {"r": 10, "g": 20, "b": 30}},
                        "runs": [],
                    }
                ],
            }
        ]
    }
    scene = MeasurementAdapter(raw).build()
    el = scene.slides[0].elements[0]
    assert el.background is None
    assert el.foreground == ArtColor(10, 20, 30)


def test_measurement_adapter_accepts_json_string_and_path(tmp_path):
    raw = (FIXTURES / "real_measurements.json").read_text(encoding="utf-8")
    scene1 = build_scene(measurements=raw)
    p = tmp_path / "m.json"
    p.write_text(raw, encoding="utf-8")
    scene2 = build_scene(measurements=str(p))
    assert len(scene1.slides) == len(scene2.slides) == 1


def test_measurement_adapter_json_string_survives_oserror_on_path_check(monkeypatch):
    # 回归：Linux 上 Path(超长 JSON 串).is_file() 抛 ENAMETOOLONG（Windows 返回 False），
    # build_scene 必须把它当 JSON 字符串处理而不是抛错。
    raw = (FIXTURES / "real_measurements.json").read_text(encoding="utf-8")

    def _raise_enametoolong(self):
        raise OSError(errno.ENAMETOOLONG, "File name too long")

    monkeypatch.setattr(Path, "is_file", _raise_enametoolong)
    scene = build_scene(measurements=raw)
    assert len(scene.slides) == 1


def test_measurement_adapter_truncates_excessive_elements():
    # per-slide 元素上限：恶意/病态超多元素页截断到上限，下游 O(n²) 分析被约束
    records = [
        {
            "id": i,
            "kind": "shape",
            "className": "",
            "tag": "div",
            "rect": {"x": i, "y": 0, "w": 1, "h": 1},
            "style": {},
            "runs": [],
        }
        for i in range(3001)
    ]
    raw = {"slides": [{"slide": {"width": 1920, "height": 1080}, "records": records}]}
    scene = MeasurementAdapter(raw).build()
    els = scene.slides[0].elements
    assert len(els) == 3000
    assert any(w.code == "art.adapter.elements_truncated" for w in scene.warnings)


def test_pptx_adapter_truncates_excessive_elements():
    from offipy.audit.models import SlideShapeSnapshot

    data = json.loads((FIXTURES / "real_audit_report.json").read_text(encoding="utf-8"))
    report = _report_from_fixture(data)
    for i in range(3001):
        report.shapes.append(
            SlideShapeSnapshot(
                slide_index=1,
                shape_id=1000 + i,
                name=f"s{i}",
                shape_type="Oval",
                role="shape",
                left=0.1,
                top=0.1,
                width=0.1,
                height=0.1,
                z_order=i,
                text="",
                is_rotated=False,
                geometry_unknown=False,
            )
        )
    scene = PptxAuditAdapter(report).build()
    els = scene.slides[0].elements
    assert len(els) == 3000
    assert any(w.code == "art.adapter.elements_truncated" for w in scene.warnings)


def test_pptx_adapter_real_shapes():
    data = json.loads((FIXTURES / "real_audit_report.json").read_text(encoding="utf-8"))
    report = _report_from_fixture(data)
    scene = PptxAuditAdapter(report).build()
    assert scene.width_unit == "pt"
    assert len(scene.slides) == 1
    slide = scene.slides[0]
    assert slide.index == 1  # SlideShapeSnapshot.slide_index 已 1-based，不做 +1
    # element_id 带 pptx- 前缀，与 measurements 的 ID 空间隔离
    assert slide.elements[0].element_id == "pptx-1-2"
    # 英寸 → pt（×72），再按页宽高归一化
    title = slide.elements[0]
    assert abs(title.x - (1.0 * 72.0) / (13.333 * 72.0)) < 1e-6
    # 无字号/颜色证据 → font_size_norm None、foreground None
    assert title.font_size_norm is None
    assert title.foreground is None
    # shape_type Picture → kind image
    assert slide.elements[1].kind == "image"


def test_pptx_adapter_skips_geometry_unknown():
    from offipy.audit.models import SlideShapeSnapshot

    report = _report_from_fixture(
        json.loads((FIXTURES / "real_audit_report.json").read_text(encoding="utf-8"))
    )
    report.shapes.append(
        SlideShapeSnapshot(
            slide_index=1,
            shape_id=99,
            name="unknown",
            shape_type="Oval",
            role="shape",
            left=None,
            top=None,
            width=None,
            height=None,
            z_order=9,
            text="",
            is_rotated=False,
            geometry_unknown=True,
        )
    )
    scene = PptxAuditAdapter(report).build()
    assert len(scene.slides[0].elements) == 2
    assert any(w.code == "art.adapter.geometry_unknown" for w in scene.warnings)


def test_pptx_adapter_keeps_blank_slides():
    data = json.loads((FIXTURES / "real_audit_report.json").read_text(encoding="utf-8"))
    data["slide_count"] = 3
    data["shapes"] = [s for s in data["shapes"] if s["slide_index"] in (1, 3)]
    report = _report_from_fixture(data)
    scene = PptxAuditAdapter(report).build()
    assert len(scene.slides) == 3  # 空白页保留，索引连续
    assert [s.index for s in scene.slides] == [1, 2, 3]
    assert scene.slides[1].elements == []  # slide 2 无 shape → 空页


def test_pptx_enriched_elements_from_real_pptx():
    # #128：PPTX-only 路径从 _ShapeRecord 富集，文字元素携带字号证据。
    # synthetic.pptx 仅 TextBox 12 有显式 run 字号 18pt；其余文字继承主题 → 字号 None。
    from offipy.audit import audit_pptx

    report = audit_pptx(str(FIXTURES.parent / "audit" / "synthetic.pptx"))
    assert report.records, "真实审计报告应携带 _ShapeRecord（富集分支入口）"
    scene = PptxAuditAdapter(report).build()
    els = scene.slides[0].elements
    assert els, "synthetic.pptx 无元素"
    text_el = next((e for e in els if e.has_text() and e.font_size is not None), None)
    assert text_el is not None, "富集后应有文字元素带 font_size"
    assert text_el.font_size == 18.0, "synthetic.pptx TextBox 12 显式字号 18pt"
    assert text_el.font_size_unit == "pt"
    assert text_el.font_size_norm is not None
    assert text_el.runs and text_el.runs[0].font_size == 18.0, "run 应带字号证据"
    # fixture 无任何颜色证据（无 srgbClr / solidFill）：不得虚构 foreground
    assert all(e.foreground is None for e in els), "无颜色证据时 foreground 应为 None"
    # 无显式字体的文字元素：字号保持 None（继承主题，不硬造证据），不崩
    assert any(e.has_text() and e.font_size is None for e in els), "无显式字号的元素保持 None"


def test_build_scene_no_source_raises():
    with pytest.raises(InvalidArgumentError):
        build_scene()


def test_build_scene_pptx_and_pptx_report_conflict(tmp_path):
    with pytest.raises(InvalidArgumentError):
        build_scene(pptx="x.pptx", pptx_report=object())


def test_merge_matches_by_shape_id():
    m_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("t1", "Title", font_size=48.0),
                    make_text_element("b1", "Body", y=0.2, font_size=24.0),
                ],
            )
        ]
    )
    p_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("t1", "Title", font_size=44.0),
                    make_text_element("b1", "Body", y=0.2, font_size=22.0),
                ],
            )
        ],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    assert merged.slides[0].elements[0].source == "merged"
    assert merged.slides[0].elements[0].evidence["pptx"]["font_size"] == 44.0
    assert merged.slides[0].elements[0].evidence["match_confidence"] == 1.0
    assert warnings == []


def test_merge_matches_by_identity_when_ids_differ():
    m_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("m-title", "标题", font_size=48.0, role="title"),
                ],
            )
        ]
    )
    p_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("pptx-title", "标题", font_size=44.0, role="title"),
                ],
            )
        ],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    els = merged.slides[0].elements
    assert len(els) == 1  # 一对一：不重复追加 secondary
    assert els[0].source == "merged"
    assert 0.7 <= els[0].evidence["match_confidence"] <= 0.9
    assert warnings == []


def test_merge_keeps_secondary_only_with_warning():
    m_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("only_measure", "M", font_size=20.0),
                ],
            )
        ]
    )
    p_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("only_pptx", "P", y=0.3, font_size=20.0),
                ],
            )
        ],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    ids = {e.element_id for e in merged.slides[0].elements}
    assert "only_measure" in ids and "only_pptx" in ids
    assert any(w.code == "art.merge.unmatched" for w in warnings)


def test_merge_keeps_secondary_only_slide():
    m_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("a", "A", font_size=20.0),
                ],
            )
        ]
    )
    p_scene = make_scene(
        [
            make_slide(1, elements=[make_text_element("a", "A", font_size=20.0)]),
            make_slide(2, elements=[make_text_element("s2", "S2", font_size=20.0)]),
        ],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    assert len(merged.slides) == 2  # secondary-only slide 保留
    assert any(w.code == "art.merge.slide_secondary_only" for w in warnings)


def test_merge_cross_vocab_text_match():
    """双源 role 词表不一致：measurement body ↔ audit content，文本相同仍匹配。"""
    m_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("m-title", "产品增长报告", role="title", font_size=48.0),
                ],
            )
        ]
    )
    p_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("pptx-1-3", "产品增长报告", role="content", font_size=44.0),
                ],
            )
        ],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    els = merged.slides[0].elements
    assert len(els) == 1  # 一对一：不重复追加 secondary
    assert els[0].source == "merged"
    assert 0.6 <= els[0].evidence["match_confidence"] <= 0.75  # 文本强佐证分支 0.7
    assert warnings == []


def test_merge_normalizes_text_soft_break():
    """软换行 \\x0b（pptx）与 \\n（measurement）归一后视为同文本。"""
    m_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element(
                        "m-t", "产品增长报告\n与下季度规划", role="body", font_size=48.0
                    ),
                ],
            )
        ]
    )
    p_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element(
                        "p-t", "产品增长报告\x0b与下季度规划", role="content", font_size=44.0
                    ),
                ],
            )
        ],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    assert len(merged.slides[0].elements) == 1
    assert merged.slides[0].elements[0].source == "merged"
    assert warnings == []


def test_merge_empty_shape_matches_by_geometry_with_role_soft():
    """空文本形状：role 词表不同（shape vs unknown）也按几何 + role 归一加分匹配。"""
    m_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_element("m-card", kind="shape", role="shape", x=0.1, y=0.3, w=0.4, h=0.3),
                ],
            )
        ]
    )
    p_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_element(
                        "p-card", kind="shape", role="unknown", x=0.1, y=0.3, w=0.4, h=0.3
                    ),
                ],
            )
        ],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    els = merged.slides[0].elements
    assert len(els) == 1  # unknown → shape 归一后 role 加分，几何 d=0 匹配
    assert els[0].source == "merged"
    assert warnings == []


def test_merge_text_identity_not_distance_gated():
    """文本身份分支不受距离门限制：同文本即使几何偏移 >0.2 仍匹配（跨源坐标参考系可能不同）。"""
    m_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element(
                        "m-t", "标题文字", role="title", x=0.1, y=0.1, font_size=48.0
                    ),
                ],
            )
        ]
    )
    p_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element(
                        "p-t", "标题文字", role="title", x=0.6, y=0.6, font_size=44.0
                    ),
                ],
            )
        ],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    els = merged.slides[0].elements
    assert len(els) == 1
    assert els[0].source == "merged"
    assert els[0].evidence["match_confidence"] == 0.8  # 身份分支，不受 d 门限制
    assert warnings == []


def test_merge_text_mismatch_guard_blocks_merge():
    """双方都有文本但不同 → 即使几何接近也不合并（避免错误并证）。"""
    m_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("m-a", "AAA", x=0.1, y=0.1, font_size=20.0),
                ],
            )
        ]
    )
    p_scene = make_scene(
        [
            make_slide(
                1,
                elements=[
                    make_text_element("p-b", "BBB", x=0.1, y=0.1, font_size=20.0),
                ],
            )
        ],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    ids = {e.element_id for e in merged.slides[0].elements}
    assert "m-a" in ids and "p-b" in ids  # 都保留、未合并
    assert any(w.code == "art.merge.unmatched" for w in warnings)


def test_merge_primary_only_slide_warns():
    m_scene = make_scene(
        [
            make_slide(1, elements=[make_text_element("a", "A", font_size=20.0)]),
            make_slide(2, elements=[make_text_element("s2", "S2", font_size=20.0)]),
        ]
    )
    p_scene = make_scene(
        [make_slide(1, elements=[make_text_element("a", "A", font_size=20.0)])],
        width_unit="pt",
    )
    merged, warnings = merge_scenes(primary=m_scene, secondary=p_scene)
    assert any(w.code == "art.merge.slide_missing" for w in warnings)
    # primary-only slide 2 kept, elements stay source "measurement"
    assert merged.slides[1].elements[0].source == "measurement"


def test_pptx_adapter_skips_index_out_of_range():
    from offipy.audit.models import SlideShapeSnapshot

    report = _report_from_fixture(
        json.loads((FIXTURES / "real_audit_report.json").read_text(encoding="utf-8"))
    )
    report.shapes.append(
        SlideShapeSnapshot(
            slide_index=5,
            shape_id=99,
            name="oob",
            shape_type="TextBox",
            role="body",
            left=1.0,
            top=1.0,
            width=1.0,
            height=1.0,
            z_order=9,
            text="",
            is_rotated=False,
            geometry_unknown=False,
        )
    )
    scene = PptxAuditAdapter(report).build()
    assert len(scene.slides[0].elements) == 2  # out-of-range shape skipped
    assert any(w.code == "art.adapter.index_out_of_range" for w in scene.warnings)


def _report_from_fixture(data):
    from offipy.audit.models import AuditConfig, PptxAuditReport, SlideShapeSnapshot

    shapes = [
        SlideShapeSnapshot(
            slide_index=s["slide_index"],
            shape_id=s["shape_id"],
            name=s["name"],
            shape_type=s["shape_type"],
            role=s["role"],
            left=s["left"],
            top=s["top"],
            width=s["width"],
            height=s["height"],
            z_order=s["z_order"],
            text=s["text"],
            is_rotated=s["is_rotated"],
            geometry_unknown=s.get("geometry_unknown", False),
        )
        for s in data["shapes"]
    ]
    return PptxAuditReport(
        schema_version=data["schema_version"],
        offipy_version=data["offipy_version"],
        path=data["path"],
        source_sha256=data["source_sha256"],
        slide_size=tuple(data["slide_size"]),
        slide_count=data["slide_count"],
        config=AuditConfig(),
        shapes=shapes,
        findings=[],
        suppressed=[],
        warnings=[],
    )


def test_make_element_preserves_pixel_evidence():
    from offipy.art.models import ElementPixelEvidence

    pe = ElementPixelEvidence(method="declared_verified", color_confidence=0.8)
    el = make_element("a", pixel_evidence=pe)
    slide = make_slide(1, elements=[el])
    assert slide.elements[0].pixel_evidence is not None
    assert slide.elements[0].pixel_evidence.method == "declared_verified"


def test_build_scene_slides_dir_invalid_dir_raises(tmp_path):
    with pytest.raises(InvalidArgumentError):
        build_scene(slides_dir=str(tmp_path / "nope"))


def test_build_scene_slides_dir_only_creates_empty_scene(tmp_path):
    pytest.importorskip("PIL")
    from PIL import Image

    d = tmp_path / "slides"
    d.mkdir()
    Image.new("RGB", (100, 50), (255, 255, 255)).save(d / "slide_1.png")
    scene = build_scene(slides_dir=str(d))
    assert len(scene.slides) == 1
    assert scene.slides[0].index == 1
    assert scene.slides[0].width == 100.0
    assert scene.width_unit == "px"


def test_build_scene_measurements_plus_slides_dir(tmp_path):
    pytest.importorskip("PIL")
    from PIL import Image

    d = tmp_path / "slides"
    d.mkdir()
    Image.new("RGB", (1920, 1080), (255, 255, 255)).save(d / "slide_1.png")
    m = tmp_path / "m.json"
    m.write_text(
        json.dumps(
            {
                "slides": [
                    {
                        "slide": {"width": 1920, "height": 1080, "background": "rgb(255,255,255)"},
                        "records": [
                            {
                                "id": 1,
                                "kind": "text",
                                "className": "title",
                                "tag": "h1",
                                "rect": {"x": 0, "y": 0, "w": 600, "h": 60},
                                "style": {"fontSize": "52px", "color": "rgb(0,0,0)"},
                                "runs": [{"text": "Title", "fontSize": 52, "color": "rgb(0,0,0)"}],
                            }
                        ],
                    }
                ]
            }
        ),
        encoding="utf-8",
    )
    scene = build_scene(measurements=str(m), slides_dir=str(d))
    assert "measurement" in scene.sources
    assert "pixel" in scene.sources
    assert scene.slides[0].elements[0].pixel_evidence is not None


def test_color_parser_hex_named_percent():
    assert _rgb_string_to_color("#1a2b3c") == ArtColor(26, 43, 60)
    assert _rgb_string_to_color("#1A2B3C") == ArtColor(26, 43, 60)
    assert _rgb_string_to_color("#f00") == ArtColor(255, 0, 0)
    assert _rgb_string_to_color("#ff000080") == ArtColor(255, 0, 0, 128 / 255)
    assert _rgb_string_to_color("white") == ArtColor(255, 255, 255)
    assert _rgb_string_to_color("rgb(100%, 0%, 0%)") == ArtColor(255, 0, 0)
    assert _rgb_string_to_color("rgba(255, 0, 0, 0.5)") == ArtColor(255, 0, 0, 0.5)
    assert _rgb_string_to_color("#f008") == ArtColor(255, 0, 0, 0x88 / 255)  # 4-digit alpha
    assert _rgb_string_to_color("rgba(255, 0, 0, 50%)") == ArtColor(255, 0, 0, 0.5)  # percent alpha
    assert _rgb_string_to_color("rgba(255, 0, 0, 0)") is None  # 全透明 → None
    assert _rgb_string_to_color("rgb(1, 2, 3, 4, 5)") is None  # 分量数错误 → None
    assert _rgb_string_to_color("not-a-color") is None
    assert _rgb_string_to_color("transparent") is None


def test_measurement_adapter_no_warn_on_valid_or_silent_colors():
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080},
                "records": [
                    {
                        "id": 1,
                        "kind": "text",
                        "className": "title",
                        "tag": "h1",
                        "rect": {"x": 0, "y": 0, "w": 600, "h": 60},
                        "style": {"fontSize": "52px", "color": "#1a2b3c"},
                        "runs": [],
                    },
                    {
                        "id": 2,
                        "kind": "shape",
                        "rect": {"x": 0, "y": 0, "w": 100, "h": 100},
                        "deco": {"hasBg": True, "bg": "transparent"},
                    },
                ],
            }
        ]
    }
    scene = MeasurementAdapter(raw).build()
    assert not any(w.code == "art.adapter.color_unparsed" for w in scene.warnings)


def test_measurement_adapter_warns_on_unparseable_color():
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080},
                "records": [
                    {
                        "id": 1,
                        "kind": "text",
                        "className": "title",
                        "tag": "h1",
                        "rect": {"x": 0, "y": 0, "w": 600, "h": 60},
                        "style": {"fontSize": "52px", "color": "obviously-broken("},
                        "runs": [],
                    }
                ],
            }
        ]
    }
    scene = MeasurementAdapter(raw).build()
    assert any(w.code == "art.adapter.color_unparsed" for w in scene.warnings)


def test_kind_map_asset_svg_deco_snapshot():
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080},
                "records": [
                    {"id": 1, "kind": "asset", "rect": {"x": 0, "y": 0, "w": 100, "h": 100}},
                    {"id": 2, "kind": "svg", "rect": {"x": 0, "y": 0, "w": 100, "h": 100}},
                    {
                        "id": 3,
                        "kind": "deco_snapshot",
                        "rect": {"x": 0, "y": 0, "w": 1920, "h": 1080},
                    },
                ],
            }
        ]
    }
    els = MeasurementAdapter(raw).build().slides[0].elements
    assert [e.kind for e in els] == ["image", "image", "shape"]
    assert els[0].role == "image"  # asset → image role
    assert els[1].role == "image"  # svg → image role
    assert els[2].role == "decoration"
    assert els[2].decoration is True
    assert [e.decoration for e in els] == [False, False, True]  # 仅 deco_snapshot 是 decoration


def test_measurement_adapter_decoded_rendered_size():
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080},
                "records": [
                    {
                        "id": 1,
                        "kind": "img",
                        "rect": {"x": 0, "y": 0, "w": 200, "h": 100},
                        "decodedSize": {"w": 1, "h": 1},
                        "renderedSize": {"w": 200, "h": 100},
                    }
                ],
            }
        ]
    }
    el = MeasurementAdapter(raw).build().slides[0].elements[0]
    assert el.decoded_width == 1.0 and el.decoded_height == 1.0
    assert el.natural_width == 200.0 and el.natural_height == 100.0  # rendered 语义保留


def test_measurement_adapter_decoded_missing_falls_back():
    # 旧数据只有 naturalSize（渲染尺寸）→ decoded 退回渲染，drift≈0 但字段不 None
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080},
                "records": [
                    {
                        "id": 1,
                        "kind": "canvas",
                        "rect": {"x": 0, "y": 0, "w": 480, "h": 270},
                        "naturalSize": {"w": 480, "h": 270},
                    }
                ],
            }
        ]
    }
    el = MeasurementAdapter(raw).build().slides[0].elements[0]
    assert el.natural_width == 480.0
    assert el.decoded_width == 480.0  # 退回 naturalSize，不 None


def test_measurement_adapter_opacity():
    # #137：元素级 opacity 从 style / deco 携带到 ArtElement
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080},
                "records": [
                    {
                        "id": 1,
                        "kind": "text",
                        "className": "title",
                        "tag": "h1",
                        "rect": {"x": 0, "y": 0, "w": 600, "h": 60},
                        "style": {"fontSize": "52px", "color": "rgb(0,0,0)", "opacity": "0.5"},
                        "runs": [],
                    },
                    {
                        "id": 2,
                        "kind": "shape",
                        "rect": {"x": 0, "y": 0, "w": 100, "h": 100},
                        "deco": {"hasBg": True, "bg": "rgb(255,0,0)", "opacity": "0.25"},
                    },
                ],
            }
        ]
    }
    els = MeasurementAdapter(raw).build().slides[0].elements
    assert els[0].opacity == 0.5
    assert els[1].opacity == 0.25  # deco 路径兜底


def test_element_field_roundtrip_and_merge_preserves_all_fields():
    # 维护面 4 处同步契约：to_dict/_element_from_dict round-trip 与 _merge_element 都不得丢字段
    el = make_element(
        "rt",
        kind="image",
        role="image",
        x=0.1,
        y=0.2,
        w=0.3,
        h=0.4,
        slide_index=1,
        opacity=0.5,
        decoded_width=800.0,
        decoded_height=600.0,
        fill_kind="gradient",
    )
    rt = _element_from_dict(el.to_dict(), 1, 1080.0)
    assert rt.opacity == 0.5
    assert rt.decoded_width == 800.0 and rt.decoded_height == 600.0
    assert rt.fill_kind == "gradient"

    secondary = make_element(
        "s",
        kind="image",
        role="image",
        x=0.1,
        y=0.2,
        w=0.3,
        h=0.4,
        slide_index=1,
    )
    merged = _merge_element(el, secondary, 1.0)
    assert merged.source == "merged"
    assert merged.opacity == 0.5
    assert merged.decoded_width == 800.0 and merged.decoded_height == 600.0
    assert merged.fill_kind == "gradient"


def test_measurement_adapter_fill_kind():
    raw = {
        "slides": [
            {
                "slide": {"width": 1920, "height": 1080},
                "records": [
                    {
                        "id": 1,
                        "kind": "shape",
                        "rect": {"x": 0, "y": 0, "w": 400, "h": 300},
                        "fill_kind": "gradient",
                    },
                    {"id": 2, "kind": "shape", "rect": {"x": 0, "y": 0, "w": 100, "h": 100}},
                    {
                        "id": 3,
                        "kind": "shape",
                        "rect": {"x": 0, "y": 0, "w": 50, "h": 50},
                        "fill_kind": "",
                    },
                ],
            }
        ]
    }
    els = MeasurementAdapter(raw).build().slides[0].elements
    assert els[0].fill_kind == "gradient"
    assert els[1].fill_kind is None
    assert els[2].fill_kind is None  # 空串 → None（(rec.get("fill_kind") or None) 语义）
