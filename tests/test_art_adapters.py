import json

import pytest

from art_helpers import make_element, make_scene, make_slide, make_text_element
from offipy.art.adapters import MeasurementAdapter, PptxAuditAdapter, build_scene
from offipy.art.merge import merge_scenes
from offipy.art.models import ArtColor
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
    # font_size_norm = 52/1080
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


def test_measurement_adapter_accepts_json_string_and_path(tmp_path):
    raw = (FIXTURES / "real_measurements.json").read_text(encoding="utf-8")
    scene1 = build_scene(measurements=raw)
    p = tmp_path / "m.json"
    p.write_text(raw, encoding="utf-8")
    scene2 = build_scene(measurements=str(p))
    assert len(scene1.slides) == len(scene2.slides) == 1


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


def test_build_scene_slides_dir_raises():
    with pytest.raises(InvalidArgumentError):
        build_scene(slides_dir="whatever")


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
