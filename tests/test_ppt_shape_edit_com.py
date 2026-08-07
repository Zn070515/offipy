"""v0.13 S1 形状编辑 op 的真机 COM 冒烟（office-real）。

经公开 RPC（client.call）驱动真实 PowerPoint：建 blank 页 + 文本框 + 嵌套
group，依次验证 read_shapes / 几何 / 文本 / 多 run 字体 / 填充 / 轮廓 / 可见性 /
group 子元素坐标 / z-order / 删除，最后保存重开验证持久化。无 Office 环境自动
跳过（与 test_ppt_export 同模式）。每组结束后 close_pres 清理自身创建的文稿。
"""

import contextlib
import sys
from pathlib import Path

import pytest

from offipy import core
from offipy.client import call

pytestmark = [
    pytest.mark.com,
    pytest.mark.skipif(
        sys.platform != "win32" or not core.running("ppt"),
        reason="需要存活的 PowerPoint（server 8890 持有）",
    ),
]

BLANK = 12  # PP_LAYOUT_BLANK：无占位符，shape 集合干净


@pytest.fixture
def blank_slide():
    """新建演示文稿 + 一张 blank 页，返回 (doc_id, slide_idx)；teardown 关闭文稿。"""
    did = call("ppt", "new_pres")
    idx = call("ppt", "add_slide", layout=BLANK, doc_id=did)
    yield did, idx
    with contextlib.suppress(Exception):
        call("ppt", "close_pres", save=False, doc_id=did)  # 持久化测试已自行关闭，忽略二次关闭


def _pres():
    return core.connect("ppt").ActivePresentation


def _slide(pres, idx):
    slide = pres.Slides(idx)
    slide.Select()
    return slide


def _new_textbox(slide, text, left=72, top=72, width=300, height=60):
    sh = slide.Shapes.AddTextbox(1, left, top, width, height)  # msoTextOrientationHorizontal
    sh.TextFrame.TextRange.Text = text
    return sh


def _select(slide, shapes):
    """把 shape 按传入顺序组合为 group，返回 group shape。

    用 Shapes.Range([名称...]).Group() 而非 Selection 多选——真机探针证实
    PowerPoint 的 Shape.Select(Replace) 经 pywin32 动态分发不追加选择，
    Selection.ShapeRange 永远只有第一个项目。shape 名用 Id 保证唯一。
    """
    names = []
    for i, sh in enumerate(shapes):
        nm = f"GS{i}_{int(sh.Id)}"
        sh.Name = nm
        names.append(nm)
    return slide.Shapes.Range(names).Group()


def _run_fonts(sh):
    """(size, RGB) 列表，覆盖全部 run——验证整范围字体传播（探针 #3）。"""
    tr = sh.TextFrame.TextRange
    n = int(tr.Runs().Count)
    return [
        (float(tr.Runs(i).Font.Size), int(tr.Runs(i).Font.Color.RGB) & 0xFFFFFF)
        for i in range(1, n + 1)
    ]


def test_top_level_edit_ops_smoke(blank_slide):
    did, idx = blank_slide
    pres = _pres()
    slide = _slide(pres, idx)
    sh = _new_textbox(slide, "Hello", left=72, top=72, width=300, height=60)
    sid = int(sh.Id)

    # read_shapes 基础读回
    recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did)
    by_id = {r["shape_id"]: r for r in recs}
    assert sid in by_id
    assert by_id[sid]["text"] == "Hello"
    assert by_id[sid]["coordinate_space"] == "slide"

    # 几何（磅）
    call(
        "ppt",
        "set_shape_geometry",
        slide_idx=idx,
        shape_id=sid,
        left=100,
        top=120,
        width=320,
        height=80,
        doc_id=did,
    )
    recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did)
    by_id = {r["shape_id"]: r for r in recs}
    assert by_id[sid]["left"] == pytest.approx(100)
    assert by_id[sid]["top"] == pytest.approx(120)
    assert by_id[sid]["width"] == pytest.approx(320)
    assert by_id[sid]["height"] == pytest.approx(80)

    # 文本替换 + 多 run 字体整范围传播
    call("ppt", "set_shape_text", slide_idx=idx, shape_id=sid, text="AlphaBeta", doc_id=did)
    tr = sh.TextFrame.TextRange
    tr.Characters(1, 5).Font.Size = 20  # "Alpha"
    tr.Characters(6, 4).Font.Size = 32  # "Beta"
    assert _run_fonts(sh) == [(20, 0), (32, 0)]
    call(
        "ppt",
        "set_shape_font",
        slide_idx=idx,
        shape_id=sid,
        size=28,
        color="#FF0000",
        bold=True,
        doc_id=did,
    )
    # 整范围字体赋值后 PowerPoint 会把相同格式的 run 合并成一个，
    # 所以只断言「所有 run 都落在目标字体」，再确认文本完整无损
    fonts = _run_fonts(sh)
    assert all(f == (28, 0xFF) for f in fonts), fonts
    assert tr.Text == "AlphaBeta"
    assert int(tr.Font.Bold) == -1

    # 填充
    call("ppt", "set_shape_fill", slide_idx=idx, shape_id=sid, color="#00FF00", doc_id=did)
    recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did)
    assert {r["shape_id"]: r for r in recs}[sid]["fill_color"] == "#00FF00"

    # 轮廓（显式 visible=True 保证线可见，read_shapes 才给色）
    call(
        "ppt",
        "set_shape_outline",
        slide_idx=idx,
        shape_id=sid,
        color="#0000FF",
        width=2.5,
        visible=True,
        doc_id=did,
    )
    recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did)
    by_id = {r["shape_id"]: r for r in recs}
    assert by_id[sid]["line_color"] == "#0000FF"
    assert by_id[sid]["line_width"] == pytest.approx(2.5)

    # 可见性
    call("ppt", "set_shape_visible", slide_idx=idx, shape_id=sid, visible=False, doc_id=did)
    recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did)
    assert {r["shape_id"]: r for r in recs}[sid]["visible"] is False
    call("ppt", "set_shape_visible", slide_idx=idx, shape_id=sid, visible=True, doc_id=did)


def test_group_child_edit_smoke(blank_slide):
    did, idx = blank_slide
    pres = _pres()
    slide = _slide(pres, idx)
    a = _new_textbox(slide, "A", left=72, top=72, width=120, height=40)
    b = _new_textbox(slide, "B", left=72, top=140, width=120, height=40)
    a_id, b_id = int(a.Id), int(b.Id)

    g1 = _select(slide, [a, b])
    g1_id = int(g1.Id)
    c = _new_textbox(slide, "C", left=400, top=72, width=120, height=40)
    c_id = int(c.Id)
    g2 = _select(slide, [g1, c])
    g2_id = int(g2.Id)

    # 真机实证：PowerPoint 打开/再次分组时会把嵌套 group 拍平——G2 直接含
    # A/B/C，G1 消失。read_shapes 的嵌套 group_path 递归只在单测 fake 里可达，
    # 这里验证拍平后的真实形态（仍是 group 子元素，走同一编辑路径）。
    recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did)
    by_id = {r["shape_id"]: r for r in recs}
    assert g1_id not in by_id
    assert set(by_id) == {g2_id, a_id, b_id, c_id}
    assert by_id[a_id]["parent_shape_id"] == g2_id
    assert by_id[a_id]["group_path"] == [g2_id]
    assert by_id[c_id]["parent_shape_id"] == g2_id
    assert by_id[c_id]["group_path"] == [g2_id]

    # group 子元素几何（绝对坐标，非旋转 group）
    call("ppt", "set_shape_geometry", slide_idx=idx, shape_id=c_id, left=200, top=300, doc_id=did)
    recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did)
    by_id = {r["shape_id"]: r for r in recs}
    assert by_id[c_id]["left"] == pytest.approx(200)
    assert by_id[c_id]["top"] == pytest.approx(300)

    # group 子元素 z-order（在 G2.GroupItems 内移到最底）
    call("ppt", "set_shape_z_order", slide_idx=idx, shape_id=c_id, z=1, doc_id=did)
    recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did)
    assert {r["shape_id"]: r for r in recs}[c_id]["z_order"] == 1

    # 删除 group 子元素 B，G2 仍含 A 和 C
    call("ppt", "delete_shape", slide_idx=idx, shape_id=b_id, doc_id=did)
    recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did)
    by_id = {r["shape_id"]: r for r in recs}
    assert b_id not in by_id
    assert set(by_id) == {g2_id, a_id, c_id}


def test_shape_edit_persistence_smoke(blank_slide, tmp_path):
    did, idx = blank_slide
    pres = _pres()
    slide = _slide(pres, idx)
    sh = _new_textbox(slide, "Persist", left=72, top=72, width=300, height=60)
    sid = int(sh.Id)
    call("ppt", "set_shape_fill", slide_idx=idx, shape_id=sid, color="#FF8800", doc_id=did)
    call(
        "ppt",
        "set_shape_geometry",
        slide_idx=idx,
        shape_id=sid,
        left=150,
        top=200,
        doc_id=did,
    )

    saved = call("ppt", "save", path=str(tmp_path / "persist.pptx"), doc_id=did)
    assert Path(saved).exists()
    call("ppt", "close_pres", save=False, doc_id=did)

    did2 = call("ppt", "open_pres", path=saved)
    try:
        recs = call("ppt", "read_shapes", slide_idx=idx, doc_id=did2)
        persist = next(r for r in recs if r["text"] == "Persist")
        assert persist["left"] == pytest.approx(150)
        assert persist["top"] == pytest.approx(200)
        assert persist["fill_color"] == "#FF8800"
    finally:
        call("ppt", "close_pres", save=False, doc_id=did2)
