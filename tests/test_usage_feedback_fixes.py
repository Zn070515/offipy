"""真实使用项目 Issue 反馈的 8 项修复测试（无 COM 依赖，纯逻辑/桩替身）。

覆盖：deck audit 目录保留（#11）、进程清理 helper（#13）、Word 合并单元格
列宽回退（#15）、Word 读文本归一化（#16）、PPT 内嵌图片（#17）、Word 插图
文末 Range（#18）、server 重建前回收僵尸进程（#19）、charts className 匹配（#20）。
"""

import json
from pathlib import Path

import pytest

from offipy import core, server
from offipy._comguard import _COM_ERROR
from offipy.charts import load_chart_boxes
from offipy.deck import _preserve_audit_dir

# ---------------------------------------------------------------- #11 deck


def test_preserve_audit_dir_renames_tmp_to_final(tmp_path):
    tmp = tmp_path / "randxyz.pptx"
    tmp_audit = tmp_path / "randxyz_audit"
    (tmp_audit / "_cache").mkdir(parents=True)
    (tmp_audit / "_cache" / "measurements.json").write_text("{}", encoding="utf-8")
    final = tmp_path / "deck.pptx"
    _preserve_audit_dir(str(tmp), str(final))
    assert (tmp_path / "deck_audit" / "_cache" / "measurements.json").exists()
    assert not tmp_audit.exists()


def test_preserve_audit_dir_overwrites_existing_final(tmp_path):
    # final 名已有 audit 目录 → 先删旧再 rename（不留两套、不残留旧文件）
    (tmp_path / "deck_audit").mkdir()
    (tmp_path / "deck_audit" / "old.txt").write_text("old", encoding="utf-8")
    (tmp_path / "randxyz_audit").mkdir()
    (tmp_path / "randxyz_audit" / "measurements.json").write_text("{}", encoding="utf-8")
    _preserve_audit_dir(str(tmp_path / "randxyz.pptx"), str(tmp_path / "deck.pptx"))
    assert (tmp_path / "deck_audit" / "measurements.json").exists()
    assert not (tmp_path / "deck_audit" / "old.txt").exists()


def test_preserve_audit_dir_noop_without_tmp_audit(tmp_path):
    _preserve_audit_dir(str(tmp_path / "x.pptx"), str(tmp_path / "deck.pptx"))
    assert not (tmp_path / "deck_audit").exists()


# ---------------------------------------------------------------- #13 core


def test_reap_process_none_noop():
    core.reap_process(None)  # 不炸、不做任何事


def test_wait_process_exit_none_true():
    assert core.wait_process_exit(None) is True


def test_app_process_pid_non_com_object_none():
    # 拿不到 HWND（无活动窗口/非 COM 对象）→ None，调用方跳过清理
    assert core.app_process_pid(object(), "excel") is None


# ---------------------------------------------------------------- #15 word


def test_set_table_col_width_falls_back_to_cells_on_merged(monkeypatch):
    # 合并单元格后 Columns(col) 本身就被拒访（「表格有混合的单元格宽度」）→
    # 彻底绕开列对象，逐行 table.Cell(r, col).Width 设宽；每个 cell 有独立
    # Width 可写，先合并后调列宽也能成立。
    if not _COM_ERROR:
        pytest.skip("pywin32 未装，无法构造 com_error 触发回退")
    from offipy import word

    err = _COM_ERROR((-2147352567, "表格有混合的单元格宽度", None, 0))

    class _Cell:
        def __init__(self):
            self.Width = None

    class _MergedCol:
        # Columns(col) 在混合列宽表格上任何属性访问都被拒访（含 .Width 与
        # .Cells）——设宽即抛错，fallback 才能走到逐行 table.Cell(r, col) 分支。
        def __setattr__(self, name, value):
            if name == "Width":
                raise err
            object.__setattr__(self, name, value)

    class _Table:
        def __init__(self):
            self._col = _MergedCol()
            self.Rows = type("R", (), {"Count": 3})()
            self.cells = [[_Cell() for _ in range(3)] for _ in range(3)]

        def Columns(self, col):
            assert col == 2
            return self._col

        def Cell(self, row, col):
            assert col == 2
            return self.cells[row - 1][col - 1]

    class _Doc:
        def __init__(self):
            self.table = _Table()

        def Tables(self, idx):
            return self.table

    doc = _Doc()
    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: doc)
    app.set_table_col_width(1, 2, 300.0, doc_id="d1")
    widths = [doc.table.cells[r][1].Width for r in range(3)]
    assert widths == [300.0, 300.0, 300.0]


def test_set_table_col_width_plain_path_no_cells_touched(monkeypatch):
    from offipy import word

    class _Col:
        def __init__(self):
            self.Width = None

    class _Table:
        def __init__(self):
            self._col = _Col()

        def Columns(self, col):
            return self._col

    class _Doc:
        def Tables(self, idx):
            return self._table

    table = _Table()
    doc = _Doc()
    doc._table = table
    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: doc)
    app.set_table_col_width(1, 1, 400.0, doc_id="d1")
    assert table._col.Width == 400.0


# ---------------------------------------------------------------- #16 word


def test_read_doc_text_normalizes_cell_and_paragraph_marks(monkeypatch):
    from offipy import word

    class _Doc:
        @property
        def Content(self):
            return type("C", (), {"Text": "a\x07b\r\nc\rd"})()

    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: _Doc())
    out = app.read_doc_text(doc_id="d1")
    assert out == "a | b\nc\nd"  # \r\n → \n（不双换行）、\x07 → " | "


# ---------------------------------------------------------------- #17 ppt


def test_add_picture_embeds_with_save_with_document(tmp_path, monkeypatch):
    import os

    from offipy import ppt

    img = tmp_path / "photo.png"
    img.write_bytes(b"fake-png")  # 文件必须真实存在（#32：入口前置校验源文件）
    captured = {}

    class _Shapes:
        def AddPicture(self, path, link_to_file, save_with_document, left, top, width, height):
            captured.update(
                path=path,
                link=link_to_file,
                save=save_with_document,
                left=left,
                top=top,
                width=width,
                height=height,
            )
            return object()

    class _Slide:
        Shapes = _Shapes()

    class _Slides:
        Count = 1

        def __call__(self, idx):
            assert idx == 1
            return _Slide()

    class _Pres:
        Slides = _Slides()

    app = ppt.PptApp.__new__(ppt.PptApp)
    monkeypatch.setattr(app, "_require_pres", lambda doc_id: _Pres())
    app.add_picture(1, str(img), 100, 200, 300, 400, doc_id="p1")
    assert captured["save"] == -1  # msoTrue：内嵌（LinkToFile=False 时传 0 会被拒）
    assert captured["link"] == 0
    assert captured["path"] == os.path.normpath(Path(str(img)).resolve())
    assert captured["left"] == 100 and captured["top"] == 200


# ---------------------------------------------------------------- #18 word


def test_insert_image_passes_normalized_path_and_end_range(tmp_path, monkeypatch):
    import os

    from offipy import word

    img = tmp_path / "pic.jpg"
    img.write_bytes(b"fake-jpeg-content")  # 文件必须真实存在（#30：入口前置校验源文件）
    captured = {}

    class _Rng:
        def Collapse(self, wd):
            captured["collapse"] = wd

    class _InlineShapes:
        Count = 1

        def AddPicture(self, path, Range=None):
            captured["path"] = path
            captured["has_range"] = Range is not None
            return type("S", (), {"Width": None, "Height": None})()

    class _Doc:
        def __init__(self):
            self.Content = _Rng()
            self.InlineShapes = _InlineShapes()

    doc = _Doc()
    app = word.WordApp.__new__(word.WordApp)
    monkeypatch.setattr(app, "_require_doc", lambda doc_id: doc)
    app.insert_image(str(img), width=200, height=150, doc_id="d1")
    assert captured["path"] == os.path.normpath(Path(str(img)).resolve())
    assert captured["has_range"] is True  # 插图落在文末 Range，不覆盖既有内容
    assert captured["collapse"] == 0  # wdCollapseEnd


# ---------------------------------------------------------------- #19 server


def test_rebuild_reaps_own_process_before_reconnect(monkeypatch):
    reaped = []

    class FakeApp:
        def reap_own_process(self):
            reaped.append("reap")

    old = dict(server._APPS)
    server._APPS.clear()
    fake = FakeApp()
    server._APPS["excel"] = fake
    monkeypatch.setattr(server, "get_app", lambda name: "new-app")
    try:
        new = server._rebuild(fake)
        assert reaped == ["reap"]
        assert "excel" not in server._APPS  # 先弹掉旧实例再重建
        assert new == "new-app"
    finally:
        server._APPS.clear()
        server._APPS.update(old)


def test_rebuild_unknown_app_returns_same():
    app = object()
    assert server._rebuild(app) is app


def test_rebuild_app_without_reap_still_reconnects(monkeypatch):
    # 无 reap_own_process 的 App（老实现）不炸，只重建
    old = dict(server._APPS)
    server._APPS.clear()
    fake = object()
    server._APPS["ppt"] = fake
    monkeypatch.setattr(server, "get_app", lambda name: "new-app")
    try:
        assert server._rebuild(fake) == "new-app"
    finally:
        server._APPS.clear()
        server._APPS.update(old)


# ---------------------------------------------------------------- #20 charts


def test_load_chart_boxes_matches_chart_class(tmp_path):
    path = tmp_path / "measurements.json"
    path.write_text(
        json.dumps(
            {
                "slides": [
                    {
                        "records": [
                            {
                                "className": "chart",
                                "rect": {"x": 96, "y": 188, "w": 1728, "h": 460},
                            },
                        ]
                    }
                ]
            }
        ),
        encoding="utf-8",
    )
    assert load_chart_boxes(str(path)) == {1: {"x": 96, "y": 188, "w": 1728, "h": 460}}


def test_load_chart_boxes_ignores_chart_note(tmp_path):
    # chart-note 分词后是 ["chart-note"]，不含 "chart" token → 不匹配
    path = tmp_path / "measurements.json"
    path.write_text(
        json.dumps(
            {
                "slides": [
                    {
                        "records": [
                            {"className": "chart-note", "rect": {"x": 0, "y": 0, "w": 1, "h": 1}},
                            {
                                "className": "chart",
                                "rect": {"x": 96, "y": 188, "w": 1728, "h": 460},
                            },
                        ]
                    }
                ]
            }
        ),
        encoding="utf-8",
    )
    assert load_chart_boxes(str(path)) == {1: {"x": 96, "y": 188, "w": 1728, "h": 460}}


def test_measure_emits_class_name_in_shape_record():
    # #20 根因：measure.py 的 shape 记录必须带 className，charts.py 才匹配得到图表容器
    from offipy import __file__ as _offipy_init

    src = Path(_offipy_init).parent / "_vendor" / "html_to_editable_pptx" / "scripts" / "measure.py"
    text = src.read_text(encoding="utf-8")
    assert "className" in text and "el.className" in text
