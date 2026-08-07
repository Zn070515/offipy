"""ShapeInfo / SHAPE_TYPE_NAMES 冻结契约（S1 Task 1）。

参考表硬编码自微软官方 MsoShapeType 枚举（VBA-Docs api/Office.MsoShapeType.md，
2026-08-07 gh 拉取核实），独立于项目常量，防「错误实现+错误测试彼此一致」。
1-20 为 round-10 真机实证；21-31 按官方文档；本次会话 probe 7 真机复核
AutoShape/Callout/Group/Line/TextBox/Table 的 .Type 值。
"""

import subprocess
import sys

from offipy import SHAPE_TYPE_NAMES, ShapeInfo, models, shape_type_name

# 微软官方 MsoShapeType 完整映射（VBA-Docs 原表，字段名用描述式 snake_case）
MS_MSO_SHAPE_TYPE = {
    -2: "mixed",  # msoShapeTypeMixed（ShapeRange.Type 返回-only）
    1: "auto_shape",  # msoAutoShape
    2: "callout",  # msoCallout
    3: "chart",  # msoChart
    4: "comment",  # msoComment
    5: "freeform",  # msoFreeform
    6: "group",  # msoGroup
    7: "embedded_ole_object",  # msoEmbeddedOLEObject
    8: "form_control",  # msoFormControl
    9: "line",  # msoLine
    10: "linked_ole_object",  # msoLinkedOLEObject
    11: "linked_picture",  # msoLinkedPicture
    12: "ole_control_object",  # msoOLEControlObject
    13: "picture",  # msoPicture
    14: "placeholder",  # msoPlaceholder
    15: "text_effect",  # msoTextEffect
    16: "media",  # msoMedia
    17: "text_box",  # msoTextBox
    18: "script_anchor",  # msoScriptAnchor
    19: "table",  # msoTable
    20: "canvas",  # msoCanvas
    21: "diagram",  # msoDiagram
    22: "ink",  # msoInk
    23: "ink_comment",  # msoInkComment
    24: "smart_art",  # msoIgxGraphic（SmartArt graphic）
    25: "slicer",  # msoSlicer
    26: "web_video",  # msoWebVideo
    27: "content_app",  # msoContentApp
    28: "graphic",  # msoGraphic
    29: "linked_graphic",  # msoLinkedGraphic
    30: "3d_model",  # mso3DModel
    31: "linked_3d_model",  # msoLinked3DModel
}

EXPECTED_SHAPEINFO_FIELDS = [
    "shape_id",
    "name",
    "shape_type",
    "shape_type_name",
    "left",
    "top",
    "width",
    "height",
    "coordinate_space",
    "coordinate_unit",
    "rotation",
    "visible",
    "fill_color",
    "fill_transparency",
    "line_color",
    "line_width",
    "has_text_frame",
    "text",
    "font_size",
    "font_name",
    "font_color",
    "is_placeholder",
    "placeholder_type",
    "placeholder_type_name",
    "parent_shape_id",
    "group_path",
    "z_order",
]


def test_shape_type_names_matches_official_enum():
    assert SHAPE_TYPE_NAMES == MS_MSO_SHAPE_TYPE


def test_shape_type_name_known_values():
    assert shape_type_name(1) == "auto_shape"
    assert shape_type_name(6) == "group"
    assert shape_type_name(9) == "line"
    assert shape_type_name(13) == "picture"
    assert shape_type_name(14) == "placeholder"
    assert shape_type_name(17) == "text_box"
    assert shape_type_name(19) == "table"


def test_shape_type_name_unknown_fallback():
    assert shape_type_name(99) == "unknown_99"
    assert shape_type_name(0) == "unknown_0"


def test_shape_info_is_exported_publicly():
    assert models.ShapeInfo is ShapeInfo
    assert hasattr(ShapeInfo, "__annotations__")


def test_shape_info_fields_frozen():
    assert set(ShapeInfo.__annotations__) == set(EXPECTED_SHAPEINFO_FIELDS)


def test_shape_info_field_types():
    from typing import get_args

    a = ShapeInfo.__annotations__
    assert a["shape_id"] is int
    assert a["z_order"] is int
    assert a["coordinate_space"] == models.CoordinateSpace
    assert get_args(a["coordinate_unit"]) == ("pt",)
    assert get_args(a["group_path"]) == (int,)


def test_coordinate_space_stays_two_values():
    # 冻结：CoordinateSpace = Literal["slide", "unknown"]，不含 "group_local"
    values = [v.strip() for v in models.CoordinateSpace.__args__]
    assert set(values) == {"slide", "unknown"}


def test_import_models_does_not_load_pptx_or_win32com():
    """纯模块契约：import offipy.models 不得在 import 阶段加载 COM / python-pptx。"""
    code = (
        "import sys\n"
        "import offipy.models\n"
        "assert 'win32com' not in sys.modules, 'win32com 被提前加载'\n"
        "assert 'pythoncom' not in sys.modules, 'pythoncom 被提前加载'\n"
        "assert 'pptx' not in sys.modules, 'python-pptx 被提前加载'\n"
    )
    proc = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert proc.returncode == 0, proc.stderr
