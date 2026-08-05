"""offipy 公共数据模型（v0.10 冻结）。

顶层只放 TypedDict 与纯映射常量，**不 import python-pptx / COM**——任何平台都能
``import offipy.models``。返回统一走 TypedDict（``server._serialize`` 对 dict/list/float
原样透传，JSON 安全）。

v0.10 只冻结并发布 ``SlideTextRecord``。``ShapeInfo`` / ``ShapeRef`` 是 v0.12 增量编辑
的草案（只写进 ADR / Milestone 文档），不在此导出（P0-3）。
"""

from typing import Literal, TypedDict

# PpPlaceholderType 官方值（微软 Learn，round-10 探针用真 PowerPoint 运行时常量 20/20 核实）。
# 未知类型不返回 None，回 f"unknown_{n}"。
PLACEHOLDER_TYPE_NAMES = {
    -2: "mixed",
    1: "title",
    2: "body",
    3: "center_title",
    4: "subtitle",
    5: "vertical_title",
    6: "vertical_body",
    7: "object",
    8: "chart",
    9: "bitmap",
    10: "media_clip",
    11: "org_chart",
    12: "table",
    13: "slide_number",
    14: "header",
    15: "footer",
    16: "date",
    17: "vertical_object",
    18: "picture",
    19: "cameo",
}


def placeholder_type_name(value: int | None) -> str | None:
    """PpPlaceholderType 数值 → 名称；未知数值回 f"unknown_{n}"，None 回 None。"""
    if value is None:
        return None
    return PLACEHOLDER_TYPE_NAMES.get(value, f"unknown_{value}")


# 坐标语义（round-10 探针实证，见 KB ADR 79）：
# - 非旋转 group（rotation%90==0）的 GroupItems(i).Left/Top/Width/Height 返回
#   **当前绝对幻灯片坐标**（含移动/缩放的当前渲染位置），直接读即 slide 空间。
# - 旋转 group（rotation%90!=0）后读值不再是绝对坐标（实测旋转 30° 后子元素
#   top 读 0.0，在幻灯片上不可能）——要真绝对坐标需读 OOXML group transform
#   （off/ext/chOff/chExt）做完整 affine transform，v0.10 不做，标注 "unknown"。
# - 从未观察到 PowerPoint 返回父 group 局部坐标（"group_local"），故 Literal 收敛为
#   slide / unknown，unknown 作一切不可信读值的兜底。
CoordinateSpace = Literal["slide", "unknown"]


class SlideTextRecord(TypedDict):
    """单条可见文本 shape 的记录（``read_slide_texts`` 的返回元素）。

    - 只含**具有文本能力**的 shape（有 TextFrame）；图片/线条/无文本图形不在此列
      （那是 ``read_shapes`` 的职责，v0.10 不提供）。
    - 坐标单位恒为磅（pt）；``coordinate_space`` 说明坐标参照系。
    - group 内文本：``parent_shape_id`` 是直接父 group；``group_path`` 是祖先 group
      链（从外层到内层）。
    """

    shape_id: int
    name: str
    text: str
    left: float
    top: float
    width: float
    height: float
    coordinate_space: CoordinateSpace
    coordinate_unit: Literal["pt"]
    is_placeholder: bool
    placeholder_type: int | None
    placeholder_type_name: str | None
    parent_shape_id: int | None
    group_path: list[int]
