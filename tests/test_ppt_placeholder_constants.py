"""PpPlaceholderType 模块常量与微软官方枚举一致（round-10 探针运行时常量 20/20 核实）。

fixture 用官方值硬编码（独立于项目常量），防「错误实现+错误测试彼此一致」。
"""

from offipy import ppt

# 微软 Learn 官方 PpPlaceholderType 完整映射（probe_placeholder_types.py 真机运行时常量实证）
MS_PP_PLACEHOLDER = {
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

# 模块暴露的常量名 → 官方枚举名（防只对数值、名字对不上）
_EXPOSED = {
    "PP_PLACEHOLDER_TITLE": "title",
    "PP_PLACEHOLDER_BODY": "body",
    "PP_PLACEHOLDER_CENTER_TITLE": "center_title",
    "PP_PLACEHOLDER_SLIDE_NUMBER": "slide_number",
    "PP_PLACEHOLDER_HEADER": "header",
    "PP_PLACEHOLDER_FOOTER": "footer",
    "PP_PLACEHOLDER_DATE": "date",
}


def test_exposed_constants_match_official_enum():
    for const_name, official_name in _EXPOSED.items():
        value = getattr(ppt, const_name)
        assert MS_PP_PLACEHOLDER[value] == official_name, (
            f"{const_name}={value} 应映射到 {official_name}，实际 {MS_PP_PLACEHOLDER[value]}"
        )


def test_exposed_constants_are_distinct():
    values = [getattr(ppt, name) for name in _EXPOSED]
    assert len(values) == len(set(values)), f"占位符常量值重复: {values}"


def test_official_reference_is_complete():
    # 官方枚举 20 个值全部在参考表内（探针核实），防参考表自身抄错
    assert len(MS_PP_PLACEHOLDER) == 20
    assert set(MS_PP_PLACEHOLDER) == set(range(1, 20)) | {-2}
