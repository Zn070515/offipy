# tests/test_charts.py
"""原生图表声明解析测试（纯 Python，不依赖 Office）。"""

import pytest

from offipy.charts import parse_chart_declarations

ATTR_HTML = (
    "<!DOCTYPE html>\n"
    "<html><body>\n"
    '<section class="slide" data-pptx-slide>\n'
    '  <h2 class="title">营收增长</h2>\n'
    '  <div class="chart" data-chart="bar" '
    'data-chart-data=\'{"categories":["Q1","Q2","Q3"],'
    '"series":[{"name":"营收","values":[40,55,70]}]}\'></div>\n'
    "</section>\n"
    '<section class="slide" data-pptx-slide>\n'
    '  <h2 class="title">趋势</h2>\n'
    '  <div class="chart" data-chart="line"></div>\n'
    "</section>\n"
    "</body></html>"
)

SCRIPT_HTML = """<!DOCTYPE html>
<html><body>
<section class="slide" data-pptx-slide>
  <h2 class="title">占比</h2>
  <div class="chart" id="share-chart" data-chart="pie"></div>
  <script type="application/json" data-chart-target="#share-chart">
    {"categories":["A","B","C"],"series":[{"name":"份额","values":[45,35,20]}]}
  </script>
</section>
</body></html>"""


def test_parse_attribute_format():
    # 属性格式正常 → 正常解析返回声明（验证 data-chart-data 解析）
    decls = parse_chart_declarations(
        '<section class="slide" data-pptx-slide>'
        '<div class="chart" data-chart="bar" '
        'data-chart-data=\'{"categories":["Q1","Q2","Q3"],'
        '"series":[{"name":"营收","values":[40,55,70]}]}\'></div>'
        "</section>"
    )
    assert len(decls) == 1
    assert decls[0].slide_index == 1
    assert decls[0].chart_type == "bar"
    assert decls[0].data.categories == ["Q1", "Q2", "Q3"]
    assert decls[0].data.series[0].name == "营收"
    assert decls[0].data.series[0].values == [40.0, 55.0, 70.0]
    # ATTR_HTML 第二页 line 无数据 → 抛 ValueError（契约：声明了图表就必须给数据）
    with pytest.raises(ValueError):
        parse_chart_declarations(ATTR_HTML)


def test_parse_script_format():
    decls = parse_chart_declarations(SCRIPT_HTML)
    assert len(decls) == 1
    assert decls[0].slide_index == 1
    assert decls[0].chart_type == "pie"
    assert decls[0].data.categories == ["A", "B", "C"]


def test_parse_no_chart_returns_empty():
    assert parse_chart_declarations("<section data-pptx-slide><h2>x</h2></section>") == []


def test_parse_invalid_type_raises():
    with pytest.raises(ValueError):
        parse_chart_declarations(
            '<section data-pptx-slide><div class="chart" data-chart="radar" '
            'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1]]}\'></div></section>'
        )


def test_parse_bad_data_json_raises():
    with pytest.raises(ValueError):
        parse_chart_declarations(
            '<section data-pptx-slide><div class="chart" data-chart="bar" '
            'data-chart-data="{oops}"></div></section>'
        )


def test_parse_series_values_must_be_numeric():
    with pytest.raises(ValueError):
        parse_chart_declarations(
            '<section data-pptx-slide><div class="chart" data-chart="bar" '
            'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":["x"]}]}\'></div></section>'
        )


def test_parse_chart_without_data_raises():
    # 契约：打了 data-chart 的容器必须有数据（属性或 script），否则报错
    with pytest.raises(ValueError):
        parse_chart_declarations(
            '<section data-pptx-slide><div class="chart" data-chart="line"></div></section>'
        )
