# tests/test_charts.py
"""原生图表声明解析测试（纯 Python，不依赖 Office）。"""

import json

import pytest
from pptx import Presentation

from offipy import charts
from offipy.charts import (
    ChartData,
    ChartDecl,
    ChartSeries,
    inject_native_charts,
    load_chart_boxes,
    parse_chart_declarations,
)

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


def _blank_pptx(tmp_path):
    from pptx.util import Emu

    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    prs.slides.add_slide(prs.slide_layouts[6])
    return prs


def test_load_chart_boxes(tmp_path):
    meas = tmp_path / "measurements.json"
    meas.write_text(
        json.dumps(
            {
                "slides": [
                    {
                        "slide": {"theme": "mckinsey"},
                        "records": [
                            {
                                "id": 1,
                                "kind": "text",
                                "tag": "h2",
                                "className": "title",
                                "rect": {"x": 96, "y": 96, "w": 800, "h": 80},
                                "text": "营收",
                            },
                            {
                                "id": 2,
                                "kind": "shape",
                                "tag": "div",
                                "className": "chart",
                                "rect": {"x": 96, "y": 260, "w": 900, "h": 500},
                                "text": "",
                            },
                            {
                                "id": 3,
                                "kind": "text",
                                "tag": "div",
                                "className": "chart-note",
                                "rect": {"x": 1020, "y": 260, "w": 360, "h": 500},
                                "text": "来源",
                            },
                        ],
                    }
                ]
            }
        ),
        encoding="utf-8",
    )
    boxes = load_chart_boxes(str(meas))
    assert boxes[1] == {"x": 96, "y": 260, "w": 900, "h": 500}
    assert 2 not in boxes  # 只认 className 恰为 "chart" 的容器，chart-note 不匹配


def test_inject_native_chart_bar(tmp_path):
    prs = _blank_pptx(tmp_path)
    slide = prs.slides[0]
    from pptx.util import Emu

    # 放一个占位矩形（模拟 convert 渲染的图表区），再放一个右上角说明（不该被删）
    slide.shapes.add_shape(1, Emu(96 * 6350), Emu(260 * 6350), Emu(900 * 6350), Emu(500 * 6350))
    note = slide.shapes.add_textbox(
        Emu(1020 * 6350), Emu(260 * 6350), Emu(360 * 6350), Emu(500 * 6350)
    )
    note.text = "来源"
    out = tmp_path / "out.pptx"
    prs.save(str(out))

    decls = [
        ChartDecl(
            slide_index=1,
            chart_type="bar",
            data=ChartData(
                categories=["Q1", "Q2"],
                series=[ChartSeries(name="营收", values=[40, 70])],
            ),
        )
    ]
    boxes = {1: {"x": 96, "y": 260, "w": 900, "h": 500}}
    inject_native_charts(str(out), decls, boxes)

    prs2 = Presentation(str(out))
    slide2 = prs2.slides[0]
    charts = [s for s in slide2.shapes if s.has_chart]
    assert len(charts) == 1
    ch = charts[0].chart
    assert ch.chart_type is not None
    assert ch.has_title is False
    # 占位矩形被移除，说明文本保留
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    rects = [s for s in slide2.shapes if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    assert len(rects) == 0
    texts = [s.text for s in slide2.shapes if s.has_text_frame]
    assert "来源" in texts


def test_inject_native_chart_pie(tmp_path):
    prs = _blank_pptx(tmp_path)
    out = tmp_path / "pie.pptx"
    prs.save(str(out))
    decls = [
        ChartDecl(
            slide_index=1,
            chart_type="pie",
            data=ChartData(
                categories=["A", "B"],
                series=[ChartSeries(name="份额", values=[45, 55])],
            ),
        )
    ]
    inject_native_charts(str(out), decls, {1: {"x": 100, "y": 200, "w": 800, "h": 500}})
    prs2 = Presentation(str(out))
    charts = [s for s in prs2.slides[0].shapes if s.has_chart]
    assert len(charts) == 1


def test_parse_top_level_non_dict_raises():
    with pytest.raises(ValueError):
        parse_chart_declarations(
            '<section data-pptx-slide><div class="chart" data-chart="bar" '
            'data-chart-data="[1,2,3]"></div></section>'
        )


def test_parse_chart_outside_slide_raises():
    # 图表容器出现在第一个 <section data-pptx-slide> 之前 → slide_index=0 → 报错而非静默
    with pytest.raises(ValueError):
        parse_chart_declarations(
            '<div class="chart" data-chart="bar" '
            'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1]}]}\'></div>'
            "<section data-pptx-slide><h2>x</h2></section>"
        )


def test_postprocess_skips_without_charts(tmp_path, monkeypatch):
    # HTML 无 data-chart → 不读 measurements、不改 pptx
    called = {}
    monkeypatch.setattr(charts, "load_chart_boxes", lambda *a, **k: called.setdefault("load", True))
    html = tmp_path / "d.html"
    html.write_text("<section data-pptx-slide><h2>x</h2></section>", encoding="utf-8")
    pptx = tmp_path / "d.pptx"
    pptx.write_bytes(b"not a real pptx")
    charts.postprocess_charts(str(html), str(pptx))  # 不抛异常、load 不被调用
    assert "load" not in called


def test_postprocess_missing_measurements_raises(tmp_path):
    html = tmp_path / "d.html"
    html.write_text(
        '<section data-pptx-slide><div class="chart" data-chart="bar" '
        'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1]}]}\'></div></section>',
        encoding="utf-8",
    )
    pptx = tmp_path / "d.pptx"
    pptx.write_bytes(b"x")
    with pytest.raises(RuntimeError, match="measurements"):
        charts.postprocess_charts(str(html), str(pptx))


def test_postprocess_calls_inject(monkeypatch, tmp_path):
    from pptx.util import Emu

    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    prs.slides.add_slide(prs.slide_layouts[6])
    pptx = tmp_path / "d.pptx"
    prs.save(str(pptx))
    html = tmp_path / "d.html"
    html.write_text(
        '<section data-pptx-slide><div class="chart" data-chart="bar" '
        'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1,2]}]}\'></div></section>',
        encoding="utf-8",
    )
    meas_dir = tmp_path / "d_audit" / "_cache"
    meas_dir.mkdir(parents=True)
    (meas_dir / "measurements.json").write_text(
        json.dumps(
            {
                "slides": [
                    {
                        "slide": {},
                        "records": [
                            {
                                "id": 1,
                                "kind": "shape",
                                "tag": "div",
                                "className": "chart",
                                "rect": {"x": 100, "y": 200, "w": 800, "h": 500},
                                "text": "",
                            }
                        ],
                    }
                ]
            }
        ),
        encoding="utf-8",
    )

    charts.postprocess_charts(str(html), str(pptx))

    prs2 = Presentation(str(pptx))
    assert any(s.has_chart for s in prs2.slides[0].shapes)
