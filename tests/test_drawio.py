"""offipy.drawio — draw.io 解析（vendored drawio_extract 包装）。"""

from __future__ import annotations

import subprocess
import sys

import pytest

from offipy.drawio import parse_drawio

SINGLE_PAGE = """\
<mxfile>
  <diagram name="Page 1">
    <mxGraphModel>
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
        <mxCell id="g" value="容器"
                style="rounded=1;fillColor=#fff2cc;strokeColor=#d6b656;container=1;"
                vertex="1" parent="1">
          <mxGeometry x="20" y="20" width="340" height="200" as="geometry"/>
        </mxCell>
        <mxCell id="a" value="节点A"
                style="rounded=0;fillColor=#dae8fc;strokeColor=#6c8ebf;fontColor=#000000;"
                vertex="1" parent="g">
          <mxGeometry x="40" y="40" width="120" height="60" as="geometry"/>
        </mxCell>
        <mxCell id="b" value="节点B"
                style="rounded=0;fillColor=#d5e8d4;strokeColor=#82b366;"
                vertex="1" parent="g">
          <mxGeometry x="200" y="40" width="120" height="60" as="geometry"/>
        </mxCell>
        <mxCell id="c" value="节点C"
                style="rounded=1;fillColor=#ffe6cc;strokeColor=#d79b00;"
                vertex="1" parent="1">
          <mxGeometry x="60" y="260" width="140" height="60" as="geometry"/>
        </mxCell>
        <mxCell id="e1" value="连接"
                style="edgeStyle=orthogonalEdgeStyle;rounded=1;strokeColor=#333333;dashed=1;"
                edge="1" source="a" target="b" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
"""

MULTI_PAGE = """\
<mxfile>
  <diagram name="第一页">
    <mxGraphModel><root>
      <mxCell id="0"/><mxCell id="1" parent="0"/>
      <mxCell id="x" value="X" vertex="1" parent="1">
        <mxGeometry x="10" y="10" width="100" height="50" as="geometry"/>
      </mxCell>
    </root></mxGraphModel>
  </diagram>
  <diagram name="第二页">
    <mxGraphModel><root>
      <mxCell id="0"/><mxCell id="1" parent="0"/>
      <mxCell id="y" value="Y" vertex="1" parent="1">
        <mxGeometry x="30" y="30" width="100" height="50" as="geometry"/>
      </mxCell>
    </root></mxGraphModel>
  </diagram>
</mxfile>
"""

SHAPES = """\
<mxfile><diagram name="P"><mxGraphModel><root>
  <mxCell id="0"/><mxCell id="1" parent="0"/>
  <mxCell id="ra" value="RA" style="rounded=0;" vertex="1" parent="1">
    <mxGeometry x="0" y="0" width="50" height="30" as="geometry"/>
  </mxCell>
  <mxCell id="rb" value="RB" style="rounded=1;" vertex="1" parent="1">
    <mxGeometry x="80" y="0" width="50" height="30" as="geometry"/>
  </mxCell>
  <mxCell id="el" value="EL" style="ellipse;" vertex="1" parent="1">
    <mxGeometry x="160" y="0" width="50" height="30" as="geometry"/>
  </mxCell>
  <mxCell id="tr" value="TR" style="triangle;" vertex="1" parent="1">
    <mxGeometry x="240" y="0" width="50" height="30" as="geometry"/>
  </mxCell>
</root></mxGraphModel></diagram></mxfile>
"""


def _write(tmp_path, text=SINGLE_PAGE, name="input.drawio"):
    p = tmp_path / name
    p.write_text(text, encoding="utf-8")
    return str(p)


def test_parse_drawio_basic(tmp_path):
    d = parse_drawio(_write(tmp_path))
    by_id = {n.id: n for n in d.nodes}
    assert set(by_id) == {"g", "a", "b", "c"}
    a = by_id["a"]
    assert a.x == 60.0 and a.y == 60.0  # 父链累加：g(20,20) + a(40,40)
    assert a.w == 120.0 and a.h == 60.0
    assert a.fill == "#dae8fc"
    assert a.stroke == "#6c8ebf"
    assert a.font_color == "#000000"
    assert a.rounded is False  # rounded=0 → 直角
    g = by_id["g"]
    assert g.container is True
    assert g.label == "容器"
    assert g.x == 20.0 and g.y == 20.0
    c = by_id["c"]
    assert c.rounded is True  # rounded=1 → 圆角
    assert len(d.edges) == 1
    e = d.edges[0]
    assert e.source == "a" and e.target == "b"
    assert e.label == "连接"
    assert e.dashed is True
    assert e.stroke == "#333333"


def test_parse_drawio_default_first_page(tmp_path):
    d = parse_drawio(_write(tmp_path, MULTI_PAGE, "multi.drawio"))
    assert [n.id for n in d.nodes] == ["x"]


def test_parse_drawio_page_by_index(tmp_path):
    src = _write(tmp_path, MULTI_PAGE, "multi.drawio")
    assert [n.id for n in parse_drawio(src, page=0).nodes] == ["x"]
    assert [n.id for n in parse_drawio(src, page=1).nodes] == ["y"]
    assert [n.id for n in parse_drawio(src, page="0").nodes] == ["x"]


def test_parse_drawio_page_by_name(tmp_path):
    src = _write(tmp_path, MULTI_PAGE, "multi.drawio")
    assert [n.id for n in parse_drawio(src, page="第二页").nodes] == ["y"]


def test_parse_drawio_missing_file(tmp_path):
    with pytest.raises(FileNotFoundError):
        parse_drawio(str(tmp_path / "nope.drawio"))


def test_parse_drawio_bad_source(tmp_path):
    p = tmp_path / "bad.drawio"
    p.write_text("not draw.io at all", encoding="utf-8")
    with pytest.raises(ValueError, match="无法解析 draw.io 源码"):
        parse_drawio(str(p))


def test_parse_drawio_empty_diagram(tmp_path):
    xml = (
        '<mxfile><diagram name="P"><mxGraphModel><root>'
        '<mxCell id="0"/><mxCell id="1" parent="0"/>'
        "</root></mxGraphModel></diagram></mxfile>"
    )
    with pytest.raises(ValueError, match="未提取到任何节点或边"):
        parse_drawio(_write(tmp_path, xml, "empty.drawio"))


def test_parse_drawio_dangling_edge_filtered(tmp_path):
    xml = (
        '<mxfile><diagram name="P"><mxGraphModel><root>'
        '<mxCell id="0"/><mxCell id="1" parent="0"/>'
        '<mxCell id="a" value="A" vertex="1" parent="1">'
        '<mxGeometry x="0" y="0" width="50" height="30" as="geometry"/></mxCell>'
        '<mxCell id="e1" value="dangle" style="edgeStyle=orthogonalEdgeStyle;" '
        'edge="1" source="a" target="ghost" parent="1">'
        '<mxGeometry relative="1" as="geometry"/></mxCell>'
        "</root></mxGraphModel></diagram></mxfile>"
    )
    d = parse_drawio(_write(tmp_path, xml, "dangle.drawio"))
    assert [n.id for n in d.nodes] == ["a"]
    assert d.edges == []  # 悬空边（target 指向不存在的节点）被过滤


def test_drawio_lazy_import_no_pptx():
    """drawio.py 顶层不得 import pptx（惰性 import 红线）。"""
    code = (
        "import sys\n"
        "import offipy.drawio\n"
        "assert 'pptx' not in sys.modules, 'drawio.py 顶层不得 import pptx'\n"
    )
    subprocess.run([sys.executable, "-c", code], check=True)
