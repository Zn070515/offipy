# src/offipy/charts.py
"""原生图表：HTML 声明 → 替换为 PowerPoint 原生可编辑图表。

HTML 给图表容器打 `data-chart="<type>"`（type ∈ bar/line/pie），数据来自：
  1. 容器自身 `data-chart-data` JSON 属性（categories + series），或
  2. 页面 `<script type="application/json" data-chart-target="<css选择器>">` 块。
convert.py 转换照常跑（图表区渲染成占位形状）；转换后读 convert 审计产物
`<out>_audit/_cache/measurements.json`（含每元素 className + rect 坐标），
用 python-pptx 把图表区占位形状替换成原生图表。1 px = 6350 EMU（转换器画布
1920×1080 px = 12192000×6858000 EMU）。
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from html.parser import HTMLParser

PX_TO_EMU = 6350

CHART_TYPES = ("bar", "line", "pie")


@dataclass
class ChartSeries:
    name: str
    values: list[float]


@dataclass
class ChartData:
    categories: list[str]
    series: list[ChartSeries]


@dataclass
class ChartDecl:
    slide_index: int  # 1-based，对齐 HTML section 顺序 / convert only-slides
    chart_type: str
    data: ChartData


@dataclass
class _Container:
    slide_index: int
    chart_type: str
    data_json: str | None = None
    el_id: str | None = None


def _parse_data(json_text: str) -> ChartData:
    raw = json.loads(json_text)
    if not isinstance(raw, dict):
        raise ValueError("图表数据必须是对象 {categories, series}")
    categories = raw.get("categories")
    series_raw = raw.get("series")
    if (
        not isinstance(categories, list)
        or not categories
        or not all(isinstance(c, str) for c in categories)
    ):
        raise ValueError("图表数据 categories 必须是非空字符串列表")
    if not isinstance(series_raw, list) or not series_raw:
        raise ValueError("图表数据 series 必须是非空列表")
    series: list[ChartSeries] = []
    for sr in series_raw:
        if not isinstance(sr, dict):
            raise ValueError("图表数据 series 每项必须是对象 {name, values}")
        name = sr.get("name", "")
        values = sr.get("values")
        if (
            not isinstance(values, list)
            or not values
            or not all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in values)
        ):
            raise ValueError("图表数据 series.values 必须是非空数值列表")
        series.append(ChartSeries(name=str(name), values=[float(v) for v in values]))
    return ChartData(categories=categories, series=series)


class _ChartHTMLParser(HTMLParser):
    """提取：每页的 .chart 容器 + 页内 <script data-chart-target> 数据块。"""

    def __init__(self) -> None:
        super().__init__()
        self.slide_index = 0
        self.cur_containers: list[_Container] = []
        self.all_containers: list[_Container] = []
        self.scripts: list[tuple[str, str]] = []  # (target_selector, json_text)
        self._script_target: str | None = None
        self._script_buf: str | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        d = {k: (v or "") for k, v in attrs}
        if tag == "section" and "data-pptx-slide" in d:
            if self.cur_containers:
                self.all_containers.extend(self.cur_containers)
            self.slide_index += 1
            self.cur_containers = []
            return
        if tag == "div" and "data-chart" in d:
            cls = d.get("class", "").split()
            if "chart" in cls:
                self.cur_containers.append(
                    _Container(
                        slide_index=self.slide_index,
                        chart_type=d["data-chart"],
                        data_json=d.get("data-chart-data"),
                        el_id=d.get("id") or None,
                    )
                )
            return
        if tag == "script" and d.get("type") == "application/json" and d.get("data-chart-target"):
            self._script_target = d["data-chart-target"]
            self._script_buf = ""
            return

    def handle_endtag(self, tag: str) -> None:
        if tag == "script" and self._script_buf is not None:
            self.scripts.append((self._script_target or "", self._script_buf.strip()))
            self._script_target = None
            self._script_buf = None
            return

    def handle_data(self, data: str) -> None:
        if self._script_buf is not None:
            self._script_buf += data


def _resolve_data(container: _Container, scripts: list[tuple[str, str]]) -> ChartData:
    if container.data_json:
        try:
            return _parse_data(container.data_json)
        except (json.JSONDecodeError, ValueError) as e:
            raise ValueError(
                f"第 {container.slide_index} 页 data-chart-data 解析失败: {e}"
            ) from None
    # 从 script 块找：data-chart-target 指向该容器
    match = None
    for sel, json_text in scripts:
        if sel.startswith("#") and container.el_id == sel[1:]:
            match = json_text
            break
    if match is None:
        raise ValueError(
            f"第 {container.slide_index} 页图表 (data-chart={container.chart_type}) "
            "缺数据：请在容器加 data-chart-data，或用 <script data-chart-target> 提供"
        )
    try:
        return _parse_data(match)
    except (json.JSONDecodeError, ValueError) as e:
        raise ValueError(
            f"第 {container.slide_index} 页 <script data-chart-target> 数据解析失败: {e}"
        ) from None


def parse_chart_declarations(html_text: str) -> list[ChartDecl]:
    """解析 HTML → 图表声明列表（slide_index 1-based）。图表类型非法或数据缺失即 ValueError。"""
    p = _ChartHTMLParser()
    p.feed(html_text)
    p.close()
    # 关键：flush 最后一段 section 的容器（feed 结束时最后一页的 cur_containers 还在缓存里）
    if p.cur_containers:
        p.all_containers.extend(p.cur_containers)
    decls: list[ChartDecl] = []
    seen: set[int] = set()  # v1: 每页最多一个图表容器
    for container in p.all_containers:
        if container.chart_type not in CHART_TYPES:
            raise ValueError(
                f"第 {container.slide_index} 页图表类型非法: "
                f"{container.chart_type!r}（可选: bar/line/pie）"
            )
        if container.slide_index <= 0:
            raise ValueError(
                f"第 {container.slide_index} 页图表出现在 slide 之外——图表容器必须在 "
                "data-pptx-slide 的 <section> 内"
            )
        if container.slide_index in seen:
            raise ValueError(f"第 {container.slide_index} 页有多个图表容器——v1 每页仅支持一个图表")
        seen.add(container.slide_index)
        data = _resolve_data(container, p.scripts)
        decls.append(
            ChartDecl(slide_index=container.slide_index, chart_type=container.chart_type, data=data)
        )
    return decls


def load_chart_boxes(measurements_path: str) -> dict[int, dict]:
    """读 measurements.json → {slide_index(1-based): {"x","y","w","h"}}（px，图表容器 bbox）。

    匹配规则：记录 className 分词后恰含 "chart"（chart-note 不算）。每页取第一个匹配
    （v1 契约：每页最多一个图表容器，与 parse_chart_declarations 一致）。
    """
    import json as _json

    with open(measurements_path, encoding="utf-8") as f:
        data = _json.load(f)
    boxes: dict[int, dict] = {}
    for i, slide in enumerate(data.get("slides", []), start=1):
        for rec in slide.get("records", []):
            cls = (rec.get("className") or "").split()
            if "chart" in cls:
                boxes[i] = dict(rec["rect"])
                break
    return boxes


def _measurements_path(pptx_path: str) -> str:
    from pathlib import Path

    p = Path(pptx_path)
    return str(p.with_name(f"{p.stem}_audit") / "_cache" / "measurements.json")


def inject_native_charts(pptx_path: str, decls: list[ChartDecl], boxes: dict[int, dict]) -> None:
    """把图表区占位形状替换成原生图表。boxes 缺某页 → 该页跳过（调用方保证完整）。"""
    from pptx import Presentation
    from pptx.enum.chart import XL_CHART_TYPE

    XL_TYPE = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }
    prs = Presentation(pptx_path)
    for decl in decls:
        box = boxes.get(decl.slide_index)
        if box is None:
            continue
        slide = prs.slides[decl.slide_index - 1]
        _replace_with_chart(slide, decl, box, XL_TYPE)
    prs.save(pptx_path)


def _palette() -> list:
    """惰性构造配色（RGBColor 需 import python-pptx，顶层 import 会拖慢无图表路径）。"""
    from pptx.dml.color import RGBColor

    return [
        RGBColor(0x22, 0x51, 0xFF),  # 主蓝（对齐 mckinsey --accent）
        RGBColor(0x0E, 0x93, 0x87),  # 青
        RGBColor(0xF5, 0x9E, 0x0B),  # 琥珀
        RGBColor(0xE0, 0x5D, 0x5D),  # 珊瑚
        RGBColor(0x7C, 0x3A, 0xED),  # 紫
        RGBColor(0x66, 0x70, 0x85),  # 灰（对齐 --muted）
    ]


def _replace_with_chart(slide, decl: ChartDecl, box: dict, xl_type) -> None:
    from pptx.chart.data import CategoryChartData

    x = int(box["x"] * PX_TO_EMU)
    y = int(box["y"] * PX_TO_EMU)
    cx = int(box["w"] * PX_TO_EMU)
    cy = int(box["h"] * PX_TO_EMU)
    cxm, cym = x + cx // 2, y + cy // 2
    # 移除中心落在图表 bbox 内的占位形状（容器 surface 矩形 + bar 矩形）
    for shape in list(slide.shapes):
        if (
            shape.left <= cxm <= shape.left + shape.width
            and shape.top <= cym <= shape.top + shape.height
        ):
            slide.shapes._spTree.remove(shape._element)
    cd = CategoryChartData()
    cd.categories = decl.data.categories
    for s in decl.data.series:
        cd.add_series(s.name, s.values)
    gframe = slide.shapes.add_chart(xl_type[decl.chart_type], x, y, cx, cy, cd)
    chart = gframe.chart
    chart.has_title = False
    chart.has_legend = len(decl.data.series) > 1
    colors = _palette()
    if decl.chart_type == "bar":
        for i, series in enumerate(chart.plots[0].series):
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = colors[i % len(colors)]
    elif decl.chart_type == "line":
        for i, series in enumerate(chart.plots[0].series):
            series.format.line.color.rgb = colors[i % len(colors)]
    elif decl.chart_type == "pie":
        plot = chart.plots[0]
        for i, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = colors[i % len(colors)]


def postprocess_charts(html_path: str, pptx_path: str) -> None:
    """转换后调用：HTML 含图表声明 → 读 measurements → 注入原生图表。

    无图表声明 → 原样返回。measurements.json 缺失（--no-visual-audit）→ RuntimeError，
    让调用方知道图表不会变成原生。图表声明/数据非法 → ValueError（从 parse 上抛）。
    """
    import os

    with open(html_path, encoding="utf-8") as f:
        html_text = f.read()
    if "data-chart" not in html_text:
        return
    decls = parse_chart_declarations(html_text)
    if not decls:
        return
    meas_path = _measurements_path(pptx_path)
    if not os.path.exists(meas_path):
        raise RuntimeError(
            f"找不到 convert 审计产物 {meas_path}——图表注入需要 measurements.json，"
            "请勿用 --no-visual-audit"
        )
    boxes = load_chart_boxes(meas_path)
    missing = [d.slide_index for d in decls if d.slide_index not in boxes]
    if missing:
        raise RuntimeError(
            f'第 {missing} 页没测到图表容器（class="chart"）——请确认 deck 用了 '
            "chart-dominant 布局且 deck make 带了 --layouts"
        )
    inject_native_charts(pptx_path, decls, boxes)
