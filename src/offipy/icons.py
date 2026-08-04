# src/offipy/icons.py
"""原生图标：内联 <svg data-icon> → PowerPoint freeform 矢量图标。

HTML 给图标容器打空 `<svg class="icon" data-icon="<set>:<name>" viewBox="..."
width=".." height=".."></svg>`（path 由注入生成）。convert 照常跑（SVG 截图成
PNG 占位）；转换后读 `<out>_audit/_cache/measurements.json`（kind='svg' record
含 outerHTML/rect/color），用 python-pptx freeform 把占位替换成矢量图标并删
占位。1 px = 6350 EMU（转换器画布 1920×1080 px = 12192000×6858000 EMU）。

图标集（vendored 到 src/offipy/assets/icons/）：
  ph: Phosphor fill 权重（256 viewBox，单 path fill="currentColor"，fill 模式）
  lu: Lucide（24 viewBox，fill="none" stroke="currentColor"，stroke 模式）
图标颜色取 measurements 里 svg 的计算 color（HTML 设 color: var(--accent)）。

顶层不 import python-pptx（镜像 charts.py）：from pptx... 全在函数内惰性 import。
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from pathlib import Path

PX_TO_EMU = 6350
ASSETS_DIR = Path(__file__).resolve().parent / "assets" / "icons"

# data-icon 前缀 → assets 子目录 / 渲染模式
SET_DIRS = {"ph": "phosphor", "lu": "lucide"}
SET_MODE = {"ph": "fill", "lu": "stroke"}

_CURVE_SAMPLES = 12  # 贝塞尔压平采样段数
_CIRCLE_SAMPLES = 32  # circle/ellipse 多边形近似段数
_ARC_SAMPLES = 24  # 圆弧压平采样段数

_TOKEN_RE = re.compile(r"[A-Za-z]|-?\d*\.?\d+(?:[eE][-+]?\d+)?")
_NUM_RE = re.compile(r"-?\d*\.?\d+(?:[eE][-+]?\d+)?")


@dataclass
class _SubPath:
    points: list[tuple[float, float]]
    close: bool


@dataclass
class IconDecl:
    slide_index: int  # 1-based，对齐 HTML section 顺序
    data_icon: str  # "ph:name"
    view_box: tuple[float, float, float, float]  # (x, y, w, h)


def _tokenize(d: str) -> list[str]:
    return _TOKEN_RE.findall(d)


def _cubic_points(p0, p1, p2, p3, n: int) -> list[tuple[float, float]]:
    pts = [p0]
    for i in range(1, n + 1):
        t = i / n
        mt = 1 - t
        x = mt**3 * p0[0] + 3 * mt**2 * t * p1[0] + 3 * mt * t**2 * p2[0] + t**3 * p3[0]
        y = mt**3 * p0[1] + 3 * mt**2 * t * p1[1] + 3 * mt * t**2 * p2[1] + t**3 * p3[1]
        pts.append((x, y))
    return pts


def _quad_points(p0, p1, p2, n: int) -> list[tuple[float, float]]:
    pts = [p0]
    for i in range(1, n + 1):
        t = i / n
        mt = 1 - t
        x = mt**2 * p0[0] + 2 * mt * t * p1[0] + t**2 * p2[0]
        y = mt**2 * p0[1] + 2 * mt * t * p1[1] + t**2 * p2[1]
        pts.append((x, y))
    return pts


def _arc_points(p0, radii, rot, large, sweep, p1, n: int) -> list[tuple[float, float]]:
    """SVG 椭圆弧 → 折线（endpoint→center 参数化，F.6.5）。"""
    rx, ry = abs(radii[0]), abs(radii[1])
    x0, y0 = p0
    x1, y1 = p1
    if (x0, y0) == (x1, y1) or rx == 0 or ry == 0:
        return [p0]
    phi = math.radians(rot)
    cos_p, sin_p = math.cos(phi), math.sin(phi)
    dx, dy = (x0 - x1) / 2, (y0 - y1) / 2
    xp = cos_p * dx + sin_p * dy
    yp = -sin_p * dx + cos_p * dy
    lam = xp * xp / (rx * rx) + yp * yp / (ry * ry)
    if lam > 1:
        s = math.sqrt(lam)
        rx *= s
        ry *= s
    sign = -1.0 if large == sweep else 1.0
    num = rx * rx * ry * ry - rx * rx * yp * yp - ry * ry * xp * xp
    den = rx * rx * yp * yp + ry * ry * xp * xp
    coef = sign * math.sqrt(max(0.0, num / den))
    cxp = coef * (rx * yp / ry)
    cyp = coef * (-ry * xp / rx)
    cx = cos_p * cxp - sin_p * cyp + (x0 + x1) / 2
    cy = sin_p * cxp + cos_p * cyp + (y0 + y1) / 2

    def angle(ux, uy, vx, vy):
        dot = ux * vx + uy * vy
        m = math.hypot(ux, uy) * math.hypot(vx, vy)
        a = math.acos(max(-1.0, min(1.0, dot / m)))
        if ux * vy - uy * vx < 0:
            a = -a
        return a

    theta1 = angle(1, 0, (xp - cxp) / rx, (yp - cyp) / ry)
    dtheta = angle((xp - cxp) / rx, (yp - cyp) / ry, (-xp - cxp) / rx, (-yp - cyp) / ry)
    if sweep == 0 and dtheta > 0:
        dtheta -= 2 * math.pi
    elif sweep == 1 and dtheta < 0:
        dtheta += 2 * math.pi
    pts = [p0]
    for i in range(1, n + 1):
        a = theta1 + dtheta * (i / n)
        ca, sa = math.cos(a), math.sin(a)
        pts.append((cos_p * rx * ca - sin_p * ry * sa + cx, sin_p * rx * ca + cos_p * ry * sa + cy))
    pts[-1] = p1  # 终点由 SVG 规范保证精确等于 p1；采样末点存在浮点漂移，收拢到 p1
    return pts


def _parse_path(d: str) -> list[_SubPath]:
    """SVG path d 属性 → 压平后的子路径列表。曲线命令已折线化。"""
    toks = _tokenize(d)
    subpaths: list[_SubPath] = []
    pts: list[tuple[float, float]] = []
    close = False
    cx = cy = 0.0
    start = (0.0, 0.0)
    cmd: str | None = None
    prev_cubic: tuple[float, float] | None = None
    prev_quad: tuple[float, float] | None = None
    i = 0
    n = len(toks)

    def read_num() -> float:
        nonlocal i
        v = float(toks[i])
        i += 1
        return v

    def read_point(rel: bool) -> tuple[float, float]:
        nonlocal cx, cy
        x, y = read_num(), read_num()
        if rel:
            x += cx
            y += cy
        return x, y

    while i < n:
        t = toks[i]
        if _NUM_RE.fullmatch(t):
            if cmd is None:
                raise ValueError(f"SVG path 首 token 必须是命令字母: {d!r}")
        else:
            cmd = t
            i += 1
        if cmd is None:
            raise ValueError(f"SVG path 无命令: {d!r}")
        c = cmd.upper()
        rel = cmd.islower()
        if c == "M":
            if pts:
                subpaths.append(_SubPath(pts, close))
            p = read_point(rel)
            pts = [p]
            start = p
            cx, cy = p
            cmd = "l" if rel else "L"  # M 后隐式点 → lineto
            prev_cubic = prev_quad = None
            close = False
        elif c == "L":
            p = read_point(rel)
            pts.append(p)
            cx, cy = p
            prev_cubic = prev_quad = None
        elif c == "H":
            x = read_num()
            cx = cx + x if rel else x
            pts.append((cx, cy))
            prev_cubic = prev_quad = None
        elif c == "V":
            y = read_num()
            cy = cy + y if rel else y
            pts.append((cx, cy))
            prev_cubic = prev_quad = None
        elif c == "C":
            p1 = read_point(rel)
            p2 = read_point(rel)
            p3 = read_point(rel)
            for p in _cubic_points((cx, cy), p1, p2, p3, _CURVE_SAMPLES)[1:]:
                pts.append(p)
            cx, cy = p3
            prev_cubic = p2
            prev_quad = None
        elif c == "S":
            p1 = (2 * cx - prev_cubic[0], 2 * cy - prev_cubic[1]) if prev_cubic else (cx, cy)
            p2 = read_point(rel)
            p3 = read_point(rel)
            for p in _cubic_points((cx, cy), p1, p2, p3, _CURVE_SAMPLES)[1:]:
                pts.append(p)
            cx, cy = p3
            prev_cubic = p2
            prev_quad = None
        elif c == "Q":
            p1 = read_point(rel)
            p2 = read_point(rel)
            for p in _quad_points((cx, cy), p1, p2, _CURVE_SAMPLES)[1:]:
                pts.append(p)
            cx, cy = p2
            prev_quad = p1
            prev_cubic = None
        elif c == "T":
            p1 = (2 * cx - prev_quad[0], 2 * cy - prev_quad[1]) if prev_quad else (cx, cy)
            p2 = read_point(rel)
            for p in _quad_points((cx, cy), p1, p2, _CURVE_SAMPLES)[1:]:
                pts.append(p)
            cx, cy = p2
            prev_quad = p1
            prev_cubic = None
        elif c == "A":
            rx = read_num()
            ry = read_num()
            rot = read_num()
            large = read_num()
            sweep = read_num()
            p2 = read_point(rel)
            for p in _arc_points((cx, cy), (rx, ry), rot, large, sweep, p2, _ARC_SAMPLES)[1:]:
                pts.append(p)
            cx, cy = p2
            prev_cubic = prev_quad = None
        elif c == "Z":
            close = True
            subpaths.append(_SubPath(pts, close))
            pts = []
            cx, cy = start
            cmd = None
            prev_cubic = prev_quad = None
        else:
            raise ValueError(f"SVG path 非法命令 {cmd!r}: {d!r}")
    if pts:
        subpaths.append(_SubPath(pts, close))
    if not subpaths:
        raise ValueError(f"SVG path 无有效子路径: {d!r}")
    return subpaths


def _parse_points_list(s: str) -> list[tuple[float, float]]:
    nums = [float(v) for v in re.findall(_NUM_RE, s)]
    return list(zip(nums[0::2], nums[1::2], strict=True))


def _geom_points(tag: str, attrs: dict) -> tuple[list[tuple[float, float]], bool]:
    """基本几何元素 → (顶点列表, 是否闭合)。"""
    if tag == "line":
        x1 = float(attrs.get("x1", 0))
        y1 = float(attrs.get("y1", 0))
        x2 = float(attrs.get("x2", 0))
        y2 = float(attrs.get("y2", 0))
        return [(x1, y1), (x2, y2)], False
    if tag == "polyline":
        return _parse_points_list(attrs.get("points", "")), False
    if tag == "polygon":
        return _parse_points_list(attrs.get("points", "")), True
    if tag == "rect":
        x = float(attrs.get("x", 0))
        y = float(attrs.get("y", 0))
        w = float(attrs.get("width", 0))
        h = float(attrs.get("height", 0))
        return [(x, y), (x + w, y), (x + w, y + h), (x, y + h)], True
    if tag == "circle":
        cx = float(attrs.get("cx", 0))
        cy = float(attrs.get("cy", 0))
        r = float(attrs.get("r", 0))
        return [
            (
                cx + r * math.cos(2 * math.pi * i / _CIRCLE_SAMPLES),
                cy + r * math.sin(2 * math.pi * i / _CIRCLE_SAMPLES),
            )
            for i in range(_CIRCLE_SAMPLES)
        ], True
    if tag == "ellipse":
        cx = float(attrs.get("cx", 0))
        cy = float(attrs.get("cy", 0))
        rx = float(attrs.get("rx", 0))
        ry = float(attrs.get("ry", 0))
        return [
            (
                cx + rx * math.cos(2 * math.pi * i / _CIRCLE_SAMPLES),
                cy + ry * math.sin(2 * math.pi * i / _CIRCLE_SAMPLES),
            )
            for i in range(_CIRCLE_SAMPLES)
        ], True
    raise ValueError(f"不支持的 SVG 元素: {tag!r}")


def load_icon_svg(data_icon: str) -> str:
    """从 vendored assets 读图标 SVG 源码。"ph:name" → assets/phosphor/name.svg。"""
    if ":" not in data_icon:
        raise ValueError(f"图标必须带集前缀 ph:/lu:，如 'ph:check': {data_icon!r}")
    set_, name = data_icon.split(":", 1)
    if set_ not in SET_DIRS:
        raise ValueError(f"未知图标集 {set_!r}（可选: {', '.join(SET_DIRS)}）")
    if not re.fullmatch(r"[a-z0-9][a-z0-9-]*", name):
        raise ValueError(f"非法图标名 {name!r}（限小写字母/数字/连字符）")
    path = ASSETS_DIR / SET_DIRS[set_] / f"{name}.svg"
    if not path.exists():
        raise ValueError(
            f"图标 {data_icon} 不在资产库（{path}）——先跑 uv run python "
            "scripts/fetch_icons.py 抓取全量图标"
        )
    return path.read_text(encoding="utf-8")


def _svg_to_subpaths(svg_text: str) -> tuple[list[_SubPath], float]:
    """资产 SVG 源码 → (子路径列表, stroke-width)。子路径已压平。"""
    import xml.etree.ElementTree as ET

    root = ET.fromstring(svg_text)
    stroke_width = float(root.get("stroke-width", "2"))
    subpaths: list[_SubPath] = []
    for el in root.iter():
        tag = el.tag.split("}")[-1].lower()
        if tag == "path":
            d = el.get("d")
            if d:
                subpaths.extend(_parse_path(d))
        elif tag in ("line", "polyline", "polygon", "rect", "circle", "ellipse"):
            pts, close = _geom_points(tag, el.attrib)
            if pts:
                subpaths.append(_SubPath(pts, close))
    return subpaths, stroke_width
