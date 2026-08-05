"""几何基础：Affine2D 仿射变换、AABB 矩形与交集（纯函数，零第三方依赖）。

Affine2D 是 2x3 矩阵 [a c tx; b d ty]，点变换 x'=a*x+c*y+tx, y'=b*x+d*y+ty。
compose(other) 返回 self∘other：**先应用 other，再应用 self**。

旋转采用数学 CCW 约定；本模块只用于 AABB 计算（bounds/margin/overlap），
矩形绕中心旋转的包围盒对旋转方向不敏感，因此方向符号不影响结果。
"""

from __future__ import annotations

import math
from dataclasses import dataclass

# ---------------------------------------------------------------- 矩形


@dataclass(frozen=True)
class Rect:
    """轴对齐矩形（AABB），坐标与尺寸单位英寸。"""

    x: float
    y: float
    width: float
    height: float

    @property
    def right(self) -> float:
        return self.x + self.width

    @property
    def bottom(self) -> float:
        return self.y + self.height

    def center(self) -> tuple[float, float]:
        return (self.x + self.width / 2.0, self.y + self.height / 2.0)

    def area(self) -> float:
        return self.width * self.height


def rect_intersection(a: Rect, b: Rect) -> Rect | None:
    """两 AABB 交集；无交集返回 None。"""
    x0 = max(a.x, b.x)
    y0 = max(a.y, b.y)
    x1 = min(a.right, b.right)
    y1 = min(a.bottom, b.bottom)
    if x1 <= x0 or y1 <= y0:
        return None
    return Rect(x0, y0, x1 - x0, y1 - y0)


def overlap_area(a: Rect, b: Rect) -> float:
    """重叠面积（无交集为 0）。"""
    inter = rect_intersection(a, b)
    return inter.area() if inter is not None else 0.0


def rect_contains(outer: Rect, inner: Rect, *, eps: float = 1e-9) -> bool:
    """outer 是否完整包含 inner（含边界，带 eps 容差）。"""
    return (
        outer.x - eps <= inner.x
        and outer.y - eps <= inner.y
        and inner.right - eps <= outer.right
        and inner.bottom - eps <= outer.bottom
    )


# ---------------------------------------------------------------- 仿射


class Affine2D:
    """2x3 仿射矩阵 [a c tx; b d ty]。"""

    __slots__ = ("a", "b", "c", "d", "tx", "ty")

    def __init__(self, a: float, b: float, c: float, d: float, tx: float, ty: float) -> None:
        self.a = a
        self.b = b
        self.c = c
        self.d = d
        self.tx = tx
        self.ty = ty

    @staticmethod
    def identity() -> Affine2D:
        return Affine2D(1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

    @staticmethod
    def translate(tx: float, ty: float) -> Affine2D:
        return Affine2D(1.0, 0.0, 0.0, 1.0, tx, ty)

    @staticmethod
    def scale(sx: float, sy: float) -> Affine2D:
        return Affine2D(sx, 0.0, 0.0, sy, 0.0, 0.0)

    @staticmethod
    def rotate(deg: float) -> Affine2D:
        """绕原点旋转（数学 CCW；AABB 用途下方向无关）。"""
        rad = math.radians(deg)
        cos_a = math.cos(rad)
        sin_a = math.sin(rad)
        return Affine2D(cos_a, sin_a, -sin_a, cos_a, 0.0, 0.0)

    @staticmethod
    def flip_h() -> Affine2D:
        """水平翻转：x → -x（镜像于竖直轴）。"""
        return Affine2D(-1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

    @staticmethod
    def flip_v() -> Affine2D:
        """垂直翻转：y → -y（镜像于水平轴）。"""
        return Affine2D(1.0, 0.0, 0.0, -1.0, 0.0, 0.0)

    def compose(self, other: Affine2D) -> Affine2D:
        """返回 self∘other：先应用 other，再应用 self（标准矩阵乘法）。"""
        return Affine2D(
            self.a * other.a + self.c * other.b,
            self.b * other.a + self.d * other.b,
            self.a * other.c + self.c * other.d,
            self.b * other.c + self.d * other.d,
            self.a * other.tx + self.c * other.ty + self.tx,
            self.b * other.tx + self.d * other.ty + self.ty,
        )

    def transform_point(self, x: float, y: float) -> tuple[float, float]:
        return (self.a * x + self.c * y + self.tx, self.b * x + self.d * y + self.ty)

    def transform_rect(self, rect: Rect) -> Rect:
        """四角变换后取 AABB。"""
        corners = (
            (rect.x, rect.y),
            (rect.right, rect.y),
            (rect.right, rect.bottom),
            (rect.x, rect.bottom),
        )
        xs: list[float] = []
        ys: list[float] = []
        for x, y in corners:
            px, py = self.transform_point(x, y)
            xs.append(px)
            ys.append(py)
        return Rect(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
