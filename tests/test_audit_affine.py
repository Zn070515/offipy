"""audit 几何：Affine2D 变换/复合/包围盒，Rect 交集/包含。"""

import math

import pytest

from offipy.audit.geometry import (
    Affine2D,
    Rect,
    overlap_area,
    rect_contains,
    rect_intersection,
)

# ---------------------------------------------------------------- 基础变换


def test_translate():
    m = Affine2D.translate(3.0, -2.0)
    assert m.transform_point(1, 1) == (4.0, -1.0)


def test_scale():
    m = Affine2D.scale(2.0, 0.5)
    assert m.transform_point(1, 4) == (2.0, 2.0)


def test_rotate_90():
    m = Affine2D.rotate(90.0)
    px, py = m.transform_point(1, 0)
    assert px == pytest.approx(0.0, abs=1e-9)
    assert py == pytest.approx(1.0, abs=1e-9)


def test_rotate_45_aabb_grows():
    # 1×1 矩形绕中心旋转 45°，AABB 边长 = sqrt(2)
    m = Affine2D.identity()
    center = Rect(0, 0, 1, 1)
    # 直接变换角点取 AABB
    rect = m.compose(Affine2D.translate(0.5, 0.5))
    rect = rect.compose(Affine2D.rotate(45.0))
    rect = rect.compose(Affine2D.translate(-0.5, -0.5))
    out = rect.transform_rect(center)
    assert out.width == pytest.approx(math.sqrt(2), abs=1e-9)
    assert out.height == pytest.approx(math.sqrt(2), abs=1e-9)


def test_flip_h():
    m = Affine2D.flip_h()
    assert m.transform_point(1, 0) == (-1.0, 0.0)


def test_flip_v():
    m = Affine2D.flip_v()
    assert m.transform_point(0, 1) == (0.0, -1.0)


def test_identity():
    m = Affine2D.identity()
    assert m.transform_point(5, -3) == (5.0, -3.0)


def test_compose_applies_other_first():
    # translate(2,0) ∘ translate(0,3)：先平移 y，再平移 x
    m = Affine2D.translate(2.0, 0.0).compose(Affine2D.translate(0.0, 3.0))
    assert m.transform_point(0, 0) == (2.0, 3.0)


def test_compose_scale_then_translate():
    # 先 scale(2)，再 translate(1,1)：p' = (2x+1, 2y+1)
    m = Affine2D.translate(1.0, 1.0).compose(Affine2D.scale(2.0, 2.0))
    assert m.transform_point(1, 0) == (3.0, 1.0)


def test_transform_rect_aabb():
    # 旋转 90° 的 1×2 矩形（中心旋转）→ AABB 2×1
    rect = Rect(0, 0, 1, 2)
    m = Affine2D.translate(0.5, 1.0)
    m = m.compose(Affine2D.rotate(90.0))
    m = m.compose(Affine2D.translate(-0.5, -1.0))
    out = m.transform_rect(rect)
    # 1×2 绕中心 (0.5,1) 转 90°：AABB x∈[-0.5,1.5] y∈[0.5,1.5]
    assert out.x == pytest.approx(-0.5, abs=1e-9)
    assert out.y == pytest.approx(0.5, abs=1e-9)
    assert out.width == pytest.approx(2.0, abs=1e-9)
    assert out.height == pytest.approx(1.0, abs=1e-9)


# ---------------------------------------------------------------- Rect


def test_rect_intersection():
    a = Rect(0, 0, 2, 2)
    b = Rect(1, 1, 2, 2)
    inter = rect_intersection(a, b)
    assert inter == Rect(1, 1, 1, 1)


def test_rect_intersection_disjoint():
    a = Rect(0, 0, 1, 1)
    b = Rect(2, 2, 1, 1)
    assert rect_intersection(a, b) is None


def test_rect_intersection_touching_edge_is_empty():
    a = Rect(0, 0, 1, 1)
    b = Rect(1, 0, 1, 1)
    assert rect_intersection(a, b) is None  # 边贴边不算重叠


def test_overlap_area():
    a = Rect(0, 0, 3, 3)
    b = Rect(1, 1, 3, 3)
    assert overlap_area(a, b) == pytest.approx(4.0)
    assert overlap_area(a, Rect(10, 10, 1, 1)) == 0.0


def test_rect_contains():
    outer = Rect(0, 0, 5, 5)
    assert rect_contains(outer, Rect(1, 1, 2, 2))
    assert not rect_contains(outer, Rect(4, 4, 2, 2))


# ---------------------------------------------------------------- 不按零旋转误处理


def test_rotate_not_silently_zero():
    # 45° 的 1×1 矩形 AABB ≠ 原矩形：证明旋转被应用而非当作 0
    m = Affine2D.translate(0.5, 0.5)
    m = m.compose(Affine2D.rotate(45.0))
    m = m.compose(Affine2D.translate(-0.5, -0.5))
    out = m.transform_rect(Rect(0, 0, 1, 1))
    assert out.width > 1.0  # 旋转后包围盒必然变大
