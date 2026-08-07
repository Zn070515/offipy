"""A4 Task 6 — rings pattern: exact count, center variance, bounded extent."""

from __future__ import annotations

import hashlib
import xml.etree.ElementTree as ET

import pytest

from offipy.assets.patterns import rings
from offipy.assets.patterns._common import BG, FG

_NS = "{http://www.w3.org/2000/svg}"


def _circles(svg: str) -> list[ET.Element]:
    return ET.fromstring(svg).findall(f".//{_NS}circle")


def _default(seed: int = 0, count: int = 5) -> str:
    return rings.build(seed=seed, background="transparent", count=count, thickness=0.5)


def test_snapshot_seed0_default() -> None:
    digest = hashlib.sha256(_default().encode("utf-8")).hexdigest()
    assert digest == "7f07a2421b24babb6469539dcf21ab190b9d5539c766862c28cb433470675f47"


def test_exact_ring_count() -> None:
    assert len(_circles(_default(count=5))) == 5


@pytest.mark.parametrize("count", [1, 12])
def test_count_extremes_valid(count: int) -> None:
    svg = _default(count=count)
    assert len(_circles(svg)) == count
    ET.fromstring(svg)


def test_same_seed_repeat_identical() -> None:
    assert _default(seed=7, count=8) == _default(seed=7, count=8)


def test_seed_changes_center_not_count() -> None:
    a = _circles(_default(seed=0))
    b = _circles(_default(seed=1))
    assert len(a) == len(b)
    assert a[0].get("cx") != b[0].get("cx")
    assert a[0].get("cy") != b[0].get("cy")


def test_thickness_changes_stroke_width_only() -> None:
    thin = _circles(rings.build(seed=3, background="transparent", count=5, thickness=0.0))
    thick = _circles(rings.build(seed=3, background="transparent", count=5, thickness=1.0))
    assert [c.get("r") for c in thin] == [c.get("r") for c in thick]
    assert {c.get("stroke-width") for c in thin} == {"1"}
    assert {c.get("stroke-width") for c in thick} == {"11"}


def test_rings_use_foreground_sentinel() -> None:
    svg = _default()
    assert FG in svg
    for c in _circles(svg):
        assert c.get("stroke") == FG
        assert c.get("fill") == "none"


def test_background_rect_first_when_opaque() -> None:
    svg = rings.build(seed=0, background="#123456", count=5, thickness=0.5)
    assert BG in svg
    assert list(ET.fromstring(svg))[0].tag == _NS + "rect"


@pytest.mark.parametrize("count", [1, 6, 12])
def test_extent_within_overscan(count: int) -> None:
    for seed in range(0, 40):
        svg = rings.build(seed=seed, background="transparent", count=count, thickness=1.0)
        for c in _circles(svg):
            cx, cy, r = float(c.get("cx", "0")), float(c.get("cy", "0")), float(c.get("r", "0"))
            for v in (cx - r, cx + r, cy - r, cy + r):
                assert -500.0 <= v <= 1500.0


def test_template_size_bounded() -> None:
    svg = rings.build(seed=0, background="transparent", count=12, thickness=1.0)
    assert len(svg.encode("utf-8")) < 25_000
