"""A4 Task 9 — gradient-orb pattern: radial gradients, blur stops, no filter."""

from __future__ import annotations

import hashlib
import re
import xml.etree.ElementTree as ET

import pytest

from offipy.assets.patterns import gradient_orb
from offipy.assets.patterns._common import BG, FG

_NS = "{http://www.w3.org/2000/svg}"

_NUM = r"[0-9.eE+-]+"


def _orbs(svg: str) -> list[tuple[float, float, float]]:
    return [
        (float(c.get("cx", "0")), float(c.get("cy", "0")), float(c.get("r", "0")))
        for c in ET.fromstring(svg).findall(f".//{_NS}circle")
    ]


def _default(seed: int = 0) -> str:
    return gradient_orb.build(seed=seed, background="transparent", orb_count=3, blur=0.5)


def test_snapshot_seed0_default() -> None:
    digest = hashlib.sha256(_default().encode("utf-8")).hexdigest()
    assert digest == "5f0142c29568192c72c5e89006bce6a88f235ef5dd3f2b255c0454dde49655a1"


def test_snapshot_seed7_default() -> None:
    digest = hashlib.sha256(_default(seed=7).encode("utf-8")).hexdigest()
    assert digest == "d505c57cac2edc6a906f5b8a25abc7be5e4f93e5ac6c87de6f27b0a7f779974d"


@pytest.mark.parametrize("orb_count", [1, 3, 6])
def test_exact_orb_count(orb_count: int) -> None:
    svg = gradient_orb.build(seed=0, background="transparent", orb_count=orb_count, blur=0.5)
    assert len(_orbs(svg)) == orb_count
    assert len(re.findall(r"<radialGradient", svg)) == orb_count
    ET.fromstring(svg)


def test_no_filter_element() -> None:
    svg = _default()
    assert "<filter" not in svg
    assert "feGaussianBlur" not in svg


def test_gradient_ids_deterministic_and_unique() -> None:
    svg = _default()
    ids = re.findall(r'id="(orb-\d+)"', svg)
    assert ids == ["orb-0", "orb-1", "orb-2"]
    assert len(set(ids)) == len(ids)


def test_fill_refs_match_defined_gradients() -> None:
    svg = _default()
    ids = set(re.findall(r'id="(orb-\d+)"', svg))
    refs = set(re.findall(r'fill="url\(#(orb-\d+)\)"', svg))
    assert refs == ids


def test_blur_spreads_gradient_stops() -> None:
    hard = gradient_orb.build(seed=0, background="transparent", orb_count=3, blur=0.0)
    soft = gradient_orb.build(seed=0, background="transparent", orb_count=3, blur=1.0)
    stops_hard = re.findall(r'<stop offset="([^"]+)"', hard)
    stops_soft = re.findall(r'<stop offset="([^"]+)"', soft)
    assert "85%" in stops_hard and "25%" in stops_soft


def test_same_seed_repeat_identical() -> None:
    assert _default(seed=7) == _default(seed=7)


def test_seed_changes_geometry() -> None:
    assert _default(seed=0) != _default(seed=1)


def test_gradients_use_foreground_sentinel() -> None:
    svg = _default()
    assert FG in svg
    assert svg.count(f'stop-color="{FG}"') == 9  # 3 orbs x 3 stops


def test_background_rect_first_when_opaque() -> None:
    svg = gradient_orb.build(seed=0, background="#123456", orb_count=3, blur=0.5)
    assert BG in svg
    assert next(iter(ET.fromstring(svg))).tag == _NS + "rect"


@pytest.mark.parametrize("orb_count", [1, 3, 6])
def test_extent_within_overscan(orb_count: int) -> None:
    for seed in range(40):
        svg = gradient_orb.build(seed=seed, background="transparent", orb_count=orb_count, blur=1.0)
        for cx, cy, radius in _orbs(svg):
            for v in (cx - radius, cx + radius, cy - radius, cy + radius):
                assert -100.0 <= v <= 1100.0


def test_template_size_bounded() -> None:
    svg = gradient_orb.build(seed=0, background="transparent", orb_count=6, blur=1.0)
    assert len(svg.encode("utf-8")) < 50_000
