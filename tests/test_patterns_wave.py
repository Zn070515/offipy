"""A4 Task 3 — wave pattern determinism, structure, and bounds."""

from __future__ import annotations

import hashlib
import re
import xml.etree.ElementTree as ET

import pytest

from offipy.assets.patterns import wave
from offipy.assets.patterns._common import BG, FG

_NS = "{http://www.w3.org/2000/svg}"
_NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")


def _paths(svg: str) -> list[ET.Element]:
    return ET.fromstring(svg).findall(f".//{_NS}path")


def _first(svg: str) -> ET.Element:
    return list(ET.fromstring(svg))[0]


def _nums_in(d: str) -> list[float]:
    return [float(m) for m in _NUM_RE.findall(d)]


def _default(seed: int = 0) -> str:
    return wave.build(seed=seed, background="transparent", density=0.5, thickness=0.5)


def test_snapshot_seed0_default() -> None:
    svg = _default()
    digest = hashlib.sha256(svg.encode("utf-8")).hexdigest()
    assert digest == "1952eb18c22840171e254478190c3e58278879c6af71a4145282f3d049b3a317"


def test_seed1_differs() -> None:
    assert _default(0) != _default(1)


def test_same_seed_repeat_identical() -> None:
    assert _default(7) == _default(7)


def test_density_changes_line_count() -> None:
    low = wave.build(seed=0, background="transparent", density=0.0, thickness=0.5)
    high = wave.build(seed=0, background="transparent", density=1.0, thickness=0.5)
    assert len(_paths(low)) == 2
    assert len(_paths(high)) == 10


def test_thickness_changes_stroke_width_only() -> None:
    thin = wave.build(seed=3, background="transparent", density=0.5, thickness=0.0)
    thick = wave.build(seed=3, background="transparent", density=0.5, thickness=1.0)
    thin_d = [p.get("d") for p in _paths(thin)]
    thick_d = [p.get("d") for p in _paths(thick)]
    assert thin_d == thick_d  # placement unchanged
    assert [p.get("stroke-width") for p in _paths(thin)] == ["1.5"] * 6
    assert [p.get("stroke-width") for p in _paths(thick)] == ["11.5"] * 6


def test_foreground_sentinel_on_paths() -> None:
    svg = _default()
    assert FG in svg
    for path in _paths(svg):
        assert path.get("stroke") == FG


def test_background_rect_omitted_when_transparent() -> None:
    svg = _default()
    assert BG not in svg
    assert _first(svg).tag == _NS + "path"


def test_background_rect_first_when_opaque() -> None:
    svg = wave.build(seed=0, background="#123456", density=0.5, thickness=0.5)
    assert f'<rect width="1000" height="1000" fill="{BG}"/>' in svg
    assert _first(svg).tag == _NS + "rect"


@pytest.mark.parametrize(
    ("density", "thickness"), [(d, t) for d in (0.0, 0.5, 1.0) for t in (0.0, 0.5, 1.0)]
)
def test_coords_within_overscan(density: float, thickness: float) -> None:
    for seed in range(0, 40):
        svg = wave.build(seed=seed, background="transparent", density=density, thickness=thickness)
        for path in _paths(svg):
            for n in _nums_in(path.get("d", "")):
                assert -50.0 <= n <= 1050.0


def test_valid_xml() -> None:
    svg = wave.build(seed=5, background="#ABCDEF", density=0.8, thickness=0.2)
    ET.fromstring(svg)  # must not raise


def test_template_size_bounded() -> None:
    svg = wave.build(seed=0, background="transparent", density=1.0, thickness=1.0)
    assert len(svg.encode("utf-8")) < 30_000
