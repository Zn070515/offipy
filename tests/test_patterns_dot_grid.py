"""A4 Task 5 — dot-grid pattern determinism, bounds, and zero-radius behavior."""

from __future__ import annotations

import hashlib
import xml.etree.ElementTree as ET

from offipy.assets.patterns import dot_grid
from offipy.assets.patterns._common import BG, FG

_NS = "{http://www.w3.org/2000/svg}"


def _circles(svg: str) -> list[ET.Element]:
    return ET.fromstring(svg).findall(f".//{_NS}circle")


def _default(seed: int = 0) -> str:
    return dot_grid.build(seed=seed, background="transparent", spacing=1.0, radius=0.15)


def test_snapshot_seed0_default() -> None:
    digest = hashlib.sha256(_default().encode("utf-8")).hexdigest()
    assert digest == "53cce071af4225de7e4189a88f2aca677e2e668c7609a629db3b4280ec90ab97"


def test_same_seed_repeat_identical() -> None:
    assert _default() == _default()


def test_seed_offsets_origin() -> None:
    assert _default(0) != _default(1)


def test_spacing_changes_cell_density() -> None:
    tight = dot_grid.build(seed=0, background="transparent", spacing=0.5, radius=0.15)
    loose = dot_grid.build(seed=0, background="transparent", spacing=2.0, radius=0.15)
    assert len(_circles(tight)) == 400
    assert len(_circles(loose)) == 49


def test_element_count_bounded() -> None:
    for seed in range(0, 20):
        svg = dot_grid.build(seed=seed, background="transparent", spacing=0.5, radius=0.15)
        assert len(_circles(svg)) < 1000


def test_radius_zero_empty_foreground() -> None:
    svg = dot_grid.build(seed=0, background="transparent", spacing=1.0, radius=0.0)
    assert _circles(svg) == []
    assert FG not in svg
    ET.fromstring(svg)  # still valid XML


def test_radius_zero_with_opaque_background_keeps_rect() -> None:
    svg = dot_grid.build(seed=0, background="#123456", spacing=1.0, radius=0.0)
    assert list(ET.fromstring(svg))[0].tag == _NS + "rect"


def test_dots_are_distinct() -> None:
    svg = _default()
    pairs = [(c.get("cx"), c.get("cy")) for c in _circles(svg)]
    assert len(set(pairs)) == len(pairs)


def test_dots_use_foreground_sentinel() -> None:
    svg = _default()
    assert FG in svg
    for c in _circles(svg):
        assert c.get("fill") == FG


def test_background_rect_first_when_opaque() -> None:
    svg = dot_grid.build(seed=0, background="#ABCDEF", spacing=1.0, radius=0.15)
    assert BG in svg
    assert list(ET.fromstring(svg))[0].tag == _NS + "rect"


def test_template_size_bounded() -> None:
    svg = dot_grid.build(seed=0, background="transparent", spacing=0.5, radius=0.5)
    assert len(svg.encode("utf-8")) < 80_000
