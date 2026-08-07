"""A4 Task 5 — square-grid pattern determinism, rectilinear geometry, bounds."""

from __future__ import annotations

import hashlib
import xml.etree.ElementTree as ET

from offipy.assets.patterns import square_grid
from offipy.assets.patterns._common import BG, FG

_NS = "{http://www.w3.org/2000/svg}"


def _lines(svg: str) -> list[ET.Element]:
    return ET.fromstring(svg).findall(f".//{_NS}line")


def _line_coords(svg: str) -> list[tuple[str, str, str, str]]:
    return [
        (line.get("x1", ""), line.get("y1", ""), line.get("x2", ""), line.get("y2", ""))
        for line in _lines(svg)
    ]


def _default(seed: int = 0) -> str:
    return square_grid.build(seed=seed, background="transparent", spacing=1.0, thickness=0.5)


def test_snapshot_seed0_default() -> None:
    digest = hashlib.sha256(_default().encode("utf-8")).hexdigest()
    assert digest == "e654fa81bf1e8e33e137c601e9717f7b301797a6ce4a4a40efccb624c2013a80"


def test_same_seed_repeat_identical() -> None:
    assert _default() == _default()


def test_seed_offsets_origin() -> None:
    assert _default(0) != _default(1)


def test_spacing_changes_line_count() -> None:
    tight = square_grid.build(seed=0, background="transparent", spacing=0.5, thickness=0.5)
    loose = square_grid.build(seed=0, background="transparent", spacing=2.0, thickness=0.5)
    assert len(_lines(tight)) == 40
    assert len(_lines(loose)) == 14


def test_line_count_bounded() -> None:
    for seed in range(0, 20):
        svg = square_grid.build(seed=seed, background="transparent", spacing=0.5, thickness=0.5)
        assert len(_lines(svg)) < 1000


def test_thickness_changes_stroke_width_only() -> None:
    thin = square_grid.build(seed=3, background="transparent", spacing=1.0, thickness=0.0)
    thick = square_grid.build(seed=3, background="transparent", spacing=1.0, thickness=1.0)
    assert _line_coords(thin) == _line_coords(thick)  # placement unchanged
    assert {line.get("stroke-width") for line in _lines(thin)} == {"1"}
    assert {line.get("stroke-width") for line in _lines(thick)} == {"9"}


def test_no_duplicate_lines() -> None:
    for seed in range(0, 20):
        svg = square_grid.build(seed=seed, background="transparent", spacing=0.5, thickness=0.5)
        coords = _line_coords(svg)
        assert len(set(coords)) == len(coords)


def test_geometry_is_rectilinear() -> None:
    svg = _default()
    for line in _lines(svg):
        assert line.get("x1") == line.get("x2") or line.get("y1") == line.get("y2")


def test_lines_use_foreground_sentinel() -> None:
    svg = _default()
    assert FG in svg
    for line in _lines(svg):
        assert line.get("stroke") == FG


def test_background_rect_first_when_opaque() -> None:
    svg = square_grid.build(seed=0, background="#ABCDEF", spacing=1.0, thickness=0.5)
    assert BG in svg
    assert list(ET.fromstring(svg))[0].tag == _NS + "rect"


def test_template_size_bounded() -> None:
    svg = square_grid.build(seed=0, background="transparent", spacing=0.5, thickness=1.0)
    assert len(svg.encode("utf-8")) < 80_000
