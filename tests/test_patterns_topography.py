"""A4 Task 7 — topography pattern: contour lines, density metrics, bounds."""

from __future__ import annotations

import hashlib
import xml.etree.ElementTree as ET

import pytest

from offipy.assets.patterns import topography
from offipy.assets.patterns._common import BG, FG

_NS = "{http://www.w3.org/2000/svg}"


def _paths(svg: str) -> list[ET.Element]:
    return ET.fromstring(svg).findall(f".//{_NS}path")


def _path_ys(d: str) -> list[float]:
    nums = [float(t) for t in d.replace("M", " ").replace("C", " ").split()]
    return [nums[i + 1] for i in range(0, len(nums) - 1, 2)]


def _all_points(svg: str) -> list[tuple[float, float]]:
    pts: list[tuple[float, float]] = []
    for p in _paths(svg):
        nums = [float(t) for t in p.get("d", "").replace("M", " ").replace("C", " ").split()]
        pts.extend((nums[i], nums[i + 1]) for i in range(0, len(nums) - 1, 2))
    return pts


def _default(seed: int = 0) -> str:
    return topography.build(seed=seed, background="transparent", density=0.5, lines=10)


def test_snapshot_seed0_default() -> None:
    digest = hashlib.sha256(_default().encode("utf-8")).hexdigest()
    assert digest == "ede3ff8967cf1a7dfdcffa13c21ad3c67615747810fe2c105adf7b3039fce6a4"


def test_exact_line_count() -> None:
    assert len(_paths(_default())) == 10


@pytest.mark.parametrize("lines", [3, 24])
def test_line_count_extremes_valid(lines: int) -> None:
    svg = topography.build(seed=0, background="transparent", density=1.0, lines=lines)
    assert len(_paths(svg)) == lines
    ET.fromstring(svg)


def test_same_seed_repeat_identical() -> None:
    assert _default(seed=7) == _default(seed=7)


def test_seed_changes_geometry() -> None:
    assert _default(seed=0) != _default(seed=1)


def test_density_scales_sample_density() -> None:
    light = topography.build(seed=0, background="transparent", density=0.0, lines=10)
    rich = topography.build(seed=0, background="transparent", density=1.0, lines=10)
    assert _paths(light)[0].get("d", "").count("C") == 40
    assert _paths(rich)[0].get("d", "").count("C") == 80


def test_density_scales_amplitude() -> None:
    light = topography.build(seed=0, background="transparent", density=0.0, lines=10)
    rich = topography.build(seed=0, background="transparent", density=1.0, lines=10)
    span_light = max(
        max(_path_ys(p.get("d", ""))) - min(_path_ys(p.get("d", ""))) for p in _paths(light)
    )
    span_rich = max(
        max(_path_ys(p.get("d", ""))) - min(_path_ys(p.get("d", ""))) for p in _paths(rich)
    )
    assert span_rich > span_light


def test_density_scales_roughness() -> None:
    light = topography.build(seed=0, background="transparent", density=0.0, lines=10)
    rich = topography.build(seed=0, background="transparent", density=1.0, lines=10)

    def max_flips(svg: str) -> int:
        best = 0
        for p in _paths(svg):
            ys = _path_ys(p.get("d", ""))
            flips = sum(
                1 for i in range(1, len(ys) - 1) if (ys[i] - ys[i - 1]) * (ys[i + 1] - ys[i]) < 0
            )
            best = max(best, flips)
        return best

    assert max_flips(rich) > max_flips(light)


def test_every_path_starts_at_left_edge() -> None:
    for p in _paths(_default()):
        nums = [float(t) for t in p.get("d", "").replace("M", " ").replace("C", " ").split()]
        assert nums[0] == 0.0


def test_paths_use_foreground_sentinel() -> None:
    svg = _default()
    assert FG in svg
    for p in _paths(svg):
        assert p.get("stroke") == FG
        assert p.get("fill") == "none"


def test_background_rect_first_when_opaque() -> None:
    svg = topography.build(seed=0, background="#123456", density=0.5, lines=10)
    assert BG in svg
    assert list(ET.fromstring(svg))[0].tag == _NS + "rect"


@pytest.mark.parametrize("density", [0.0, 0.5, 1.0])
@pytest.mark.parametrize("lines", [3, 10, 24])
def test_extent_within_overscan(density: float, lines: int) -> None:
    for seed in range(0, 40):
        svg = topography.build(seed=seed, background="transparent", density=density, lines=lines)
        for x, y in _all_points(svg):
            assert -100.0 <= x <= 1100.0
            assert -100.0 <= y <= 1100.0


def test_template_size_bounded() -> None:
    svg = topography.build(seed=0, background="transparent", density=1.0, lines=24)
    assert len(svg.encode("utf-8")) < 150_000


def test_many_max_calls_stable() -> None:
    expected = topography.build(seed=3, background="transparent", density=1.0, lines=24)
    for _ in range(100):
        assert topography.build(seed=3, background="transparent", density=1.0, lines=24) == expected
