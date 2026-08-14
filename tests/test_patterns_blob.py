"""A4 Task 4 — blob pattern: closed single path, complexity mapping, bounds."""

from __future__ import annotations

import hashlib
import re
import xml.etree.ElementTree as ET

import pytest

from offipy.assets.patterns import blob
from offipy.assets.patterns._common import BG, FG

_NS = "{http://www.w3.org/2000/svg}"
_NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")


def _paths(svg: str) -> list[ET.Element]:
    return ET.fromstring(svg).findall(f".//{_NS}path")


def _d(svg: str) -> str:
    paths = _paths(svg)
    assert len(paths) == 1
    return paths[0].get("d", "")


def test_snapshot_seed0_default() -> None:
    svg = blob.build(seed=0, background="transparent", complexity=5)
    digest = hashlib.sha256(svg.encode("utf-8")).hexdigest()
    assert digest == "c30867c943e18a1939e4413edadf483c1cc537c828cfdd3e3c3f4a558567966c"


def test_complexity_changes_anchor_count() -> None:
    low = _d(blob.build(seed=0, background="transparent", complexity=2))
    high = _d(blob.build(seed=0, background="transparent", complexity=8))
    assert low.count("C") == 8
    assert high.count("C") == 16


def test_single_closed_path() -> None:
    svg = blob.build(seed=0, background="transparent", complexity=5)
    assert len(_paths(svg)) == 1
    nums = [float(m) for m in _NUM_RE.findall(_d(svg))]
    assert (nums[0], nums[1]) == (nums[-2], nums[-1])  # closed back to start


def test_same_seed_repeat_identical() -> None:
    assert blob.build(seed=9, background="transparent", complexity=7) == blob.build(
        seed=9, background="transparent", complexity=7
    )


def test_seed_changes_shape() -> None:
    a = blob.build(seed=0, background="transparent", complexity=5)
    b = blob.build(seed=1, background="transparent", complexity=5)
    assert a != b


def test_foreground_sentinel_fill() -> None:
    svg = blob.build(seed=0, background="transparent", complexity=5)
    assert FG in svg
    assert _paths(svg)[0].get("fill") == FG


def test_background_rect_first_when_opaque() -> None:
    svg = blob.build(seed=0, background="#123456", complexity=5)
    assert BG in svg
    assert next(iter(ET.fromstring(svg))).tag == _NS + "rect"


def test_background_omitted_when_transparent() -> None:
    svg = blob.build(seed=0, background="transparent", complexity=5)
    assert BG not in svg


@pytest.mark.parametrize("complexity", [2, 5, 8])
def test_coords_within_overscan(complexity: int) -> None:
    for seed in range(40):
        svg = blob.build(seed=seed, background="transparent", complexity=complexity)
        for n in _NUM_RE.findall(_d(svg)):
            assert -100.0 <= float(n) <= 1100.0


def test_valid_xml() -> None:
    ET.fromstring(blob.build(seed=3, background="#ABCDEF", complexity=8))


def test_template_size_bounded() -> None:
    svg = blob.build(seed=0, background="transparent", complexity=8)
    assert len(svg.encode("utf-8")) < 20_000
