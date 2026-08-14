"""A4 Task 8 — circuit pattern: unique nodes, orthogonal routes, bounded edges."""

from __future__ import annotations

import hashlib
import re
import xml.etree.ElementTree as ET

import pytest

from offipy.assets.patterns import circuit
from offipy.assets.patterns._common import BG, FG

_NS = "{http://www.w3.org/2000/svg}"

_NUM = r"[0-9.eE+-]+"


def _nodes(svg: str) -> list[tuple[str, str]]:
    return [
        (c.get("cx", ""), c.get("cy", "")) for c in ET.fromstring(svg).findall(f".//{_NS}circle")
    ]


def _paths(svg: str) -> list[str]:
    return [p.get("d", "") for p in ET.fromstring(svg).findall(f".//{_NS}path")]


def _default(seed: int = 0) -> str:
    return circuit.build(seed=seed, background="transparent", nodes=20, density=0.5)


def test_snapshot_seed0_default() -> None:
    digest = hashlib.sha256(_default().encode("utf-8")).hexdigest()
    assert digest == "35b436e355865aed477a286b50089a299d8cbedc9718e37f99bc6776c92b54e7"


def test_snapshot_seed7_default() -> None:
    digest = hashlib.sha256(_default(seed=7).encode("utf-8")).hexdigest()
    assert digest == "7ace0d34fdcf28043a5c30a04207a0f9d954a25e65ac7dccceb8e06fb3b56415"


@pytest.mark.parametrize("nodes", [4, 20, 60])
def test_exact_unique_node_count(nodes: int) -> None:
    svg = circuit.build(seed=0, background="transparent", nodes=nodes, density=0.5)
    coords = _nodes(svg)
    assert len(coords) == nodes
    assert len(set(coords)) == nodes
    ET.fromstring(svg)


def test_no_duplicate_nodes_across_seeds() -> None:
    for seed in range(40):
        coords = _nodes(circuit.build(seed=seed, background="transparent", nodes=60, density=1.0))
        assert len(set(coords)) == len(coords)


def test_all_segments_axis_aligned() -> None:
    for d in _paths(_default()):
        cmds = re.findall(f"[MHV]{_NUM}", d)
        assert cmds, d
        for tok in cmds:
            assert tok[0] in ("M", "H", "V"), (tok, d)


def test_density_scales_edge_count() -> None:
    def edges(density: float) -> int:
        return len(
            _paths(circuit.build(seed=0, background="transparent", nodes=60, density=density))
        )

    assert edges(0.0) == 59  # minimum spanning set: nodes - 1
    assert edges(1.0) > edges(0.0)
    assert edges(1.0) <= 2 * 60


def test_same_seed_exact_bytes() -> None:
    assert _default(seed=7) == _default(seed=7)


def test_seed_changes_geometry() -> None:
    assert _default(seed=0) != _default(seed=1)


def test_paths_and_nodes_use_foreground_sentinel() -> None:
    svg = _default()
    root = ET.fromstring(svg)
    assert FG in svg
    for p in root.findall(f".//{_NS}path"):
        assert p.get("stroke") == FG
        assert p.get("fill") == "none"
    for c in root.findall(f".//{_NS}circle"):
        assert c.get("fill") == FG


def test_background_rect_first_when_opaque() -> None:
    svg = circuit.build(seed=0, background="#123456", nodes=20, density=0.5)
    assert BG in svg
    assert next(iter(ET.fromstring(svg))).tag == _NS + "rect"


def test_template_size_bounded() -> None:
    svg = circuit.build(seed=0, background="transparent", nodes=60, density=1.0)
    assert len(svg.encode("utf-8")) < 180_000
