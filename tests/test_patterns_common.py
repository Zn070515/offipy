"""A4 Task 1 — deterministic procedural SVG foundation.

Covers the SplitMix32 PRNG fixed sequences, the stable float formatter, the
deterministic SVG root tag, and cross-process determinism (the CI matrix runs
this suite on Python 3.10–3.13, so the subprocess test pins byte identity
across those versions).
"""

from __future__ import annotations

import os
import subprocess
import sys

import pytest

from offipy.assets.patterns import _common
from offipy.exceptions import InvalidArgumentError

_INT32_MIN = -(1 << 31)
_INT32_MAX = (1 << 31) - 1

# (seed, first 8 u32 outputs) — pinned so any PRNG change breaks loudly.
_SEQ_CASES: tuple[tuple[int, list[int]], ...] = (
    (
        0,
        [
            2462723854,
            1020716019,
            454327756,
            1275600319,
            1215922603,
            3678440605,
            2025593743,
            3627053797,
        ],
    ),
    (
        1,
        [
            2527132011,
            314344336,
            2535364964,
            2041432039,
            1495043544,
            3445983177,
            4176287394,
            1522731872,
        ],
    ),
    (
        -1,
        [
            920564995,
            4230986166,
            697614773,
            1778835764,
            280495159,
            1500331647,
            1240119404,
            1650816001,
        ],
    ),
    (
        _INT32_MIN,
        [
            3172115818,
            304882621,
            205969554,
            3182573129,
            2043043288,
            3787605635,
            1301340159,
            2358788073,
        ],
    ),
    (
        _INT32_MAX,
        [
            849629901,
            219964201,
            1510287064,
            2144703427,
            3612074208,
            1332670368,
            2209325893,
            3014266354,
        ],
    ),
)


@pytest.mark.parametrize(("seed", "expected"), _SEQ_CASES)
def test_rng_fixed_sequences(seed: int, expected: list[int]) -> None:
    rng = _common._Rng(seed)
    assert [rng.u32() for _ in range(8)] == expected


def test_rng_seed_normalized_to_32bit() -> None:
    # signed seed masks to the same 32-bit state as its two's-complement twin
    a = _common._Rng(-1)
    b = _common._Rng(0xFFFFFFFF)
    assert [a.u32() for _ in range(4)] == [b.u32() for _ in range(4)]


def test_rng_unit_range() -> None:
    rng = _common._Rng(7)
    for _ in range(100):
        u = rng.unit()
        assert 0.0 <= u < 1.0


def test_rng_uniform_range() -> None:
    rng = _common._Rng(7)
    for _ in range(100):
        u = rng.uniform(-10.0, 30.0)
        assert -10.0 <= u < 30.0


def test_rng_choice_index_range() -> None:
    rng = _common._Rng(7)
    for _ in range(100):
        assert 0 <= rng.choice_index(12) < 12


def test_rng_choice_index_rejects_nonpositive() -> None:
    with pytest.raises(InvalidArgumentError):
        _common._Rng(1).choice_index(0)
    with pytest.raises(InvalidArgumentError):
        _common._Rng(1).choice_index(-3)


def test_rng_deterministic_across_instances() -> None:
    a = _common._Rng(12345)
    b = _common._Rng(12345)
    seq_a = [a.u32() for _ in range(100)]
    seq_b = [b.u32() for _ in range(100)]
    assert seq_a == seq_b


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (0.0, "0"),
        (-0.0, "0"),
        (5.0, "5"),
        (5.5, "5.5"),
        (0.1, "0.1"),
        (123.4567, "123.457"),
        (-0.0004, "0"),
        (1e-7, "0"),
        (0.0005, "0.001"),
        (42.0, "42"),
        (-3.3333, "-3.333"),
        (1000.0001, "1000"),
    ],
)
def test_fmt_table(value: float, expected: str) -> None:
    assert _common._fmt(value) == expected


@pytest.mark.parametrize(
    ("value", "digits", "expected"),
    [
        (1.23456, 4, "1.2346"),
        (0.123456, 4, "0.1235"),
        (2.0, 4, "2"),
        (0.0, 4, "0"),
    ],
)
def test_fmt_digits_param(value: float, digits: int, expected: str) -> None:
    assert _common._fmt(value, digits=digits) == expected


@pytest.mark.parametrize("bad", [float("nan"), float("inf"), float("-inf")])
def test_fmt_rejects_non_finite(bad: float) -> None:
    with pytest.raises(InvalidArgumentError):
        _common._fmt(bad)


def test_svg_open_root_bytes() -> None:
    assert _common.svg_open() == '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000">'


def test_svg_open_extra_attrs_order_fixed() -> None:
    got = _common.svg_open((("role", "img"), ("data-x", "1")))
    assert got == (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" role="img" data-x="1">'
    )


def test_sentinel_constants() -> None:
    assert _common.FG == "__OFFIPY_ASSET_FG__"
    assert _common.BG == "__OFFIPY_ASSET_BG__"


def test_subprocess_cross_interpreter_determinism() -> None:
    """Fresh interpreter must reproduce the in-process sequence (pins byte
    identity across the CI matrix's Python versions)."""
    script = (
        "from offipy.assets.patterns._common import _Rng; print(_Rng(42).u32(), _Rng(-7).u32())"
    )
    proc = subprocess.run(
        [sys.executable, "-c", script],
        check=True,
        capture_output=True,
        text=True,
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
    )
    in_process = [_common._Rng(42).u32(), _common._Rng(-7).u32()]
    assert proc.stdout.split() == [str(v) for v in in_process]
