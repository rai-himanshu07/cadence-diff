"""Correctness tests for the bounded, non-production BIFF12 Ptg token-walk
spike (tests/biff12_ptg_spike.py). Never imported by production code.
"""

from __future__ import annotations

import struct

from tests.biff12_ptg_spike import (
    TokenWalkResult,
    locate_rgce,
    walk_ptg_tokens,
)


def _int_token(value: int) -> bytes:
    return bytes([0x1E]) + struct.pack("<H", value)


def _num_token(value: float) -> bytes:
    return bytes([0x1F]) + struct.pack("<d", value)


def test_walks_a_simple_addition_fully_classified() -> None:
    # =1+2 : PtgInt(1) PtgInt(2) PtgAdd
    rgce = _int_token(1) + _int_token(2) + bytes([0x03])

    result = walk_ptg_tokens(rgce)

    assert result.fully_classified
    assert result.stopped_at == len(rgce)
    assert result.stopped_opcode is None
    assert result.opcode_counts == {"PtgInt": 2, "PtgAdd": 1}


def test_walks_every_confidently_sized_operator() -> None:
    operators = {
        0x03: "PtgAdd", 0x04: "PtgSub", 0x05: "PtgMul", 0x06: "PtgDiv",
        0x07: "PtgPower", 0x08: "PtgConcat", 0x09: "PtgLt", 0x0A: "PtgLe",
        0x0B: "PtgEq", 0x0C: "PtgGe", 0x0D: "PtgGt", 0x0E: "PtgNe",
        0x0F: "PtgIsect", 0x10: "PtgUnion", 0x11: "PtgRange",
        0x12: "PtgUplus", 0x13: "PtgUminus", 0x14: "PtgPercent",
        0x15: "PtgParen", 0x16: "PtgMissArg",
    }
    for opcode, name in operators.items():
        result = walk_ptg_tokens(bytes([opcode]))
        assert result.fully_classified, name
        assert result.opcode_counts == {name: 1}


def test_walks_a_ref_and_area_token() -> None:
    # PtgRef family (opcode 0x24: base 0x04 | class bits 0x20): row(4)+col(2).
    # The Ref/Value/Array class LABEL is not verified in this spike (only
    # that class-bit variants of the same base id are recognized and sized
    # identically) -- what matters for an inventory is the base token kind.
    ref = bytes([0x24]) + struct.pack("<IH", 0, 0)
    # PtgArea family (opcode 0x45: base 0x05 | class bits 0x40): 2x(row4+col2).
    area = bytes([0x45]) + struct.pack("<IHIH", 0, 0, 9, 3)

    ref_result = walk_ptg_tokens(ref)
    area_result = walk_ptg_tokens(area)

    assert ref_result.fully_classified
    assert sum(ref_result.opcode_counts.values()) == 1
    assert next(iter(ref_result.opcode_counts)).startswith("PtgRef:")
    assert area_result.fully_classified
    assert sum(area_result.opcode_counts.values()) == 1
    assert next(iter(area_result.opcode_counts)).startswith("PtgArea:")


def test_stops_at_unrecognized_opcode_without_guessing() -> None:
    # A known PtgAdd, then a byte with no table entry (0xF0 is unused).
    rgce = bytes([0x03, 0xF0, 0x03])

    result = walk_ptg_tokens(rgce)

    assert not result.fully_classified
    assert result.stopped_at == 1
    assert result.stopped_opcode == 0xF0
    assert result.opcode_counts == {"PtgAdd": 1}  # only what was seen BEFORE the stop


def test_stops_on_unsized_string_token_without_guessing() -> None:
    # PtgStr has no confidently-known size in this spike -- must stop, not
    # skip an arbitrary number of bytes.
    rgce = bytes([0x17]) + b"whatever-follows-is-unknown"

    result = walk_ptg_tokens(rgce)

    assert not result.fully_classified
    assert result.stopped_at == 0
    assert result.stopped_opcode == 0x17
    assert result.opcode_counts == {}


def test_stops_on_truncated_fixed_size_operand() -> None:
    # PtgNum declares an 8-byte operand but only 3 bytes remain.
    rgce = bytes([0x1F, 0x01, 0x02, 0x03])

    result = walk_ptg_tokens(rgce)

    assert not result.fully_classified
    assert result.stopped_at == 0
    assert result.stopped_opcode == 0x1F


def test_never_raises_on_arbitrary_bytes() -> None:
    """Fuzz-safety: no input can crash the walker -- it only ever stops."""
    import random

    rng = random.Random(20260904)
    for _ in range(500):
        length = rng.randint(0, 64)
        garbage = bytes(rng.randrange(256) for _ in range(length))
        result = walk_ptg_tokens(garbage)
        assert isinstance(result, TokenWalkResult)
        assert result.stopped_at <= len(garbage)


def test_locate_rgce_self_validates_against_declared_cce() -> None:
    """A synthetic BrtFmlaNum-shaped payload with a correct cce is accepted;
    tampering the cce so the arithmetic no longer holds is rejected."""
    rgce = _int_token(1) + _int_token(2) + bytes([0x03])
    col, xf, cached = 0, 0, 3.0
    good_payload = (
        struct.pack("<IId", col, xf, cached)
        + struct.pack("<H", 0)  # 2-byte flags candidate
        + struct.pack("<I", len(rgce))
        + rgce
    )

    found = locate_rgce(0x0009, good_payload)

    assert found == rgce

    bad_payload = (
        struct.pack("<IId", col, xf, cached)
        + struct.pack("<H", 0)
        + struct.pack("<I", len(rgce) + 5)  # cce lies about the length
        + rgce
    )
    assert locate_rgce(0x0009, bad_payload) is None


def test_locate_rgce_returns_none_for_an_unmapped_record_id() -> None:
    assert locate_rgce(0x9999, b"\x00" * 32) is None
