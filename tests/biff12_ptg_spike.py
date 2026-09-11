"""Non-production research spike: BIFF12 Ptg formula-token inventory.

plan-20260904-large-workbook-load-and-formula-compare.md, Step 5. Scope, deliberately
bounded and disclosed up front:

- This module NEVER ships in the production loader (`qc_tool/`) and is
  imported by no shipped code.
- It inventories Ptg opcode KINDS and COUNTS from raw BIFF12 formula-token
  streams (``rgce``) -- never formula text, never cell coordinates, never
  sheet names.
- Cell/area reference tokens (PtgRef/PtgArea/PtgRefN/PtgAreaN/3d variants)
  are recognized and sized (their byte layout -- 4-byte row + 2-byte column,
  the documented BIFF12 "wide" cell-reference convention -- is the most
  stable, well-established part of the format), but this spike does NOT
  render A1 formula text and does NOT attempt to interpret the
  relative/absolute reference flag bits. That interpretation, PtgExp/shared-
  formula follower reconstruction, string/array/name literal decoding, and
  the variable-length PtgAttr "choose" jump table are all explicitly
  UNVERIFIED here and are treated as bounds-safe "unknown, stop" tokens --
  never guessed. Getting a fixed-size guess wrong would misalign every
  subsequent token in the same formula, silently corrupting the inventory;
  refusing to guess keeps this module honest and safe on arbitrary input.
- Every parse is self-validating: a formula's declared ``cce`` byte count
  must be consumed EXACTLY (or the walk stops at a genuinely unrecognized
  opcode) before its tokens count as "fully classified" evidence.
- The formula-record payload prefix (before ``cce``/``rgce``) is confirmed
  per record type by exactly this cce-consistency check, not asserted from
  an external spec citation.

Findings from this module feed a go/no-go report; per the plan, a "no-go"
verdict is a legitimate, complete spike outcome.
"""

from __future__ import annotations

import struct
from collections import Counter
from dataclasses import dataclass, field

# --- Ptg opcode table --------------------------------------------------------

#: Unclassified (control/operator/constant) tokens: opcode < 0x20.
#: name -> fixed operand byte count after the opcode byte, or None if this
#: spike does not have a confidently-sized layout for it.
_UNCLASSIFIED: dict[int, tuple[str, int | None]] = {
    0x01: ("PtgExp", 6),  # BIFF12 sizing: anchor row(4) + col(2)
    0x02: ("PtgTbl", 6),
    0x03: ("PtgAdd", 0),
    0x04: ("PtgSub", 0),
    0x05: ("PtgMul", 0),
    0x06: ("PtgDiv", 0),
    0x07: ("PtgPower", 0),
    0x08: ("PtgConcat", 0),
    0x09: ("PtgLt", 0),
    0x0A: ("PtgLe", 0),
    0x0B: ("PtgEq", 0),
    0x0C: ("PtgGe", 0),
    0x0D: ("PtgGt", 0),
    0x0E: ("PtgNe", 0),
    0x0F: ("PtgIsect", 0),
    0x10: ("PtgUnion", 0),
    0x11: ("PtgRange", 0),
    0x12: ("PtgUplus", 0),
    0x13: ("PtgUminus", 0),
    0x14: ("PtgPercent", 0),
    0x15: ("PtgParen", 0),
    0x16: ("PtgMissArg", 0),
    0x17: ("PtgStr", None),  # variable string encoding -- unverified
    0x19: ("PtgAttr", None),  # variable "choose" jump table -- unverified
    0x1C: ("PtgErr", 1),
    0x1D: ("PtgBool", 1),
    0x1E: ("PtgInt", 2),
    0x1F: ("PtgNum", 8),
}

#: Classified (Ref=0x00/Value=0x20/Array=0x40 class bits) tokens: opcode
#: masked with 0x1F gives this base id. name -> fixed operand byte count,
#: or None if unverified.
_CLASSIFIED_BASE: dict[int, tuple[str, int | None]] = {
    0x00: ("PtgArray", None),  # data lives in trailing rgcb -- unverified
    0x01: ("PtgFunc", 2),
    0x02: ("PtgFuncVar", 3),
    0x03: ("PtgName", None),  # name-index layout unverified
    0x04: ("PtgRef", 6),
    0x05: ("PtgArea", 12),
    0x06: ("PtgMemArea", None),
    0x07: ("PtgMemErr", None),
    0x08: ("PtgMemNoMem", None),
    0x09: ("PtgMemFunc", None),
    0x0A: ("PtgRefErr", 6),
    0x0B: ("PtgAreaErr", 12),
    0x0C: ("PtgRefN", 6),
    0x0D: ("PtgAreaN", 12),
    0x0E: ("PtgMemAreaN", None),
    0x0F: ("PtgMemNoMemN", None),
    0x19: ("PtgNameX", None),  # external-name layout unverified
    0x1A: ("PtgRef3d", 8),
    0x1B: ("PtgArea3d", 14),
    0x1C: ("PtgRefErr3d", 8),
    0x1D: ("PtgAreaErr3d", 14),
}

_CLASS_SUFFIX = {0x00: ":Ref", 0x20: ":Value", 0x40: ":Array"}


def _opcode_lookup(opcode: int) -> tuple[str, int | None] | None:
    if opcode < 0x20:
        return _UNCLASSIFIED.get(opcode)
    base = opcode & 0x1F
    class_bits = opcode & 0x60
    entry = _CLASSIFIED_BASE.get(base)
    if entry is None or class_bits not in _CLASS_SUFFIX:
        return None
    name, size = entry
    return name + _CLASS_SUFFIX[class_bits], size


@dataclass(slots=True)
class TokenWalkResult:
    """Aggregate-only outcome of walking one ``rgce`` byte stream."""

    opcode_counts: Counter[str] = field(default_factory=Counter)
    fully_classified: bool = False
    #: Byte offset within rgce where the walk stopped (== len(rgce) when
    #: `fully_classified` is True).
    stopped_at: int = 0
    #: The unrecognized/unsized opcode that stopped the walk, if any.
    stopped_opcode: int | None = None


def walk_ptg_tokens(rgce: bytes) -> TokenWalkResult:
    """Inventory token kinds in one formula's raw Ptg byte stream.

    Never raises; any truncation or unrecognized/unsized opcode stops the
    walk immediately rather than guessing a byte count.
    """
    result = TokenWalkResult()
    pos = 0
    n = len(rgce)
    while pos < n:
        opcode = rgce[pos]
        entry = _opcode_lookup(opcode)
        if entry is None:
            result.stopped_at = pos
            result.stopped_opcode = opcode
            return result
        name, size = entry
        if size is None:
            result.stopped_at = pos
            result.stopped_opcode = opcode
            return result
        if pos + 1 + size > n:
            result.stopped_at = pos
            result.stopped_opcode = opcode
            return result
        result.opcode_counts[name] += 1
        pos += 1 + size
    result.fully_classified = True
    result.stopped_at = pos
    return result


# --- record-payload prefix discovery ----------------------------------------

#: Candidate fixed-width cached-value widths to try before `cce`/`rgce`,
#: keyed by the naive-decoded record id already verified elsewhere in this
#: project (see qc_tool/io/xlsb_formula.py's record-id comment). Each
#: candidate is self-validated per record via the cce-consistency check
#: below -- this table is a starting guess, not an assumed truth.
_CANDIDATE_VALUE_WIDTHS: dict[int, tuple[int, ...]] = {
    0x0008: (0,),  # BrtFmlaString: cached value is itself a string -- skip
    0x0009: (8,),  # BrtFmlaNum: cached IEEE double
    0x000A: (1, 2, 4, 8),  # BrtFmlaBool: cached bool, width unverified
    0x000B: (1, 2, 4, 8),  # BrtFmlaError: cached error byte, width unverified
}
#: col(4) + ixfe/style(4) precede the cached value in every retained cell
#: record (proven -- pyxlsb reads exactly this for BrtFmlaNum today; see
#: tests/fixtures/xlsb_writer.py `_sheet_part`).
_CELL_PREFIX_BYTES = 8
#: Widest plausible search range for the offset of ``cce`` -- brute-forced
#: and accepted only via the exact arithmetic self-check below, since this
#: spike does not have a verified spec citation for the cached-value/flags
#: layout that precedes the formula token stream.
_MAX_CCE_SEARCH_OFFSET = 64


def locate_rgce(record_id: int, payload: bytes) -> bytes | None:
    """Best-effort, self-validated ``rgce`` slice, or None if inconsistent.

    Brute-forces every plausible ``cce`` field offset and accepts one only
    if the ``cce`` value it implies exactly matches the remaining payload
    length -- an exact arithmetic identity, not a guess accepted on faith.
    """
    if record_id not in _CANDIDATE_VALUE_WIDTHS:
        return None
    limit = min(_MAX_CCE_SEARCH_OFFSET, len(payload) - 4)
    for cce_offset in range(_CELL_PREFIX_BYTES, max(_CELL_PREFIX_BYTES, limit) + 1):
        if cce_offset + 4 > len(payload):
            break
        (cce,) = struct.unpack_from("<I", payload, cce_offset)
        rgce_offset = cce_offset + 4
        if cce and rgce_offset + cce == len(payload):
            return payload[rgce_offset : rgce_offset + cce]
    return None


def scan_worksheet_formula_tokens(
    data: bytes, formula_record_ids: frozenset[int]
) -> tuple[Counter[str], int, int]:
    """Aggregate opcode counts across every formula record in one worksheet
    part. Returns ``(opcode_counts, formulas_fully_classified,
    formulas_total)``. Reuses this project's own proven record reader
    (imported locally to avoid a hard dependency for callers that only want
    the pure token-walk logic above).
    """
    from qc_tool.io.xlsb_formula import (
        XlsbFormulaScanError,
        _read_record_id,
        _read_record_length,
    )

    counts: Counter[str] = Counter()
    total = 0
    fully = 0
    pos = 0
    n = len(data)
    while pos < n:
        try:
            record_id, pos = _read_record_id(data, pos)
            length, pos = _read_record_length(data, pos)
        except XlsbFormulaScanError:
            break  # truncated/malformed tail -- stop, never guess
        payload = data[pos : pos + length]
        pos += length
        if record_id not in formula_record_ids:
            continue
        rgce = locate_rgce(record_id, payload)
        if rgce is None:
            total += 1
            continue
        total += 1
        result = walk_ptg_tokens(rgce)
        counts.update(result.opcode_counts)
        if result.fully_classified:
            fully += 1
    return counts, fully, total
