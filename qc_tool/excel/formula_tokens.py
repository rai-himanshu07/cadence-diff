"""Canonical tokenization for modern Excel reference syntax."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token

DynamicReferenceKind = Literal["spill", "implicit"]

_CELL = r"\$?[A-Z]{1,3}\$?[1-9]\d*"
_SHEET = r"(?:'(?:[^']|'')+'|[A-Z_\\][A-Z0-9_.\\]*)!"
_REFERENCE = rf"(?:{_SHEET})?{_CELL}"
_RANGE = rf"{_REFERENCE}(?::{_CELL})?"
_SPILL_RE = re.compile(
    rf"(?<![A-Z0-9_.\[\]])(?P<reference>{_REFERENCE})#",
    re.IGNORECASE,
)
_IMPLICIT_RE = re.compile(
    rf"(?<![A-Z0-9_.\[\]])@(?P<reference>{_RANGE})",
    re.IGNORECASE,
)
_ANCHOR_FUNCTION_RE = re.compile(
    r"(?<![A-Z0-9_.])(?:_xlfn\.)?ANCHORARRAY\(",
    re.IGNORECASE,
)
_SINGLE_FUNCTION_RE = re.compile(
    r"(?<![A-Z0-9_.])(?:_xlfn\.)?SINGLE\(",
    re.IGNORECASE,
)
_ANCHOR_WRAPPER_RE = re.compile(
    r"^(?:_xlfn\.)?ANCHORARRAY\((?P<reference>.+)\)$",
    re.IGNORECASE,
)
_SINGLE_WRAPPER_RE = re.compile(
    r"^(?:_xlfn\.)?SINGLE\((?P<reference>.+)\)$",
    re.IGNORECASE,
)


@dataclass(frozen=True, slots=True)
class FormulaReferenceOperand:
    value: str
    kind: DynamicReferenceKind | None = None


def _rewrite_segment(segment: str) -> str:
    segment = _ANCHOR_FUNCTION_RE.sub("_xlfn.ANCHORARRAY(", segment)
    segment = _SINGLE_FUNCTION_RE.sub("_xlfn.SINGLE(", segment)
    segment = _SPILL_RE.sub(
        lambda match: f"_xlfn.ANCHORARRAY({match.group('reference')})",
        segment,
    )
    return _IMPLICIT_RE.sub(
        lambda match: f"_xlfn.SINGLE({match.group('reference')})",
        segment,
    )


def canonicalize_dynamic_formula(formula: str) -> str:
    """Rewrite unquoted spill/intersection syntax into canonical wrappers."""
    rendered: list[str] = []
    segment: list[str] = []
    index = 0
    while index < len(formula):
        character = formula[index]
        if character != '"':
            segment.append(character)
            index += 1
            continue
        rendered.append(_rewrite_segment("".join(segment)))
        segment.clear()
        quoted = ['"']
        index += 1
        while index < len(formula):
            quoted.append(formula[index])
            if formula[index] != '"':
                index += 1
                continue
            if index + 1 < len(formula) and formula[index + 1] == '"':
                quoted.append('"')
                index += 2
                continue
            index += 1
            break
        rendered.append("".join(quoted))
    rendered.append(_rewrite_segment("".join(segment)))
    return "".join(rendered)


def tokenize_formula(formula: str) -> list[Token]:
    return Tokenizer(canonicalize_dynamic_formula(formula)).items


def formula_reference_operands(formula: str) -> list[FormulaReferenceOperand]:
    operands: list[FormulaReferenceOperand] = []
    function_stack: list[DynamicReferenceKind | None] = []
    for token in tokenize_formula(formula):
        if token.type == "FUNC" and token.subtype == "OPEN":
            function = token.value[:-1].casefold()
            if function.endswith("anchorarray"):
                function_stack.append("spill")
            elif function.endswith("single"):
                function_stack.append("implicit")
            else:
                function_stack.append(None)
            continue
        if token.type == "FUNC" and token.subtype == "CLOSE":
            if function_stack:
                function_stack.pop()
            continue
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        kind = function_stack[-1] if function_stack else None
        if kind == "spill":
            operands.append(FormulaReferenceOperand(f"{token.value}#", kind))
        elif kind == "implicit":
            operands.append(FormulaReferenceOperand(f"@{token.value}", kind))
        else:
            operands.append(FormulaReferenceOperand(token.value))
    return operands


def parse_dynamic_reference(
    target: str,
) -> tuple[DynamicReferenceKind, str] | None:
    normalized = target.strip()
    anchor = _ANCHOR_WRAPPER_RE.fullmatch(normalized)
    if anchor is not None:
        return "spill", anchor.group("reference").strip()
    single = _SINGLE_WRAPPER_RE.fullmatch(normalized)
    if single is not None:
        return "implicit", single.group("reference").strip()
    if normalized.endswith("#") and "#REF!" not in normalized.upper():
        return "spill", normalized[:-1].strip()
    if normalized.startswith("@"):
        return "implicit", normalized[1:].strip()
    return None
