"""Canonical tokenization for modern Excel reference syntax."""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import StrEnum
from functools import lru_cache
from typing import Literal

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token

DynamicReferenceKind = Literal["spill", "implicit"]
FormulaPatternKey = tuple[tuple[str, str, str], ...]

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
_LOCAL_SYMBOL_RE = re.compile(r"^[A-Z_\\][A-Z0-9_.\\]*$", re.IGNORECASE)


@dataclass(frozen=True, slots=True)
class FormulaReferenceOperand:
    value: str
    kind: DynamicReferenceKind | None = None


class FormulaPrecedentKind(StrEnum):
    REFERENCE = "reference"
    LOCAL_SYMBOL = "local_symbol"
    UNSUPPORTED = "unsupported"


@dataclass(frozen=True, slots=True)
class FormulaPrecedent:
    value: str
    kind: FormulaPrecedentKind
    dynamic: DynamicReferenceKind | None = None
    reason: str = ""


@dataclass(frozen=True, slots=True)
class FormulaPrecedentExtraction:
    precedents: tuple[FormulaPrecedent, ...]

    @property
    def references(self) -> tuple[FormulaPrecedent, ...]:
        return tuple(
            item
            for item in self.precedents
            if item.kind is FormulaPrecedentKind.REFERENCE
        )

    @property
    def local_symbols(self) -> tuple[FormulaPrecedent, ...]:
        return tuple(
            item
            for item in self.precedents
            if item.kind is FormulaPrecedentKind.LOCAL_SYMBOL
        )

    @property
    def unsupported(self) -> tuple[FormulaPrecedent, ...]:
        return tuple(
            item
            for item in self.precedents
            if item.kind is FormulaPrecedentKind.UNSUPPORTED
        )


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


@lru_cache(maxsize=4096)
def _tokenize_formula(formula: str) -> tuple[Token, ...]:
    return tuple(Tokenizer(canonicalize_dynamic_formula(formula)).items)


def tokenize_formula(formula: str) -> list[Token]:
    return list(_tokenize_formula(formula))


def formula_pattern_key(formula: str) -> FormulaPatternKey:
    """Return a case-normalized token key while preserving quoted text."""
    return tuple(
        (
            token.type,
            token.subtype,
            (
                token.value
                if token.type == "OPERAND" and token.subtype == "TEXT"
                else token.value.casefold()
            ),
        )
        for token in _tokenize_formula(formula)
    )


def _function_name(token: Token) -> str:
    return token.value[:-1].casefold().rsplit(".", 1)[-1]


def _function_closes(tokens: tuple[Token, ...]) -> dict[int, int]:
    stack: list[int] = []
    closes: dict[int, int] = {}
    for position, token in enumerate(tokens):
        if token.type == "FUNC" and token.subtype == "OPEN":
            stack.append(position)
        elif token.type == "FUNC" and token.subtype == "CLOSE":
            if not stack:
                raise ValueError("formula contains an unmatched function close")
            closes[stack.pop()] = position
    if stack:
        raise ValueError("formula contains an unclosed function")
    return closes


def _argument_ranges(
    tokens: tuple[Token, ...], start: int, end: int
) -> tuple[tuple[int, int], ...]:
    ranges: list[tuple[int, int]] = []
    argument_start = start
    depth = 0
    for position in range(start, end):
        token = tokens[position]
        if token.subtype == "OPEN" and token.type in {"FUNC", "PAREN", "ARRAY"}:
            depth += 1
        elif token.subtype == "CLOSE" and token.type in {"FUNC", "PAREN", "ARRAY"}:
            depth -= 1
        elif token.type == "SEP" and token.subtype == "ARG" and depth == 0:
            ranges.append((argument_start, position))
            argument_start = position + 1
    ranges.append((argument_start, end))
    return tuple(ranges)


def _local_symbol(
    tokens: tuple[Token, ...], start: int, end: int
) -> str | None:
    significant = [token for token in tokens[start:end] if token.type != "WSPACE"]
    if len(significant) != 1:
        return None
    token = significant[0]
    if token.type != "OPERAND" or token.subtype != "RANGE":
        return None
    value = token.value
    if _LOCAL_SYMBOL_RE.fullmatch(value) is None:
        return None
    if re.fullmatch(_CELL, value, re.IGNORECASE) is not None:
        return None
    if value.casefold() in {"r", "c"}:
        return None
    return value


class _PrecedentExtractor:
    def __init__(self, tokens: tuple[Token, ...]) -> None:
        self._tokens = tokens
        self._closes = _function_closes(tokens)

    @staticmethod
    def _local(value: str) -> FormulaPrecedent:
        return FormulaPrecedent(value, FormulaPrecedentKind.LOCAL_SYMBOL)

    @staticmethod
    def _unsupported(reason: str) -> FormulaPrecedent:
        return FormulaPrecedent("", FormulaPrecedentKind.UNSUPPORTED, reason=reason)

    def _walk_let(
        self,
        start: int,
        end: int,
        scope: frozenset[str],
    ) -> list[FormulaPrecedent]:
        arguments = _argument_ranges(self._tokens, start, end)
        if len(arguments) < 3 or len(arguments) % 2 == 0:
            return [self._unsupported("malformed_let")]
        bindings: list[str] = []
        for argument in arguments[:-1:2]:
            binding = _local_symbol(self._tokens, *argument)
            if binding is None or binding.casefold() in bindings:
                return [self._unsupported("malformed_let")]
            bindings.append(binding.casefold())

        precedents: list[FormulaPrecedent] = []
        current_scope = scope
        for pair_index in range(0, len(arguments) - 1, 2):
            binding = _local_symbol(self._tokens, *arguments[pair_index])
            if binding is None:  # validated above
                return [self._unsupported("malformed_let")]
            precedents.append(self._local(binding))
            precedents.extend(
                self._walk(*arguments[pair_index + 1], current_scope, None)
            )
            current_scope = current_scope | {binding.casefold()}
        precedents.extend(self._walk(*arguments[-1], current_scope, None))
        return precedents

    def _walk_lambda(
        self,
        start: int,
        end: int,
        scope: frozenset[str],
    ) -> list[FormulaPrecedent]:
        arguments = _argument_ranges(self._tokens, start, end)
        if len(arguments) < 2:
            return [self._unsupported("malformed_lambda")]
        parameters: list[str] = []
        parameter_values: list[str] = []
        for argument in arguments[:-1]:
            parameter = _local_symbol(self._tokens, *argument)
            if parameter is None or parameter.casefold() in parameters:
                return [self._unsupported("malformed_lambda")]
            parameters.append(parameter.casefold())
            parameter_values.append(parameter)
        precedents = [self._local(parameter) for parameter in parameter_values]
        precedents.extend(
            self._walk(*arguments[-1], scope | set(parameters), None)
        )
        return precedents

    def _walk(
        self,
        start: int,
        end: int,
        scope: frozenset[str],
        dynamic: DynamicReferenceKind | None,
    ) -> list[FormulaPrecedent]:
        precedents: list[FormulaPrecedent] = []
        position = start
        while position < end:
            token = self._tokens[position]
            if token.type == "FUNC" and token.subtype == "OPEN":
                close = self._closes.get(position)
                if close is None or close >= end:
                    raise ValueError("formula function boundary is malformed")
                function = _function_name(token)
                if function == "let":
                    precedents.extend(self._walk_let(position + 1, close, scope))
                elif function == "lambda":
                    precedents.extend(self._walk_lambda(position + 1, close, scope))
                else:
                    function_dynamic: DynamicReferenceKind | None
                    if function == "anchorarray":
                        function_dynamic = "spill"
                    elif function == "single":
                        function_dynamic = "implicit"
                    else:
                        function_dynamic = None
                    precedents.extend(
                        self._walk(position + 1, close, scope, function_dynamic)
                    )
                position = close + 1
                continue
            if token.type == "OPERAND" and token.subtype == "RANGE":
                if token.value.casefold() in scope:
                    precedents.append(self._local(token.value))
                else:
                    value = token.value
                    if dynamic == "spill":
                        value = f"{value}#"
                    elif dynamic == "implicit":
                        value = f"@{value}"
                    precedents.append(
                        FormulaPrecedent(
                            value,
                            FormulaPrecedentKind.REFERENCE,
                            dynamic=dynamic,
                        )
                    )
            position += 1
        return precedents

    def extract(self) -> FormulaPrecedentExtraction:
        return FormulaPrecedentExtraction(
            tuple(self._walk(0, len(self._tokens), frozenset(), None))
        )


@lru_cache(maxsize=4096)
def extract_formula_precedents(formula: str) -> FormulaPrecedentExtraction:
    """Classify formula operands without resolving lexical symbols as cells."""
    return _PrecedentExtractor(_tokenize_formula(formula)).extract()


def formula_reference_operands(formula: str) -> list[FormulaReferenceOperand]:
    return [
        FormulaReferenceOperand(precedent.value, precedent.dynamic)
        for precedent in extract_formula_precedents(formula).references
    ]


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
