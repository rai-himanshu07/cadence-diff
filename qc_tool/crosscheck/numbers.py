"""Format-aware figure parsing and display-precision matching.

Deck figures are rounded, scaled, and decorated ("$1.2M", "12%",
"(1,234)"). A figure matches a workbook value when the value, projected
into the figure's display space (scale applied, percent multiplied out),
falls within half a unit of the figure's last displayed decimal — robust
to the different rounding modes of Excel and Python.
"""

import re
from dataclasses import dataclass

_SUFFIX_SCALE = {
    "k": 1e3,
    "m": 1e6,
    "mn": 1e6,
    "b": 1e9,
    "bn": 1e9,
}

_FIGURE_RE = re.compile(
    r"(?<![\w.,-])"  # not glued to a word, number, or hyphenated label (Jan-26)
    r"(?P<open>\()?\s*"
    r"(?P<currency>[$\u20ac\u00a3\u20b9])?\s*"
    r"(?P<sign>-)?"
    r"(?P<number>\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)"
    r"\s*(?P<suffix>%|bn|mn|k|m|b)?"
    r"(?P<close>\))?"
    r"(?![\w%])",
    re.IGNORECASE,
)

#: Numeric tokens blanked out — used for line skeletons that stay stable
#: while figures refresh each cycle.
_NUMERIC_TOKEN_RE = re.compile(r"[-+]?[\d.,]+\s*%?")


def numeric_skeleton(text: str) -> str:
    """Text with numeric tokens blanked; equal skeletons = same wording."""
    return _NUMERIC_TOKEN_RE.sub("#", text)


@dataclass(frozen=True, slots=True)
class ParsedFigure:
    raw: str
    mantissa: float  # the number as displayed (1.2 for "$1.2M")
    scale: float  # 1e6 for M, 1e3 for k, 1 otherwise
    is_percent: bool
    decimals: int  # displayed decimal places of the mantissa
    negative: bool

    @property
    def value(self) -> float:
        """Normalized numeric value ($1.2M -> 1_200_000; 12% -> 0.12)."""
        signed = -self.mantissa if self.negative else self.mantissa
        if self.is_percent:
            return signed / 100.0
        return signed * self.scale


def _from_match(match: re.Match[str]) -> ParsedFigure:
    number = match["number"].replace(",", "")
    mantissa = float(number)
    decimals = len(number.rsplit(".", 1)[1]) if "." in number else 0
    suffix = (match["suffix"] or "").lower()
    negative = bool(match["sign"]) or bool(match["open"] and match["close"])
    return ParsedFigure(
        raw=match.group(0).strip(),
        mantissa=mantissa,
        scale=_SUFFIX_SCALE.get(suffix, 1.0),
        is_percent=suffix == "%",
        decimals=decimals,
        negative=negative,
    )


def parse_figure(text: str) -> ParsedFigure | None:
    """Parse a single figure token; None if the text is not one figure."""
    match = _FIGURE_RE.search(text.strip())
    if match is None or match.group(0).strip() != text.strip():
        return None
    return _from_match(match)


def extract_figures(text: str) -> list[ParsedFigure]:
    """All figure tokens in a line of display text."""
    return [_from_match(m) for m in _FIGURE_RE.finditer(text)]


def display_matches(figure: ParsedFigure, cell_value: float) -> bool:
    """True if ``cell_value`` displays as ``figure`` at its shown precision."""
    scaled = cell_value * 100.0 if figure.is_percent else cell_value / figure.scale
    shown = -figure.mantissa if figure.negative else figure.mantissa
    tolerance = 0.5 * 10.0 ** (-figure.decimals) + 1e-9
    return abs(scaled - shown) <= tolerance


def relative_difference(figure: ParsedFigure, cell_value: float) -> float:
    """Relative distance between a figure and a value (for near-miss ranking)."""
    reference = figure.value
    if reference == 0:
        return abs(cell_value)
    return abs(cell_value - reference) / abs(reference)


def format_figure_like(figure: ParsedFigure, value: float) -> str:
    """Render ``value`` with the source figure's currency/scale/precision style."""
    raw = figure.raw.strip()
    currency_match = re.search(r"[$\u20ac\u00a3\u20b9]", raw)
    currency = currency_match.group(0) if currency_match else ""
    suffix_match = re.search(r"(%|bn|mn|k|m|b)\s*\)?$", raw, re.IGNORECASE)
    suffix = suffix_match.group(1) if suffix_match else ""
    scaled = value * 100.0 if figure.is_percent else value / figure.scale
    negative = scaled < 0
    magnitude = abs(scaled)
    use_grouping = "," in raw
    number = (
        f"{magnitude:,.{figure.decimals}f}"
        if use_grouping
        else f"{magnitude:.{figure.decimals}f}"
    )
    rendered = f"{currency}{number}{suffix}"
    if negative:
        return f"({rendered})" if raw.startswith("(") else f"-{rendered}"
    return rendered


def replace_figure_ordinals(
    text: str, replacements: dict[int, str], *, start_index: int = 0
) -> tuple[str, int]:
    """Replace global figure ordinals in one text fragment; return next index."""
    index = start_index

    def replace(match: re.Match[str]) -> str:
        nonlocal index
        current = index
        index += 1
        replacement = replacements.get(current)
        if replacement is None:
            return match.group(0)
        raw = match.group(0)
        leading = raw[: len(raw) - len(raw.lstrip())]
        trailing = raw[len(raw.rstrip()) :]
        return f"{leading}{replacement}{trailing}"

    return _FIGURE_RE.sub(replace, text), index
