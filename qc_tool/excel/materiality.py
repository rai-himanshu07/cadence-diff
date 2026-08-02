"""Numeric magnitude, display evidence, and local temporal context.

A changed constant is *representation noise* when an analyst reviewing the
rendered deliverable could not see it and the delta is floating-point-scale:
either within a few ULPs of the larger magnitude, or display-identical under
the cell's own number format while relatively tiny. Display equivalence uses
a bounded number-format parser (General plus fixed/thousands/percent/
scientific decimal families); anything else fails closed to the ULP rule so
an unparseable format can never manufacture a noise claim.

Magnitude precedence is noise, within_tolerance, then material. Recency is a
separate evidence axis and never changes the magnitude classification.
"""

import math
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import date
from enum import StrEnum
from itertools import pairwise

from qc_tool.config.profile import RestatementWindows
from qc_tool.excel.periods import Period, PeriodKind, is_period_after
from qc_tool.findings import FindingEvidenceTag, FindingTemporalContext, Materiality

#: Delta at or below this many ULPs of the larger magnitude is always noise.
NOISE_ULPS = 4
#: Display-identical deltas must also be relatively tiny to count as noise,
#: so coarse formats (e.g. millions scaling) cannot hide real restatements.
NOISE_MAX_RELATIVE = 1e-6
#: Below this magnitude, display-identical residues are noise regardless of
#: relative delta (sum residues around zero have unbounded relative error).
NOISE_ZERO_FLOOR = 1e-9

_GENERAL_SIG_DECIMALS = 10  # 11 significant digits, Excel's General ceiling
_DIGIT_PLACEHOLDERS = frozenset("0#?")
_DATETIME_LETTERS = frozenset("ymdhsagYMDHSAG")


class DisplayKind(StrEnum):
    GENERAL = "general"
    DECIMAL = "decimal"
    SCIENTIFIC = "scientific"


@dataclass(frozen=True, slots=True)
class DisplayRule:
    """How a supported number format renders a numeric value."""

    kind: DisplayKind
    decimals: int = 0
    percents: int = 0
    scaling_commas: int = 0


def _first_section(format_code: str) -> str | None:
    """Positive-value section, honoring quotes, escapes, and pad characters."""
    section: list[str] = []
    i = 0
    length = len(format_code)
    while i < length:
        char = format_code[i]
        if char == '"':
            end = format_code.find('"', i + 1)
            if end == -1:
                return None
            section.append(format_code[i : end + 1])
            i = end + 1
            continue
        if char in "\\_*":
            if i + 1 >= length:
                return None
            section.append(format_code[i : i + 2])
            i += 2
            continue
        if char == ";":
            break
        section.append(char)
        i += 1
    return "".join(section)


def parse_number_format(format_code: str) -> DisplayRule | None:
    """Parse a supported numeric format; None means unparseable (fail closed).

    Supported: General; fixed-decimal families with thousands separators,
    literal text, colors/conditions/currency brackets, percent scaling, and
    trailing-comma thousands scaling; decimal scientific notation. Date/time,
    fraction, and text formats are rejected.
    """
    section = _first_section(format_code)
    if section is None:
        return None
    if section.strip().lower() == "general":
        return DisplayRule(kind=DisplayKind.GENERAL)

    percents = 0
    scientific = False
    seen_decimal_point = False
    in_exponent = False
    decimals = 0
    placeholder_count = 0
    last_placeholder_index = -1
    comma_indices: list[tuple[int, bool]] = []  # (index, in_decimal_region)

    i = 0
    length = len(section)
    while i < length:
        char = section[i]
        if char == '"':
            end = section.find('"', i + 1)
            if end == -1:
                return None
            i = end + 1
            continue
        if char in "\\_*":
            i += 2
            continue
        if char == "[":
            end = section.find("]", i + 1)
            if end == -1:
                return None
            i = end + 1
            continue
        if char == "%":
            percents += 1
            i += 1
            continue
        if char in "Ee":
            if i + 1 < length and section[i + 1] in "+-":
                if scientific:
                    return None
                scientific = True
                in_exponent = True
                i += 2
                continue
            return None
        if char in _DATETIME_LETTERS:
            return None
        if char in "@/":
            return None
        if char in _DIGIT_PLACEHOLDERS:
            if not in_exponent:
                placeholder_count += 1
                last_placeholder_index = i
                if seen_decimal_point:
                    decimals += 1
            i += 1
            continue
        if char == ".":
            if seen_decimal_point or in_exponent:
                return None
            seen_decimal_point = True
            i += 1
            continue
        if char == ",":
            if not in_exponent:
                comma_indices.append((i, seen_decimal_point))
            i += 1
            continue
        # Other literal characters (currency symbols, parentheses, spaces...)
        i += 1

    if placeholder_count == 0:
        return None

    scaling_commas = 0
    for index, in_decimal_region in comma_indices:
        if index > last_placeholder_index:
            scaling_commas += 1
        elif in_decimal_region:
            return None  # separator inside the decimal digits is malformed

    return DisplayRule(
        kind=DisplayKind.SCIENTIFIC if scientific else DisplayKind.DECIMAL,
        decimals=decimals,
        percents=percents,
        scaling_commas=scaling_commas,
    )


def rendered_number(value: float, rule: DisplayRule) -> str | None:
    """Canonical rendered form of `value` under `rule` for equivalence tests.

    Optional `#`/`?` decimals render at full width: Excel trims trailing
    zeros after rounding, so equal rounded forms are equal displayed forms.
    """
    if not math.isfinite(value):
        return None
    if rule.kind is DisplayKind.GENERAL:
        text = format(value + 0.0, f".{_GENERAL_SIG_DECIMALS}e")
    else:
        scaled = value * (100.0**rule.percents) / (1000.0**rule.scaling_commas)
        if not math.isfinite(scaled):
            return None
        spec = "e" if rule.kind is DisplayKind.SCIENTIFIC else "f"
        text = format(scaled, f".{rule.decimals}{spec}")
    if text.startswith("-") and float(text) == 0.0:
        text = text[1:]
    return text


def _as_finite_float(value: object) -> float | None:
    if isinstance(value, bool) or not isinstance(value, int | float):
        return None
    result = float(value)
    return result if math.isfinite(result) else None


def within_ulps(base: float, curr: float, ulps: int = NOISE_ULPS) -> bool:
    """True when the delta is at most `ulps` ULPs of the larger magnitude."""
    delta = abs(curr - base)
    if delta == 0.0:
        return True
    scale = max(abs(base), abs(curr))
    return delta <= ulps * math.ulp(scale)


def is_representation_noise(
    base: object,
    curr: object,
    number_format: str | None,
    *,
    ulps: int = NOISE_ULPS,
) -> bool:
    """True when a changed numeric pair is floating-point/display noise.

    Equal pairs return False: they never produce a finding, and this
    predicate describes *changed* values only.
    """
    base_float = _as_finite_float(base)
    curr_float = _as_finite_float(curr)
    if base_float is None or curr_float is None or base_float == curr_float:
        return False
    if within_ulps(base_float, curr_float, ulps):
        return True
    rule = parse_number_format(number_format or "")
    if rule is None:
        return False
    base_rendered = rendered_number(base_float, rule)
    curr_rendered = rendered_number(curr_float, rule)
    if base_rendered is None or base_rendered != curr_rendered:
        return False
    delta = abs(curr_float - base_float)
    scale = max(abs(base_float), abs(curr_float))
    return delta / scale <= NOISE_MAX_RELATIVE or scale <= NOISE_ZERO_FLOOR


def numeric_evidence_tags(
    base: object,
    curr: object,
    number_format: str | None,
    *,
    ulps: int = NOISE_ULPS,
) -> set[FindingEvidenceTag]:
    """Directly proved numeric representation evidence for a changed pair."""
    base_float = _as_finite_float(base)
    curr_float = _as_finite_float(curr)
    if base_float is None or curr_float is None or base_float == curr_float:
        return set()
    evidence: set[FindingEvidenceTag] = set()
    if within_ulps(base_float, curr_float, ulps):
        evidence.add(FindingEvidenceTag.ULP_SCALE)
    rule = parse_number_format(number_format or "")
    if rule is not None:
        base_rendered = rendered_number(base_float, rule)
        curr_rendered = rendered_number(curr_float, rule)
        if base_rendered is not None and base_rendered == curr_rendered:
            evidence.add(FindingEvidenceTag.DISPLAY_EQUIVALENT)
    return evidence


def classify_numeric_pair(
    base: object,
    curr: object,
    number_format: str | None,
    *,
    within_acceptance: bool = False,
    ulps: int = NOISE_ULPS,
) -> Materiality | None:
    """Materiality tier for a changed pair; None when either side is non-numeric."""
    if _as_finite_float(base) is None or _as_finite_float(curr) is None:
        return None
    if is_representation_noise(base, curr, number_format, ulps=ulps):
        return Materiality.NOISE
    if within_acceptance:
        return Materiality.WITHIN_TOLERANCE
    return Materiality.MATERIAL


def is_anomalous_magnitude(base: object, curr: object) -> bool:
    """Whether a numeric change crosses a hard analyst-review boundary."""
    base_float = _as_finite_float(base)
    curr_float = _as_finite_float(curr)
    if base_float is None or curr_float is None or base_float == curr_float:
        return False
    if base_float == 0.0 or curr_float == 0.0:
        return True
    if (base_float < 0.0) != (curr_float < 0.0):
        return True
    ratio = max(abs(base_float), abs(curr_float)) / min(
        abs(base_float), abs(curr_float)
    )
    return ratio >= 10.0


#: Median day spacing at or below each bound infers that cadence for
#: date-keyed axes. Sparser axes get no recency window (fail closed).
_DATE_CADENCE_BOUNDS: tuple[tuple[int, PeriodKind], ...] = (
    (10, "week"),
    (45, "month"),
    (120, "quarter"),
)


def infer_date_cadence(periods: list[Period]) -> PeriodKind | None:
    """Cadence of a date-keyed axis from the median gap of distinct dates."""
    keys = sorted({period.sort_key for period in periods if period.kind == "date"})
    if len(keys) < 2:
        return None
    gaps = sorted((date(*b) - date(*a)).days for a, b in pairwise(keys))
    median_gap = gaps[len(gaps) // 2]
    for bound, kind in _DATE_CADENCE_BOUNDS:
        if median_gap <= bound:
            return kind
    return None


def _window_for(kind: PeriodKind, windows: RestatementWindows) -> int:
    if kind == "week":
        return windows.week
    if kind == "month":
        return windows.month
    if kind == "quarter":
        return windows.quarter
    return 0


def _period_bands(
    periods_by_position: Mapping[int, Period],
) -> list[dict[int, Period]]:
    """Split one axis into independent contiguous cadence runs."""
    bands: list[dict[int, Period]] = []
    current: dict[int, Period] = {}
    previous_position: int | None = None
    previous_period: Period | None = None
    for position, period in sorted(periods_by_position.items()):
        continues = (
            previous_position is not None
            and previous_period is not None
            and position == previous_position + 1
            and period.kind == previous_period.kind
            and (
                period.sort_key >= previous_period.sort_key
                or is_period_after(period, previous_period)
            )
        )
        if current and not continues:
            bands.append(current)
            current = {}
        current[position] = period
        previous_position = position
        previous_period = period
    if current:
        bands.append(current)
    return bands


def _recent_band_positions(
    periods_by_position: Mapping[int, Period],
    windows: RestatementWindows,
) -> set[int]:
    periods = list(periods_by_position.values())
    if not periods:
        return set()
    kind = periods[0].kind
    window_kind: PeriodKind | None = kind
    if kind == "date":
        window_kind = infer_date_cadence(periods)
    window = _window_for(window_kind, windows) if window_kind is not None else 0
    if window <= 0:
        return set()
    recent_keys = set(
        sorted({period.sort_key for period in periods}, reverse=True)[:window]
    )
    return {
        position
        for position, period in periods_by_position.items()
        if period.sort_key in recent_keys
    }


def recent_positions(
    periods_by_position: Mapping[int, Period],
    windows: RestatementWindows,
) -> set[int]:
    """Axis positions whose period lies in the trailing restatement window.

    Windows count distinct periods back from each cadence kind's own growth
    edge, so mixed-cadence axes and irregular calendars need no date
    arithmetic. Date-kind periods use the window of their inferred cadence;
    positions of unparseable or unknown-cadence periods are never recent.
    """
    return set().union(
        *(
            _recent_band_positions(band, windows)
            for band in _period_bands(periods_by_position)
        )
    ) if periods_by_position else set()


def temporal_contexts(
    periods_by_position: Mapping[int, Period],
    windows: RestatementWindows,
) -> dict[int, FindingTemporalContext]:
    """Classify each proved period position against its local observed edge."""
    contexts: dict[int, FindingTemporalContext] = {}
    for band in _period_bands(periods_by_position):
        recent = _recent_band_positions(band, windows)
        latest = max(period.sort_key for period in band.values())
        for position, period in band.items():
            if position not in recent:
                contexts[position] = FindingTemporalContext.HISTORICAL
            elif period.sort_key == latest:
                contexts[position] = FindingTemporalContext.CURRENT_PERIOD
            else:
                contexts[position] = FindingTemporalContext.RECENT_WINDOW
    return contexts
