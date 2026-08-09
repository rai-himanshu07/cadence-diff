"""Reporting-period evidence used by PowerPoint claim checks."""

from __future__ import annotations

import re

from qc_tool.excel.periods import Period, parse_period

_PERIOD_PATTERNS = (
    re.compile(
        r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[A-Za-z]*"
        r"[-_ ]?'?\d{2,4}\b",
        re.IGNORECASE,
    ),
    re.compile(r"\b\d{4}-\d{2}(?:-\d{2})?\b"),
    re.compile(
        r"\b(?:CW|WK|W)[-_ ]?\d{1,2}(?:[-_ ]?'?\d{2,4})?\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"\bQ[1-4](?:[-_ ]?(?:FY)?[-_ ]?'?\d{2,4})?\b",
        re.IGNORECASE,
    ),
)
CURRENT_PERIOD_CONTEXT = re.compile(
    r"\b(?:as of|reporting period|cycle|month ending|week ending|quarter ending)\b",
    re.IGNORECASE,
)


def periods_in_text(text: str) -> list[tuple[str, Period]]:
    """Return every recognizable reporting-period label in source order."""
    periods: list[tuple[int, int, str, Period]] = []
    for pattern_index, pattern in enumerate(_PERIOD_PATTERNS):
        for match in pattern.finditer(text):
            parsed = parse_period(match.group(0))
            if parsed is not None:
                periods.append(
                    (match.start(), pattern_index, match.group(0), parsed)
                )
    periods.sort(key=lambda item: (item[0], item[1]))
    return [(label, period) for _start, _pattern, label, period in periods]
