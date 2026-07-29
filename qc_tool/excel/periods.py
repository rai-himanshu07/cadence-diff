"""Period-label parsing: detection and ordering of cadence labels.

Real deliverables label periods inconsistently (weekly, monthly, quarterly;
"Jan-26", "2026-01", "W05", "Q1 FY26", real dates). The parser normalizes a
label into a `Period` with a sortable key so region detection can find
period axes and the alignment engine can order growth.
"""

import datetime as dt
import re
from dataclasses import dataclass
from typing import Literal

from qc_tool.io.model import display_cell_value

PeriodKind = Literal["month", "week", "quarter", "date"]

_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}  # fmt: skip

_MONTH_NAME_RE = re.compile(
    r"^(?P<mon>jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[-_ ]?'?(?P<year>\d{2,4})$",
    re.IGNORECASE,
)
_ISO_RE = re.compile(r"^(?P<year>\d{4})-(?P<mon>\d{2})(?:-(?P<day>\d{2}))?$")
_WEEK_RE = re.compile(
    r"^(?:CW|WK|W)[-_ ]?(?P<week>\d{1,2})"
    r"(?:[-_ ]?'?(?P<year>\d{2,4}))?$",
    re.IGNORECASE,
)
_ISO_WEEK_RE = re.compile(
    r"^(?P<year>\d{4})[-_ ]?(?:CW|WK|W)[-_ ]?(?P<week>\d{1,2})$",
    re.IGNORECASE,
)
_QUARTER_RE = re.compile(
    r"^Q(?P<q>[1-4])(?:[-_ ]?(?:FY)?[-_ ]?'?(?P<year>\d{2,4}))?$", re.IGNORECASE
)


@dataclass(frozen=True, slots=True)
class Period:
    kind: PeriodKind
    sort_key: tuple[int, int, int]
    label: str


def _year(raw: str | None) -> int:
    if raw is None:
        return 0
    year = int(raw)
    return year + 2000 if year < 100 else year


def parse_period(value: object) -> Period | None:
    """Parse a cell value into a `Period`, or None if it is not one."""
    if isinstance(value, dt.datetime | dt.date):
        return Period(
            kind="date", sort_key=(value.year, value.month, getattr(value, "day", 1)),
            label=display_cell_value(value),
        )
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None
    if match := _MONTH_NAME_RE.match(text):
        return Period(
            kind="month",
            sort_key=(_year(match["year"]), _MONTHS[match["mon"].lower()], 0),
            label=text,
        )
    if match := _ISO_RE.match(text):
        return Period(
            kind="date" if match["day"] else "month",
            sort_key=(int(match["year"]), int(match["mon"]), int(match["day"] or 0)),
            label=text,
        )
    if match := _ISO_WEEK_RE.match(text):
        return Period(
            kind="week",
            sort_key=(_year(match["year"]), 0, int(match["week"])),
            label=text,
        )
    if match := _WEEK_RE.match(text):
        return Period(
            kind="week",
            sort_key=(_year(match["year"]), 0, int(match["week"])),
            label=text,
        )
    if match := _QUARTER_RE.match(text):
        return Period(
            kind="quarter", sort_key=(_year(match["year"]), int(match["q"]), 0), label=text
        )
    return None


def is_period_label(value: object) -> bool:
    return parse_period(value) is not None


def is_period_after(current: Period, previous: Period) -> bool:
    """Return whether ``current`` follows ``previous`` in cadence order.

    Yearless week labels need a bounded rollover rule because ``W01`` has no
    intrinsic year. Only a late-year to early-year transition is inferred;
    other backwards week movements remain regressions.
    """
    if current.kind != previous.kind:
        return False
    if current.kind == "week":
        current_year, _, current_week = current.sort_key
        previous_year, _, previous_week = previous.sort_key
        if current_year == previous_year == 0 and current_week < previous_week:
            return previous_week >= 40 and current_week <= 13
    return current.sort_key > previous.sort_key
