"""Deterministic domain data shared by the Excel and PowerPoint fixture builders.

All values are closed-form functions of (period, region) so that the workbook,
the deck, and the cross-check ground truth always agree without an evaluation
engine. Seeded-defect deltas are named constants so tests can reference them.
"""

MONTH_LABELS = ["Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26"]
REGION_LABELS = ["North", "South", "East", "West"]
WEEK_LABELS = [f"W{i:02d}" for i in range(1, 22)]
HEADCOUNT_REGIONS = ["North", "South"]

BASELINE_MONTHS = 5
CURRENT_MONTHS = 6
BASELINE_WEEKS = 20
CURRENT_WEEKS = 21
DELETED_WEEK = "W03"
ROLLING_WINDOW = 4

# Seeded-defect deltas (single source of truth for builders and tests).
E01_DELTA = 1234.0  # Long_Monthly!C7 historical revenue edit (Feb-26 / South)
E06_DELTA = 777.0  # Wide_Weekly W05 revenue edit
P04_DELTA = 500.0  # deck table North / Apr-26 edit
P05_DELTA = 2000.0  # deck trend chart Mar-26 edit
X03_MARGIN_OFFSET = 0.007  # KPI slide margin figure offset vs workbook

FIXTURE_PASSWORD = "qc-test"


def monthly_revenue(month: int, region: int) -> float:
    """Baseline revenue for a month index (0-based) and region index."""
    return 100_000.0 + 5_137.0 * month + 11_213.0 * region + ((month * 7 + region * 13) % 97) * 10

def monthly_cost(month: int, region: int) -> float:
    return round(monthly_revenue(month, region) * 0.62, 2)


def current_monthly_revenue(month: int, region: int) -> float:
    """Current-cycle revenue: baseline values plus the seeded E01 edit."""
    value = monthly_revenue(month, region)
    if (month, region) == (1, 1):  # Feb-26 / South -> Long_Monthly!C7
        value += E01_DELTA
    return value


def weekly_revenue(week: int) -> float:
    """Revenue for a week index (0-based over WEEK_LABELS)."""
    return 24_000.0 + 310.0 * week + ((week * 17) % 23) * 7.0


def weekly_cost(week: int) -> float:
    return round(weekly_revenue(week) * 0.58, 2)


def headcount(month: int, region: int) -> float:
    return 40.0 + 3.0 * month + 5.0 * region


def month_total_revenue(month: int, *, current: bool) -> float:
    fn = current_monthly_revenue if current else monthly_revenue
    return sum(fn(month, r) for r in range(len(REGION_LABELS)))


def month_total_margin(month: int, *, current: bool) -> float:
    fn = current_monthly_revenue if current else monthly_revenue
    return sum(fn(month, r) - monthly_cost(month, r) for r in range(len(REGION_LABELS)))


def total_revenue(*, current: bool) -> float:
    months = CURRENT_MONTHS if current else BASELINE_MONTHS
    return sum(month_total_revenue(m, current=current) for m in range(months))


def total_cost(*, current: bool) -> float:
    months = CURRENT_MONTHS if current else BASELINE_MONTHS
    return sum(monthly_cost(m, r) for m in range(months) for r in range(len(REGION_LABELS)))


def margin_ratio(*, current: bool) -> float:
    rev = total_revenue(current=current)
    return (rev - total_cost(current=current)) / rev


def fmt_millions(value: float) -> str:
    return f"${value / 1_000_000:.2f}M"


def fmt_pct(value: float) -> str:
    return f"{value:.1%}"
