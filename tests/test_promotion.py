"""Reviewed-finding promotion into explicit reporting-contract rules."""

import datetime as dt

import pytest

from qc_tool.config.profile import DeliverableProfile
from qc_tool.config.promotion import (
    PromotionKind,
    PromotionRequest,
    available_promotions,
    promote_finding,
)
from qc_tool.findings import Finding, FindingClass


def _finding(finding_class: FindingClass) -> Finding:
    return Finding(
        finding_id="F1",
        artifact="excel",
        finding_class=finding_class,
        sheet="Data",
        location="B2",
        element="cell",
        message="synthetic",
    )


def test_waiver_promotion_is_narrow_scoped_and_requires_reason_and_expiry() -> None:
    profile = DeliverableProfile(name="contract")
    finding = _finding(FindingClass.FORMULA_LOGIC_CHANGED)

    draft = promote_finding(
        profile,
        finding,
        PromotionRequest(
            kind=PromotionKind.WAIVER,
            reason="approved exception",
            expires=dt.date(2026, 12, 31),
        ),
    )

    assert profile.waivers == []
    waiver = draft.waivers[0]
    assert waiver.finding_class is finding.finding_class
    assert waiver.sheet == "Data" and waiver.location == "B2"
    assert waiver.element == "cell"
    with pytest.raises(ValueError, match="reason and expiry"):
        PromotionRequest(kind=PromotionKind.WAIVER)


def test_control_templates_exist_only_for_compatible_finding_classes() -> None:
    required = _finding(FindingClass.REQUIRED_VALUE_MISSING)
    bound = _finding(FindingClass.NUMERIC_BOUND_VIOLATION)
    value = _finding(FindingClass.VALUE_CHANGED)
    formula = _finding(FindingClass.FORMULA_LOGIC_CHANGED)

    assert available_promotions(required) == (
        PromotionKind.WAIVER,
        PromotionKind.REQUIRED_RANGE,
    )
    assert PromotionKind.NUMERIC_BOUND in available_promotions(bound)
    assert PromotionKind.ACCEPTANCE_BAND in available_promotions(value)
    assert available_promotions(formula) == (PromotionKind.WAIVER,)

    with pytest.raises(ValueError, match="does not support"):
        promote_finding(
            DeliverableProfile(name="contract"),
            formula,
            PromotionRequest(kind=PromotionKind.REQUIRED_RANGE, name="bad"),
        )


def test_required_bound_and_acceptance_promotions_require_user_parameters() -> None:
    profile = DeliverableProfile(name="contract")
    required = promote_finding(
        profile,
        _finding(FindingClass.REQUIRED_VALUE_MISSING),
        PromotionRequest(kind=PromotionKind.REQUIRED_RANGE, name="Current KPI"),
    )
    assert required.excel.controls.required_ranges[0].cell_range == "B2"

    bounded = promote_finding(
        profile,
        _finding(FindingClass.NUMERIC_BOUND_VIOLATION),
        PromotionRequest(
            kind=PromotionKind.NUMERIC_BOUND,
            name="Margin",
            minimum=0,
            maximum=1,
        ),
    )
    assert bounded.excel.controls.numeric_bounds[0].maximum == 1

    accepted = promote_finding(
        profile,
        _finding(FindingClass.VALUE_CHANGED),
        PromotionRequest(
            kind=PromotionKind.ACCEPTANCE_BAND,
            name="Revenue tolerance",
            absolute=1,
        ),
    )
    assert accepted.excel.sheets["Data"].acceptance_bands[0].absolute == 1

    with pytest.raises(ValueError, match="minimum or maximum"):
        PromotionRequest(kind=PromotionKind.NUMERIC_BOUND, name="bad")
    with pytest.raises(ValueError, match="positive bound"):
        PromotionRequest(kind=PromotionKind.ACCEPTANCE_BAND, name="bad")
