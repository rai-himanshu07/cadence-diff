"""Drift guard for the closed field-classification registry
(``qc_tool.config.field_classification``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 1.
A newly added ``DeliverableProfile``/``RunRequest`` field must be classified
here before it can silently escape Re-QC scope-comparability review.
"""

from __future__ import annotations

import pytest

from qc_tool.config.field_classification import (
    PROFILE_FIELD_CLASSIFICATION,
    RUN_REQUEST_FIELD_CLASSIFICATION,
    classification_for,
    discover_profile_fields,
    discover_run_request_fields,
    stale_profile_classifications,
    stale_run_request_classifications,
    unclassified_profile_fields,
    unclassified_run_request_fields,
)


def test_every_discovered_profile_field_is_classified() -> None:
    assert unclassified_profile_fields() == set()


def test_no_stale_profile_classification_entries_remain() -> None:
    assert stale_profile_classifications() == set()


def test_every_discovered_run_request_field_is_classified() -> None:
    assert unclassified_run_request_fields() == set()


def test_no_stale_run_request_classification_entries_remain() -> None:
    assert stale_run_request_classifications() == set()


def test_discovery_finds_a_nonempty_field_set_for_both_models() -> None:
    """A regression guard against the walker silently discovering nothing
    (which would make the drift tests above vacuously pass).
    """
    assert len(discover_profile_fields()) > 50
    assert len(discover_run_request_fields()) == len(RUN_REQUEST_FIELD_CLASSIFICATION)


def test_all_four_categories_are_actually_used() -> None:
    used = set(PROFILE_FIELD_CLASSIFICATION.values()) | set(
        RUN_REQUEST_FIELD_CLASSIFICATION.values()
    )
    assert used == {
        "scope_semantics",
        "output_representation",
        "presentation_only",
        "operational_only",
    }


def test_classification_for_known_and_unknown_paths() -> None:
    assert classification_for("tolerance.absolute") == "scope_semantics"
    assert classification_for("requested_output_mode") == "output_representation"
    with pytest.raises(KeyError):
        classification_for("not.a.real.field")
