"""Step 10: importing a private diagnostic script must never disable
logging process-globally (Criterion 18) -- the disable call belongs inside
each script's ``main()``, not at module level, so a future test importing
one of these modules for its pure helpers doesn't silently break an
unrelated later test's logging assertions.
"""

from __future__ import annotations

import importlib
import logging
import sys


def _assert_import_has_no_logging_side_effect(module_name: str) -> None:
    sys.modules.pop(module_name, None)
    logging.disable(logging.NOTSET)
    try:
        importlib.import_module(module_name)
        current = logging.root.manager.disable
    finally:
        sys.modules.pop(module_name, None)
        logging.disable(logging.NOTSET)
    assert current == logging.NOTSET, (
        f"importing {module_name} disabled logging process-globally"
    )


def test_formula_phase_diagnostic_import_has_no_logging_side_effect() -> None:
    _assert_import_has_no_logging_side_effect("scripts.formula_phase_diagnostic")


def test_windows_large_workbook_throughput_acceptance_import_has_no_logging_side_effect() -> None:
    _assert_import_has_no_logging_side_effect(
        "scripts.windows_large_workbook_throughput_acceptance"
    )


def test_xlsb_load_locality_diagnostic_import_has_no_logging_side_effect() -> None:
    _assert_import_has_no_logging_side_effect("scripts.xlsb_load_locality_diagnostic")
