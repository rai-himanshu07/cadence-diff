"""Package skeleton smoke test: every subpackage must import."""

import importlib

import pytest

MODULES = [
    "qc_tool",
    "qc_tool.config",
    "qc_tool.io",
    "qc_tool.excel",
    "qc_tool.ppt",
    "qc_tool.crosscheck",
    "qc_tool.triage",
    "qc_tool.report",
    "qc_tool.history",
    "qc_tool.ui",
]


@pytest.mark.parametrize("name", MODULES)
def test_module_imports(name: str) -> None:
    importlib.import_module(name)
