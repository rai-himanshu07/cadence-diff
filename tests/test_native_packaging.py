"""Release contracts for the optional stable-ABI XLSB kernel wheel."""

from __future__ import annotations

import tomllib
from pathlib import Path

import yaml

_ROOT = Path(__file__).parents[1]


def test_native_extra_and_stable_abi_are_declared() -> None:
    project = tomllib.loads((_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    cargo = tomllib.loads(
        (_ROOT / "native/xlsbkernel/Cargo.toml").read_text(encoding="utf-8")
    )

    assert project["project"]["optional-dependencies"]["native"] == [
        "xlsbkernel>=0.1.0,<0.2"
    ]
    assert "abi3-py311" in cargo["dependencies"]["pyo3"]["features"]


def test_native_wheel_workflow_covers_supported_matrix_and_bundle_install() -> None:
    path = _ROOT / ".github/workflows/xlsbkernel-wheels.yml"
    workflow = yaml.safe_load(path.read_text(encoding="utf-8"))

    for job_name in ("linux", "windows"):
        versions = workflow["jobs"][job_name]["strategy"]["matrix"][
            "python-version"
        ]
        assert versions == ["3.11", "3.12"]
    text = path.read_text(encoding="utf-8")
    assert text.count("-cp311-abi3-") == 2
    assert text.count("--no-index --no-deps") == 2
    assert text.count("import qc_tool, xlsbkernel") == 2
    assert "pull_request:" in text and "push:" in text
