"""Release contracts for the stable-ABI cadence-diff native wheels."""

from __future__ import annotations

import tomllib
from pathlib import Path
from typing import Any, cast

import yaml

_ROOT = Path(__file__).parents[1]
_NATIVE_ROOT = _ROOT / "native" / "cadence_diff_native"


class _UniqueKeyLoader(yaml.SafeLoader):
    pass


def _construct_unique_mapping(
    loader: _UniqueKeyLoader,
    node: yaml.MappingNode,
    deep: bool = False,
) -> dict[object, object]:
    mapping: dict[object, object] = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        if key in mapping:
            raise ValueError(f"duplicate YAML key: {key}")
        mapping[key] = loader.construct_object(value_node, deep=deep)
    return mapping


_UniqueKeyLoader.add_constructor(
    yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
    _construct_unique_mapping,
)


def _workflow(path: Path) -> dict[str, Any]:
    payload = yaml.load(path.read_text(encoding="utf-8"), Loader=_UniqueKeyLoader)
    assert isinstance(payload, dict)
    return cast(dict[str, Any], payload)


def test_main_package_automatically_installs_exact_native_helper() -> None:
    project = tomllib.loads((_ROOT / "pyproject.toml").read_text(encoding="utf-8"))

    assert project["project"]["requires-python"] == ">=3.11"
    assert "cadence-diff-native==2.0.0" in project["project"]["dependencies"]
    assert project["project"]["optional-dependencies"]["native"] == [
        "cadence-diff-native==2.0.0"
    ]


def test_native_distribution_and_fallback_share_release_contract() -> None:
    cargo = tomllib.loads(
        (_NATIVE_ROOT / "Cargo.toml").read_text(encoding="utf-8")
    )
    native_project = tomllib.loads(
        (_NATIVE_ROOT / "pyproject.toml").read_text(encoding="utf-8")
    )
    fallback_project = tomllib.loads(
        (_NATIVE_ROOT / "fallback" / "pyproject.toml").read_text(encoding="utf-8")
    )

    assert cargo["package"]["name"] == "cadence-diff-native"
    assert cargo["package"]["version"] == "2.0.0"
    assert cargo["lib"]["name"] == "cadence_diff_native"
    assert "abi3-py311" in cargo["dependencies"]["pyo3"]["features"]
    assert native_project["project"]["name"] == "cadence-diff-native"
    assert native_project["project"]["requires-python"] == ">=3.11"
    assert fallback_project["project"]["name"] == "cadence-diff-native"
    assert fallback_project["project"]["version"] == "2.0.0"
    assert fallback_project["project"]["requires-python"] == ">=3.11"


def test_native_wheel_workflow_builds_three_wheels_without_sdist() -> None:
    path = _ROOT / ".github/workflows/native-wheels.yml"
    text = path.read_text(encoding="utf-8")
    workflow = _workflow(path)
    jobs = workflow["jobs"]
    assert isinstance(jobs, dict)
    assert "manylinux2014_x86_64" in text
    assert "win_amd64" in text
    assert "py3-none-any" in text
    assert "python-version: ['3.11', '3.12']" in text
    assert "cadence_diff_native" in text
    assert "sdist" not in text.casefold()
    assert "audit_native_distributions.py helper-dist" in text
    assert text.count("--only-binary=cadence-diff-native") == 3
    assert text.count("python -m build --wheel --outdir main-dist") == 3
    assert "pull_request:" in text and "push:" in text
    assert set(jobs) == {
        "build-linux",
        "build-windows",
        "build-fallback",
        "audit-artifacts",
        "smoke-linux",
        "smoke-windows",
        "smoke-fallback",
    }


def test_publish_workflow_orders_helper_main_and_github_release() -> None:
    path = _ROOT / ".github/workflows/publish.yml"
    text = path.read_text(encoding="utf-8")
    workflow = _workflow(path)
    jobs = workflow["jobs"]
    assert isinstance(jobs, dict)

    assert "workflow_dispatch:" in text
    assert "types: [published]" not in text
    assert jobs["publish-native"]["environment"] == "pypi-native"
    assert jobs["publish-main"]["environment"] == "pypi"
    assert "verify-native" in jobs["publish-main"]["needs"]
    assert jobs["smoke-main-fallback"]["needs"] == "verify-main"
    assert jobs["smoke-main-fallback"]["runs-on"] == "macos-14"
    assert text.count("--only-binary=:all: cadence-diff==2.0.0") == 2
    assert "NativeKernelStatus.FALLBACK_STUB" in text
    assert "verify-main" in jobs["github-release"]["needs"]
    assert "smoke-main-fallback" in jobs["github-release"]["needs"]
    assert text.count("pypa/gh-action-pypi-publish@") == 2
    assert text.count("audit_native_distributions.py") == 2
    assert text.count("--dest published-helper") == 3
    assert "sha256sum --check ../helper-release/SHA256SUMS" in text
    assert "gh release create" in text
    assert "skip-existing" not in text
