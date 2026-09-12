"""Synthetic artifact and resolver tests for cadence-diff-native releases."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from scripts.audit_native_distributions import main

VERSION = "2.0.0"
LINUX_TAG = "cp311-abi3-manylinux2014_x86_64"
WINDOWS_TAG = "cp311-abi3-win_amd64"
FALLBACK_TAG = "py3-none-any"
CLASSIFIERS = (
    "Programming Language :: Python :: 3.11",
    "Programming Language :: Python :: 3.12",
    "Programming Language :: Rust",
    "Operating System :: Microsoft :: Windows :: Windows 10",
    "Operating System :: POSIX :: Linux",
)
PROJECT_URLS = (
    "Homepage, https://github.com/rai-himanshu07/cadence-diff",
    "Repository, https://github.com/rai-himanshu07/cadence-diff",
)


def _metadata(**overrides: str) -> str:
    values = {
        "Name": "cadence-diff-native",
        "Version": VERSION,
        "Requires-Python": ">=3.11",
        "Summary": "Native Excel analysis accelerator for cadence-diff",
        "License-Expression": "Apache-2.0",
        **overrides,
    }
    lines = ["Metadata-Version: 2.4"]
    lines.extend(f"{name}: {value}" for name, value in values.items())
    lines.extend(f"Classifier: {value}" for value in CLASSIFIERS)
    lines.extend(f"Project-URL: {value}" for value in PROJECT_URLS)
    return "\n".join(lines) + "\n"


def _write_wheel(
    directory: Path,
    tag: str,
    *,
    metadata_overrides: dict[str, str] | None = None,
    internal_tag: str | None = None,
    include_extension: bool | None = None,
    include_fallback: bool | None = None,
    fallback_source: str | None = None,
    extra_members: dict[str, bytes | str] | None = None,
) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"cadence_diff_native-{VERSION}-{tag}.whl"
    native = tag != FALLBACK_TAG
    include_extension = native if include_extension is None else include_extension
    include_fallback = not native if include_fallback is None else include_fallback
    dist_info = f"cadence_diff_native-{VERSION}.dist-info"
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr(
            f"{dist_info}/METADATA",
            _metadata(**(metadata_overrides or {})),
        )
        archive.writestr(
            f"{dist_info}/WHEEL",
            "\n".join(
                (
                    "Wheel-Version: 1.0",
                    "Generator: synthetic-test",
                    f"Root-Is-Purelib: {'false' if native else 'true'}",
                    f"Tag: {internal_tag or tag}",
                    "",
                )
            ),
        )
        if include_extension:
            suffix = ".pyd" if "win_amd64" in tag else ".abi3.so"
            archive.writestr(f"cadence_diff_native/cadence_diff_native{suffix}", b"native")
        if include_fallback:
            archive.writestr(
                "cadence_diff_native/__init__.py",
                fallback_source
                or (
                    '__version__ = "2.0.0"\n'
                    "__kernel_api_version__ = 1\n"
                    "__native_available__ = False\n"
                ),
            )
        for name, payload in (extra_members or {}).items():
            archive.writestr(name, payload)
        archive.writestr(f"{dist_info}/RECORD", "")
    return path


def _good_set(directory: Path) -> None:
    _write_wheel(directory, LINUX_TAG)
    _write_wheel(directory, WINDOWS_TAG)
    _write_wheel(directory, FALLBACK_TAG)


def test_good_three_wheel_set_passes(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    _good_set(tmp_path)

    assert main(["audit_native_distributions.py", str(tmp_path)]) == 0
    assert "three wheel-only artifacts" in capsys.readouterr().out


def test_native_auditor_runs_as_a_direct_script(tmp_path: Path) -> None:
    _good_set(tmp_path)
    completed = subprocess.run(
        [
            sys.executable,
            "scripts/audit_native_distributions.py",
            str(tmp_path),
        ],
        cwd=Path(__file__).parents[1],
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stdout + completed.stderr
    assert "three wheel-only artifacts" in completed.stdout
    assert completed.stderr == ""


@pytest.mark.parametrize("mutation", ["missing", "extra", "sdist"])
def test_artifact_set_must_be_exact(tmp_path: Path, mutation: str) -> None:
    _good_set(tmp_path)
    if mutation == "missing":
        next(tmp_path.glob(f"*-{WINDOWS_TAG}.whl")).unlink()
    elif mutation == "extra":
        _write_wheel(tmp_path, "cp311-abi3-macosx_11_0_arm64")
    else:
        (tmp_path / f"cadence_diff_native-{VERSION}.tar.gz").write_bytes(b"sdist")

    assert main(["audit_native_distributions.py", str(tmp_path)]) == 1


@pytest.mark.parametrize(
    ("overrides", "internal_tag"),
    [
        ({"Name": "different-name"}, None),
        ({"Version": "2.0.1"}, None),
        ({"Requires-Python": ">=3.12"}, None),
        ({"Summary": "different"}, None),
        ({"License-Expression": "MIT"}, None),
        ({}, "cp311-abi3-manylinux_2_28_x86_64"),
    ],
)
def test_metadata_and_internal_tag_must_match(
    tmp_path: Path,
    overrides: dict[str, str],
    internal_tag: str | None,
) -> None:
    _write_wheel(
        tmp_path,
        LINUX_TAG,
        metadata_overrides=overrides,
        internal_tag=internal_tag,
    )
    _write_wheel(tmp_path, WINDOWS_TAG)
    _write_wheel(tmp_path, FALLBACK_TAG)

    assert main(["audit_native_distributions.py", str(tmp_path)]) == 1


@pytest.mark.parametrize(
    ("tag", "include_extension", "include_fallback"),
    [
        (LINUX_TAG, False, True),
        (WINDOWS_TAG, False, False),
        (FALLBACK_TAG, True, True),
        (FALLBACK_TAG, False, False),
    ],
)
def test_native_and_fallback_payloads_cannot_be_inverted(
    tmp_path: Path,
    tag: str,
    include_extension: bool,
    include_fallback: bool,
) -> None:
    _good_set(tmp_path)
    (tmp_path / f"cadence_diff_native-{VERSION}-{tag}.whl").unlink()
    _write_wheel(
        tmp_path,
        tag,
        include_extension=include_extension,
        include_fallback=include_fallback,
    )

    assert main(["audit_native_distributions.py", str(tmp_path)]) == 1


def test_manylinux_pep600_dual_tag_is_accepted(tmp_path: Path) -> None:
    _write_wheel(
        tmp_path,
        "cp311-abi3-manylinux_2_17_x86_64.manylinux2014_x86_64",
    )
    _write_wheel(tmp_path, WINDOWS_TAG)
    _write_wheel(tmp_path, FALLBACK_TAG)

    assert main(["audit_native_distributions.py", str(tmp_path)]) == 0


def test_fallback_status_contract_is_required(tmp_path: Path) -> None:
    _write_wheel(tmp_path, LINUX_TAG)
    _write_wheel(tmp_path, WINDOWS_TAG)
    _write_wheel(
        tmp_path,
        FALLBACK_TAG,
        fallback_source='__version__ = "2.0.0"\n__native_available__ = True\n',
    )

    assert main(["audit_native_distributions.py", str(tmp_path)]) == 1


@pytest.mark.parametrize(
    ("member", "payload"),
    [
        ("data/history.sqlite3", b"private"),
        ("cadence_diff_native/secret.pem", b"private"),
        (
            "cadence_diff_native/note.txt",
            "-----BEGIN " + "PRIVATE KEY-----\nnot-a-real-key",
        ),
        (
            "cadence_diff_native/probe.txt",
            bytes.fromhex("54584f").decode("ascii"),
        ),
        ("../escape.txt", b"unsafe"),
    ],
)
def test_forbidden_paths_and_content_fail(
    tmp_path: Path,
    member: str,
    payload: bytes | str,
) -> None:
    _write_wheel(tmp_path, LINUX_TAG, extra_members={member: payload})
    _write_wheel(tmp_path, WINDOWS_TAG)
    _write_wheel(tmp_path, FALLBACK_TAG)

    assert main(["audit_native_distributions.py", str(tmp_path)]) == 1


@pytest.mark.parametrize(
    ("platform", "python_version", "implementation", "abi", "expected_tag"),
    [
        ("manylinux2014_x86_64", "3.11", "cp", "abi3", LINUX_TAG),
        ("manylinux2014_x86_64", "3.12", "cp", "abi3", LINUX_TAG),
        ("win_amd64", "3.11", "cp", "abi3", WINDOWS_TAG),
        ("win_amd64", "3.12", "cp", "abi3", WINDOWS_TAG),
        ("musllinux_1_2_x86_64", "3.11", "cp", "cp311", FALLBACK_TAG),
        ("macosx_11_0_arm64", "3.12", "cp", "cp312", FALLBACK_TAG),
        ("manylinux2014_aarch64", "3.11", "cp", "cp311", FALLBACK_TAG),
        ("manylinux2014_x86_64", "3.11", "pp", "none", FALLBACK_TAG),
    ],
)
def test_pip_selects_native_or_fallback_without_an_index(
    tmp_path: Path,
    platform: str,
    python_version: str,
    implementation: str,
    abi: str,
    expected_tag: str,
) -> None:
    wheelhouse = tmp_path / "wheelhouse"
    destination = tmp_path / "downloads"
    _good_set(wheelhouse)
    destination.mkdir()
    completed = subprocess.run(
        [
            sys.executable,
            "-m",
            "pip",
            "download",
            "--quiet",
            "--no-index",
            "--no-deps",
            "--only-binary=:all:",
            "--find-links",
            str(wheelhouse),
            "--dest",
            str(destination),
            "--platform",
            platform,
            "--python-version",
            python_version,
            "--implementation",
            implementation,
            "--abi",
            abi,
            f"cadence-diff-native=={VERSION}",
        ],
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
    downloads = list(destination.glob("*.whl"))
    assert len(downloads) == 1
    assert downloads[0].name.endswith(f"-{expected_tag}.whl")
