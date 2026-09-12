"""Exact PyPI release visibility contracts."""

from __future__ import annotations

from pathlib import Path

import pytest

from scripts.verify_pypi_release import load_expected, verify_payload


def _payload(files: dict[str, str]) -> dict[str, object]:
    return {
        "urls": [
            {"filename": filename, "digests": {"sha256": digest}}
            for filename, digest in files.items()
        ]
    }


def test_exact_release_files_and_hashes_pass() -> None:
    expected = {"example-2.0.0.whl": "a" * 64}

    assert verify_payload(_payload(expected), expected) == (
        True,
        "exact release files and hashes are visible",
    )


@pytest.mark.parametrize(
    ("observed", "message"),
    [
        ({}, "missing expected files"),
        (
            {"example-2.0.0.whl": "a" * 64, "extra.whl": "b" * 64},
            "unexpected release files",
        ),
        ({"example-2.0.0.whl": "b" * 64}, "SHA-256 mismatch"),
    ],
)
def test_partial_extra_or_mismatched_release_fails(
    observed: dict[str, str],
    message: str,
) -> None:
    passed, detail = verify_payload(
        _payload(observed),
        {"example-2.0.0.whl": "a" * 64},
    )

    assert passed is False
    assert message in detail


def test_checksum_manifest_is_strict(tmp_path: Path) -> None:
    path = tmp_path / "SHA256SUMS"
    path.write_text(f"{'a' * 64}  example-2.0.0.whl\n", encoding="utf-8")
    assert load_expected(path) == {"example-2.0.0.whl": "a" * 64}

    path.write_text("not-a-digest  example.whl\n", encoding="utf-8")
    with pytest.raises(ValueError, match="invalid row"):
        load_expected(path)
