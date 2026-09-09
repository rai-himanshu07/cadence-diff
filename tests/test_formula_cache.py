"""Formula-extraction cache: key derivation, safety, and transparency."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from qc_tool.findings_store import FindingsStoreError, encode_block
from qc_tool.io.formula_cache import (
    CACHE_SCHEMA_VERSION,
    EXTRACTION_CONTRACT_VERSION,
    FormulaCacheKey,
    FormulaExtractionCache,
    coordinate_digest,
    extract_with_cache,
    package_digest,
)
from qc_tool.io.formula_enrichment import ExtractedDefinedName, FormulaExtraction
from qc_tool.io.xlsb_formula import XlsbFormulaScan


def _scan() -> XlsbFormulaScan:
    return XlsbFormulaScan(formula_cells={"Data": frozenset({(1, 1), (2, 2)})})


def _extraction() -> FormulaExtraction:
    return FormulaExtraction(
        formulas={"Data": {(1, 1): "=A2+1", (2, 2): "=SUM(A1:A5)"}},
        engine="test-engine:1.0",
        detail="test extraction",
        defined_names=(ExtractedDefinedName(name="MyRange", target="Data!$A$1:$A$5"),),
        defined_names_complete=True,
    )


def _key(**overrides: object) -> FormulaCacheKey:
    defaults: dict[str, object] = {
        "package_sha256": "a" * 64,
        "coordinate_digest": coordinate_digest(_scan()),
        "coordinate_count": _scan().formula_count,
        "adapter_family": "libreoffice",
        "adapter_fingerprint": "libreoffice:LibreOffice 24.2.4.2",
    }
    defaults.update(overrides)
    return FormulaCacheKey(**defaults)  # type: ignore[arg-type]


def test_key_digest_is_stable() -> None:
    assert _key().digest() == _key().digest()


@pytest.mark.parametrize(
    "override",
    [
        {"package_sha256": "b" * 64},
        {"coordinate_digest": "different"},
        {"coordinate_count": 999},
        {"adapter_family": "excel"},
        {"adapter_fingerprint": "libreoffice:different"},
        {"schema_version": CACHE_SCHEMA_VERSION + 1},
        {"contract_version": EXTRACTION_CONTRACT_VERSION + 1},
    ],
)
def test_key_digest_is_sensitive_to_every_field(override: dict[str, object]) -> None:
    assert _key().digest() != _key(**override).digest()


def test_coordinate_digest_is_order_independent() -> None:
    forward = XlsbFormulaScan(formula_cells={"Data": frozenset({(1, 1), (2, 2)})})
    backward = XlsbFormulaScan(
        formula_cells={"Data": frozenset({(2, 2), (1, 1)})}
    )
    assert coordinate_digest(forward) == coordinate_digest(backward)


def test_coordinate_digest_never_depends_on_formula_text() -> None:
    """The digest is purely coordinate-shaped -- confirmed by construction:
    `coordinate_digest` never receives formula text as an argument."""
    scan_a = XlsbFormulaScan(formula_cells={"Data": frozenset({(1, 1)})})
    scan_b = XlsbFormulaScan(formula_cells={"Data": frozenset({(1, 1)})})
    assert coordinate_digest(scan_a) == coordinate_digest(scan_b)


def test_store_then_lookup_round_trips_extraction(tmp_path: Path) -> None:
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    key = _key()
    extraction = _extraction()

    cache.store(key, extraction)
    result = cache.lookup(key)

    assert result is not None
    assert result.formulas == extraction.formulas
    assert result.engine == extraction.engine
    assert result.detail == extraction.detail
    assert result.defined_names == extraction.defined_names
    assert result.defined_names_complete == extraction.defined_names_complete


def test_lookup_miss_returns_none_for_unknown_key(tmp_path: Path) -> None:
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    assert cache.lookup(_key()) is None


# --- schema v2: formulas_r1c1 round trip -------------------------------------


def test_round_trip_preserves_formulas_r1c1_distinctly_from_a1(tmp_path: Path) -> None:
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    key = _key()
    extraction = FormulaExtraction(
        formulas={"Data": {(1, 1): "=A2+1", (2, 2): "=SUM(A1:A5)"}},
        engine="native-biff12:1.0.0",
        detail="native kernel",
        formulas_r1c1={"Data": {(1, 1): "=R[1]C", (2, 2): "=SUM(R[-1]C:R[-5]C)"}},
    )

    cache.store(key, extraction)
    result = cache.lookup(key)

    assert result is not None
    assert result.formulas == extraction.formulas
    assert result.formulas_r1c1 == extraction.formulas_r1c1


def test_round_trip_keeps_formulas_r1c1_none_when_the_adapter_never_produced_it(
    tmp_path: Path,
) -> None:
    """Excel/LibreOffice extractions leave `formulas_r1c1=None`; the cache
    must not invent an empty dict for them.
    """
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    key = _key()

    cache.store(key, _extraction())  # _extraction() never sets formulas_r1c1
    result = cache.lookup(key)

    assert result is not None
    assert result.formulas_r1c1 is None


def test_round_trip_keeps_formulas_r1c1_empty_dict_distinct_from_none(
    tmp_path: Path,
) -> None:
    """A native extraction that produced zero r1c1 rows (`{}`) must round
    trip as `{}`, not be conflated with an adapter that never supports r1c1
    at all (`None`).
    """
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    key = _key()
    extraction = FormulaExtraction(
        formulas={"Data": {(1, 1): "=A2+1"}},
        engine="native-biff12:1.0.0",
        detail="native kernel",
        formulas_r1c1={},
    )

    cache.store(key, extraction)
    result = cache.lookup(key)

    assert result is not None
    assert result.formulas_r1c1 == {}
    assert result.formulas_r1c1 is not None


def test_entry_is_private_file_under_private_directory(tmp_path: Path) -> None:
    root = tmp_path / "formula-cache"
    cache = FormulaExtractionCache(root)
    cache.store(_key(), _extraction())

    if os.name == "posix":
        assert (root.stat().st_mode & 0o777) == 0o700
        entries = list(root.glob("*.qfxc"))
        assert len(entries) == 1
        assert (entries[0].stat().st_mode & 0o777) == 0o600


def test_corrupt_entry_is_quarantined_and_treated_as_miss(tmp_path: Path) -> None:
    root = tmp_path / "formula-cache"
    cache = FormulaExtractionCache(root)
    key = _key()
    cache.store(key, _extraction())
    entry_path = cache._entry_path(key)
    entry_path.write_bytes(b"not a real cache entry at all")

    assert cache.lookup(key) is None
    assert not entry_path.exists()  # quarantined


def test_r1c1_coordinates_outside_a1_coordinates_are_quarantined(tmp_path: Path) -> None:
    """A stored entry whose r1c1 coordinates are not a subset of its own a1
    coordinates (a tampered file, or a genuinely broken adapter) must never
    be trusted -- it is quarantined and treated as a miss, exactly like any
    other corruption.
    """
    root = tmp_path / "formula-cache"
    cache = FormulaExtractionCache(root)
    key = _key()
    extraction = FormulaExtraction(
        formulas={"Data": {(1, 1): "=A2+1"}},
        engine="native-biff12:1.0.0",
        detail="native kernel",
        formulas_r1c1={"Data": {(1, 1): "=R[1]C", (9, 9): "=unexpected"}},
    )
    cache.store(key, extraction)
    entry_path = cache._entry_path(key)

    assert cache.lookup(key) is None
    assert not entry_path.exists()  # quarantined


def test_lookup_never_follows_a_symlink(tmp_path: Path) -> None:
    root = tmp_path / "formula-cache"
    root.mkdir(parents=True)
    outside = tmp_path / "outside-secret.txt"
    outside.write_text("must never be touched")
    key = _key()
    cache = FormulaExtractionCache(root)
    entry_path = cache._entry_path(key)
    entry_path.symlink_to(outside)

    assert cache.lookup(key) is None
    assert outside.read_text() == "must never be touched"  # untouched


def test_over_cap_entry_is_not_stored(tmp_path: Path) -> None:
    cache = FormulaExtractionCache(tmp_path / "formula-cache", max_entry_bytes=1)
    cache.store(_key(), _extraction())
    assert cache.status() == {"entry_count": 0, "total_bytes": 0}


def test_entry_count_cap_evicts_oldest_first(tmp_path: Path) -> None:
    cache = FormulaExtractionCache(tmp_path / "formula-cache", max_entries=2)
    keys = [_key(package_sha256=str(i) * 64) for i in range(3)]
    for key in keys:
        cache.store(key, _extraction())

    assert cache.status()["entry_count"] == 2
    assert cache.lookup(keys[0]) is None  # oldest evicted
    assert cache.lookup(keys[1]) is not None
    assert cache.lookup(keys[2]) is not None


def test_total_bytes_cap_evicts_oldest_first(tmp_path: Path) -> None:
    one_entry_bytes = None
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    key0 = _key(package_sha256="0" * 64)
    cache.store(key0, _extraction())
    one_entry_bytes = cache.status()["total_bytes"]
    assert one_entry_bytes > 0

    bounded = FormulaExtractionCache(
        tmp_path / "formula-cache-2",
        max_total_bytes=int(one_entry_bytes * 1.5),
    )
    key1 = _key(package_sha256="1" * 64)
    key2 = _key(package_sha256="2" * 64)
    bounded.store(key1, _extraction())
    bounded.store(key2, _extraction())

    assert bounded.status()["entry_count"] == 1
    assert bounded.lookup(key1) is None
    assert bounded.lookup(key2) is not None


def test_clear_removes_all_entries_and_returns_count(tmp_path: Path) -> None:
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    cache.store(_key(package_sha256="0" * 64), _extraction())
    cache.store(_key(package_sha256="1" * 64), _extraction())

    assert cache.clear() == 2
    assert cache.status() == {"entry_count": 0, "total_bytes": 0}
    assert cache.lookup(_key(package_sha256="0" * 64)) is None


def test_concurrent_identical_writers_are_harmless(tmp_path: Path) -> None:
    """A race between two writers for the same key is harmless: both produce
    byte-identical content, so whichever atomic rename lands last wins."""
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    key = _key()
    extraction = _extraction()

    cache.store(key, extraction)
    cache.store(key, extraction)

    result = cache.lookup(key)
    assert result is not None
    assert result.formulas == extraction.formulas


def test_package_digest_matches_sha256() -> None:
    import hashlib

    data = b"some xlsb bytes"
    assert package_digest(data) == hashlib.sha256(data).hexdigest()


# --- extract_with_cache -----------------------------------------------------


def test_extract_with_cache_disabled_when_cache_is_none() -> None:
    calls = []

    def extractor(data: bytes, scan: XlsbFormulaScan) -> FormulaExtraction:
        calls.append(1)
        return _extraction()

    result = extract_with_cache(
        b"data", _scan(), cache=None, engine="libreoffice", extractor=extractor
    )
    result2 = extract_with_cache(
        b"data", _scan(), cache=None, engine="libreoffice", extractor=extractor
    )

    assert len(calls) == 2  # never cached when disabled
    assert result.formulas == result2.formulas == _extraction().formulas


def test_extract_with_cache_disabled_when_fingerprint_unavailable(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import qc_tool.io.formula_cache as formula_cache_module

    monkeypatch.setattr(
        formula_cache_module, "discover_adapter_fingerprint", lambda engine: None
    )
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    calls = []

    def extractor(data: bytes, scan: XlsbFormulaScan) -> FormulaExtraction:
        calls.append(1)
        return _extraction()

    extract_with_cache(b"data", _scan(), cache=cache, engine="libreoffice", extractor=extractor)
    extract_with_cache(b"data", _scan(), cache=cache, engine="libreoffice", extractor=extractor)

    assert len(calls) == 2  # fingerprint unavailable -- caching stays off
    assert cache.status()["entry_count"] == 0


def test_extract_with_cache_hits_on_second_call(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import qc_tool.io.formula_cache as formula_cache_module

    monkeypatch.setattr(
        formula_cache_module,
        "discover_adapter_fingerprint",
        lambda engine: ("libreoffice", "libreoffice:test-version"),
    )
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    calls = []

    def extractor(data: bytes, scan: XlsbFormulaScan) -> FormulaExtraction:
        calls.append(1)
        return _extraction()

    first = extract_with_cache(
        b"same-bytes", _scan(), cache=cache, engine="libreoffice", extractor=extractor
    )
    second = extract_with_cache(
        b"same-bytes", _scan(), cache=cache, engine="libreoffice", extractor=extractor
    )

    assert len(calls) == 1  # second call was a cache hit
    assert first.formulas == second.formulas == _extraction().formulas


def test_extract_with_cache_revalidates_and_falls_back_on_tampered_hit(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A stored entry whose coordinates disagree with a fresh scan (despite a
    matching key) fails live validation, is quarantined, and a normal
    extraction runs instead -- the QC result never uses the bad entry."""
    import qc_tool.io.formula_cache as formula_cache_module

    monkeypatch.setattr(
        formula_cache_module,
        "discover_adapter_fingerprint",
        lambda engine: ("libreoffice", "libreoffice:test-version"),
    )
    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    scan = _scan()
    key = FormulaCacheKey(
        package_sha256=package_digest(b"same-bytes"),
        coordinate_digest=coordinate_digest(scan),
        coordinate_count=scan.formula_count,
        adapter_family="libreoffice",
        adapter_fingerprint="libreoffice:test-version",
    )
    bad_extraction = FormulaExtraction(
        formulas={"Data": {(1, 1): "=1", (2, 2): "=2", (9, 9): "=unexpected"}},
        engine="test-engine",
        detail="tampered",
    )
    cache.store(key, bad_extraction)
    calls = []

    def extractor(data: bytes, scan: XlsbFormulaScan) -> FormulaExtraction:
        calls.append(1)
        return _extraction()

    result = extract_with_cache(
        b"same-bytes", scan, cache=cache, engine="libreoffice", extractor=extractor
    )

    assert len(calls) == 1  # fell back to a real extraction
    assert result.formulas == _extraction().formulas
    assert cache.lookup(key) is not None  # the fallback re-stored a good entry


def test_manifest_missing_blocks_raises_on_read(tmp_path: Path) -> None:
    root = tmp_path / "formula-cache"
    cache = FormulaExtractionCache(root)
    key = _key()
    cache.store(key, _extraction())
    entry_path = cache._entry_path(key)

    # Truncate to just the magic bytes -- malformed footer.
    entry_path.write_bytes(b"QFXC1\n" + encode_block([[1, 1, "=1"]]))

    assert cache.lookup(key) is None
    assert not entry_path.exists()


def test_decode_block_error_is_treated_as_a_miss(tmp_path: Path) -> None:
    """A block whose bytes fail zlib decompression is a corrupt entry, not a
    crash -- `decode_block` raising `FindingsStoreError` must be caught."""
    with pytest.raises(FindingsStoreError):
        from qc_tool.findings_store import decode_block

        decode_block(b"not zlib data")


# --- CLI surface -------------------------------------------------------------


def test_formula_cache_cli_reports_status(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from qc_tool import cli

    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    cache.store(_key(), _extraction())

    assert cli.main(["formula-cache", "status", "--data-dir", str(tmp_path)]) == 0
    out = capsys.readouterr().out
    assert "formula-cache: 1 entries" in out
    assert "native formula engine:" in out


def test_formula_cache_cli_clears(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from qc_tool import cli

    cache = FormulaExtractionCache(tmp_path / "formula-cache")
    cache.store(_key(), _extraction())

    assert cli.main(["formula-cache", "clear", "--data-dir", str(tmp_path)]) == 0
    assert "cleared 1 entry" in capsys.readouterr().out
    assert cache.status() == {"entry_count": 0, "total_bytes": 0}


def test_formula_cache_cli_status_on_empty_cache(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from qc_tool import cli

    assert cli.main(["formula-cache", "status", "--data-dir", str(tmp_path)]) == 0
    out = capsys.readouterr().out
    assert "formula-cache: 0 entries" in out
    assert "native formula engine:" in out
