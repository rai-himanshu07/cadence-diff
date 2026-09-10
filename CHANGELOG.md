# Changelog

All notable changes to `cadence-diff` are documented in this file. The format
loosely follows [Keep a Changelog](https://keepachangelog.com/), and this
project uses [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- Optional native Rust/PyO3 XLSB kernel (`native/xlsbkernel/`) that
  accelerates XLSB value and formula decoding. It is its own, separately
  versioned package built with maturin, never a build- or install-time
  dependency of `cadence-diff`: absent, or on any runtime failure, loading
  falls back to the existing pyxlsb / Excel-COM / LibreOffice paths
  unchanged, with the fallback disclosed in run coverage.
- Group-first population findings: a homogeneous run of same-shape
  formula/format differences is now summarized as one population finding
  with bounded sampled evidence instead of one atomic finding per cell,
  under an explicit versioned policy that leaves every legacy run
  byte-identical. Population member access, carry-forward, and sample
  -excerpt loading are all bounded (streamed paging, a disposable owned
  worker process for excerpts) so the UI server never materializes an
  entire population or holds two full large workbooks itself.
- Formula-cache schema v2, preserving canonical R1C1 formulas and
  engine/detail fields across a private, content-addressed cache round trip.
- Attestation schema v4, composing the existing v2 sign-off and v3
  package/input-role checks cumulatively with new typed population
  -manifest and annotation-lineage validation.
- Ranked/sorted-table detection: a pre-diff heuristic flags an
  unconfigured positional block region that a composite row identity would
  explain more reliably than raw row position. Surfaced through a typed,
  value-free `RankedTableEvidence` run-action contract (v2) and a rebuilt
  "Review row matching" dialog: chip-based column selection, one
  -region-at-a-time navigation, every duplicate-identity policy visible
  with its exact consequence, profile optimistic concurrency, and
  package-member qualification.
- Findings JSON schema v3, selected automatically whenever a report
  contains a population finding; v1/v2 outputs and readers remain
  compatible for atomic/package runs without one.
- CI workflow building and testing the native kernel's wheels.
- A versioned run-level finding-output contract, `FindingOutputMode`
  (`profile` / `decision` / `atomic`), threaded through the run request,
  worker IPC, history, stored-run rehydration, Excel/HTML/JSON reports,
  attestation, Re-QC, and package-member runs, separately from `QCRunMode`
  and without changing `profile_sha256`. The UI exposes a segmented
  "Finding output" control that preselects `decision` (population-grouped,
  compact) for new cycle comparisons; `profile` keeps today's exact
  behavior and `atomic` is the forensic/advanced lane, explicitly outside
  the interactive SLA.

### Changed

- Native XLSB formula rendering now fails closed per cell: a render
  failure is absent from the formula surface (and lowers formula-text
  coverage) rather than ever appearing as an empty string or a bare `=`.
- The resolved formula engine and adapter fingerprint are recorded on
  every run and disclosed across Re-QC, carry-forward, and attestation, so
  automatic engine selection never causes an undisclosed evidence change
  between two machines with different optional kernels installed.
- Re-QC delta accounting uses multiset semantics, so two populations that
  happen to share an identity key can no longer collapse into one.
- Blocked-run CLI and UI messages now name the package member whenever it
  is not `primary`.
- Formula comparison memoizes per-pair wrapper classification across a
  run (bounded, 128 MiB cap) once real-workload telemetry showed compare
  time dominating a compact decision-mode run; population candidate
  storage now uses a delta-encoded, lazily-reconstructed codec instead of
  eagerly building a full finding per member. Both keep exact finding and
  review parity with the prior, uncached behavior.

### Fixed

- The native-kernel Python boundary is now total: mismatched formula
  -surface vector lengths, an out-of-range definition id, or any PyO3
  runtime failure become a disclosed, bounded `FormulaEnrichmentError` and
  fall back to formula-presence-only coverage, instead of an unhandled
  `ValueError`, `IndexError`, or `RuntimeError` from a normal run.
- Re-QC's change-summary banner no longer compares findings naively across
  an output-mode change: `requested_output_mode` differing between the
  compared runs now discloses the representation change explicitly instead
  of showing a resolved/new count computed across two incompatible
  identity spaces (population vs. atomic).

## [1.2.0]

Scale whole-column dependency analysis with grouped rectangle storage,
regime/delta closures, and parts-based impact annotation. See git history
for details predating this file.
