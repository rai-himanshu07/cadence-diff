# Changelog

All notable changes to `cadence-diff` are documented in this file. The format
loosely follows [Keep a Changelog](https://keepachangelog.com/), and this
project uses [Semantic Versioning](https://semver.org/).

## [Unreleased]

## [2.0.0] - 2026-09-18

### Added

- A unified **Configure & Run** workspace for every browser-started run,
  including progressive Office-free setup analysis, reconnectable private
  sidecar previews, mode-aware file roles, one active member/sheet/table
  editor, explicit scenario checks, and frozen per-run resolved configuration.
- A task-first four-tab table editor (**Rows**, **Data bounds**, **Scenario
  checks**, and **Advanced**) with effective two-sided data-range disclosure,
  reversible table removal, next-needs-attention navigation, bounded OOXML
  formula reveal, and explicit XLSB presence-only behavior.
- Automatic Rust/PyO3 native engine (`native/cadence_diff_native/`) for XLSB
  value/formula decoding and high-volume formula-delta classification. The
  separately built `cadence-diff-native` implementation distribution is an
  exact dependency of `cadence-diff`: pip selects a precompiled stable-ABI
  wheel on certified targets and a universal fallback wheel elsewhere, without
  requiring a local Rust toolchain.
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
  value-free `RankedTableEvidence` run-action contract (v2). A blocked attempt
  returns bounded, package-member-qualified proposals to the same Configure &
  Run workspace, where the analyst explicitly chooses keyed, positional, or
  excluded handling before submission.
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
- The standalone `cadence-diff-native` wheel now targets CPython's 3.11 stable
  ABI and installs automatically with `cadence-diff`; the legacy `[native]`
  extra remains a compatibility alias. Its formula-delta
  classifier uses a 16 MiB batch cap and per-string pre-screen with exact
  per-row Python fallback. A same-tree real Windows A/B reduced formula
  comparison by 51.34% and total time by 32.13%, clearing its ship gates.
- Parent-supervised Windows acceptance with a hard wall-clock deadline,
  periodic process-tree RSS sampling, Office-process accounting, complete
  aggregate phase/service/formula/population telemetry, and source rehashing.

### Changed

- Explicit table modes now drive the comparison engine: positional mode forces
  positional row alignment and bypasses ranked-table screening, while excluded
  tables are removed from matched and unpaired value/formula/structure analysis
  with an auditable coverage disclosure. Active runs also suppress historical
  recovery controls, whose callbacks revalidate persisted eligibility.
- Native XLSB formula rendering now fails closed per cell: a render
  failure is absent from the formula surface (and lowers formula-text
  coverage) rather than ever appearing as an empty string or a bare `=`.
- The resolved formula engine and adapter fingerprint are recorded on
  every run and disclosed across Re-QC, carry-forward, and attestation, so
  automatic engine selection never causes an undisclosed evidence change
  between two machines with different native-engine availability.
- The resolved cached-values decoder is also recorded per workbook role and
  preserved through history, stored-run reconstruction, carry-forward,
  reports, JSON schemas, sign-off, and signed attestation validation.
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
- Population candidate templates now stop at an explicit 4,096-key cap;
  finalization streams capped group state to lazy result blocks, reconstructing
  full atomic findings only for groups that must replay.
- Multi-workbook ranked-table checks collect all currently blocked package
  members into one bounded action, so one return-to-setup pass can surface
  every known row-identity decision before Re-QC.

### Fixed

- Configure & Run now preserves files, scope, profile, selectors, ranges, and
  table choices across Back, reload, Resume, and blocked-run recovery. Returning
  to setup consumes the paused no-run attempt, prevents competing recovery
  controls, and never restarts QC automatically.
- Spawned workers now validate and forward the exact resolved input
  configuration and canonical digest consumed by the authoritative engine;
  warning acknowledgements and run-only safety overrides also survive reload.
- Ranked-candidate detection is persisted separately from transient
  confirmation. A detected ranked table left on Automatic cannot pass bulk
  confirmation or legacy-session restore, while ordinary non-ranked tables may
  return to Automatic without a false blocker.
- The profile editor now materializes omitted optional arrays and missing
  dictionary ancestors transactionally, including population classes, input
  contract members, and comparison prerequisites, instead of raising
  `KeyError` from a normal named-profile Add action.
- Review and queue timers survive transient Socket.IO reconnects. Review-time
  ownership follows every live run view across multiple tabs and dynamic result
  replacement; deleting the final view pauses timing and cancels its timer.
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
- Re-QC also suppresses a delta when the requested mode is unchanged but the
  effective population policy or observed population/atomic representation is
  incompatible.
- Default population classes now have canonical ordering, and resolved output
  policies are independent deeply-frozen snapshots, eliminating hash-seed
  profile drift and mutation through aliased nested models.

## [1.2.0]

Scale whole-column dependency analysis with grouped rectangle storage,
regime/delta closures, and parts-based impact annotation. See git history
for details predating this file.
