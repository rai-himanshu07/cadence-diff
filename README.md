# cadence-diff (QC Tool)

Local, read-only QC for recurring ("cadence") reporting deliverables —
Excel workbooks and PowerPoint decks. Everything runs on your machine,
binds to `127.0.0.1` only, never modifies source files, and uses no LLMs
or external services.

## Install and first start

```bash
# Works from an activated pip or Conda environment
python -m pip install cadence-diff
python -m qc_tool

# Windows: create a per-user Desktop shortcut for this exact environment
python -m qc_tool shortcut install
```

Open `http://127.0.0.1:8080` if the browser does not open automatically.
On Windows, `python -m qc_tool` is the reliable fallback when the Python
Scripts directory is not in `PATH` and `cadence-diff` is not recognized.

A fresh installation starts in **Current-file preflight**. Upload one current
Excel workbook and/or PowerPoint deck and run with the built-in `default`
profile. No baseline or YAML profile is needed for this first run. QC Tool then
remembers the last mode you select; Re-QC always restores the original run mode.

```bash
# Inspect or remove the Windows shortcut later
python -m qc_tool shortcut status
python -m qc_tool shortcut remove
```

The shortcut uses the environment's absolute Python path, so neither `PATH` nor
Conda activation is needed. Installation never changes the Desktop
automatically.

![QC Tool review queue](https://raw.githubusercontent.com/rai-himanshu07/cadence-diff/main/qc_tool/assets/review-queue.png)

## Three QC modes

| Mode | Inputs | What it answers |
|---|---|---|
| **Current-file preflight** | latest Excel workbook member(s) and/or PPT | Is this package internally sound? (error literals, cleared/inconsistent formulas, period sequence, draft tokens, blanks, broken links…) |
| **Cycle comparison** | baseline + current workbook members and/or PPT pair | What changed vs last cycle — with members paired by stable ID and expected cadence growth isolated from real errors. |
| **Final-package QC** | one to eight current Excel members + current PPT | Do the deck's figures reconcile to the nominated workbook members? (suggest → confirm → persist source mappings; coverage reporting.) |

## First review

1. Read the run outcome and capability status; unavailable or degraded checks
  narrow what a low finding count proves.
2. Select an open review item and compare baseline/current evidence and nearby
  cells.
3. Confirm or change severity and add a specific evidence-based comment.
4. Continue until material items are dispositioned, then export or finalize.

On local Windows, the header settings dialog can enable **Desktop Office
focus** after explicit consent. Open the exact saved workbook or presentation
in desktop Office, then use **Bind → Confirm binding → Focus** on an atomic
finding. AutoSave, ambiguous documents, changed bytes, active content, LAN mode,
and unreadable identity refuse the action. Disabling the setting immediately
clears every focus token and binding. `--desktop-focus` remains a one-launch
compatibility option.

<details>
<summary><strong>Detailed capability reference</strong></summary>

Every run reports explicit coverage — checked / degraded / unavailable —
so a check that could not run is never silently treated as passed.
Findings support analyst severity overrides and comments; exports
(annotated Excel workbook, self-contained HTML) regenerate from the
reviewed state. A reviewed run can be finalized into immutable reports and a
verified signed attestation. Re-QC presents unchanged prior decisions for
explicit evidence-fingerprint-gated carry-forward; nothing is accepted
automatically.

Analyst-facing views partition findings into deterministic **semantic pattern
groups**, the primary analyst decisions. Spatial groups remain a separate
backward-compatible layout metric. Pattern review-item counts answer how many
decisions remain; atomic-finding counts preserve the complete cell-level
evidence. Results open on the review queue with a selected-decision evidence
panel; stories, coverage, and every atomic finding are one click away.
A finding-class multi-select composes with severity and text filtering, while
the disjoint All / Needs review / Reviewed selector keeps partially reviewed
groups in the action queue. While that queue owns focus, `j`/`k` or Arrow keys
navigate visible decisions, `1`–`4` set severity for unreviewed members, and
`c` confirms the current severity before advancing.
JSON, attestations, Re-QC identity, coverage counts, and CI failure thresholds
remain atomic and backward-compatible.

Severity carries analyst judgment, not just class labels. Numeric value
changes carry independent **magnitude** and **temporal context** axes.
Display-identical ULP-scale noise reports as Info; declared acceptance bands
remain visible as within-tolerance Info. Current/recent material changes are
Warning only after hard guards: sign flips, zero-boundary changes, and a
10x-or-greater magnitude ratio remain Critical. Historical material changes
stay Critical. Implicit numeric block refreshes are Warning; only explicit
profile refresh ranges may be Expected. Inherited explicit `NA()` formulas can
be Info, formula-backed systematic data-state errors can be Warning, and
structural, new, changed, or unsupported inherited errors remain Critical.
Systematic formula rollouts are
recognized when the old logic survives as a subtree inside a new wrapper, and
in-place key changes distinguish formula-derived labels from genuine history
rewrites. **Change stories** are a dedicated results view, linking structural
drivers to the formula changes that reference them and isolating noise,
refresh, and inherited populations — everything unexplained lands in an explicit
residual review queue. Stories never alter severities or counts.

The structural engine models Excel tables and structured references, mixed
cadence bands, every plot and named series in combo charts, chart axes/legends/
labels/geometry, data validation, conditional formatting, bounded and symbolic
formula dependencies, and explicit future-blank availability rules. PowerPoint
uses collision-safe shape identities, semantic table/chart/plot/series matching,
separate speaker notes, and visible chart-label anchors for Excel reconciliation.
Unsupported representations degrade the relevant check instead of becoming a
false pass.

Workbook state outside the grid is compared too, each behind its own coverage
row. **Defined names** are compared with explicit scope, so the same name in two
sheets stays two identities and neither shadows the other. **VBA modules** are
inventoried and their source text diffed for XLSM, XLSB, and any XLSX carrying a
project; findings report module name, line counts, changed line ranges, and a
digest, never the macro source, and no macro is ever executed. **Cell comments**,
**Power Query definitions**, and **data connections** are compared as content,
separately from the existing presence/risk detection. Connections are modelled
as a fixed classification plus a target digest — the connection string, URL, and
command text are never read into a finding, report, or log — and nothing is
fetched and no query is run.

The review queue is ordered by deterministic guided priority, scored from
severity, materiality, temporal position, provenance, downstream impacts,
population size, and story membership, with the cited counts shown for each
decision. Ordering is a permutation: it may de-emphasize, never hide. Decisions
already made — waived, reviewed, expected — defer below everything still open.

Named contracts accumulate a longitudinal decision dossier across direct Re-QC
lineage. Missing historical observations are labelled rather than inferred, and
recurrence nudges require three finalized occurrences plus at least two fresh
manual decisions; no policy is applied automatically. A read-only what-if
preview reuses typed numeric evidence to show how temporary acceptance bounds
or a materiality review floor would move atomics and decisions while keeping
accepted findings visible as Info.

Formula-logic findings include a bounded token-level R1C1 diff in the UI and
HTML report. Cycle comparisons persist a factual alignment trust manifest for
every workbook member and region, including paired/skipped/unpaired counts.
A separate compact formula graph detects circular references without expanding
large ranges, and PowerPoint preflight checks semantically identical claims
repeated across slides while treating ambiguous identity as unavailable.

Multi-workbook packages use a versioned, path-free manifest with stable member
IDs and a cap of eight Excel workbooks per side. Duplicate sheet names remain
member-qualified through findings, review groups, mappings, history, reports,
signed evidence, desktop focus, strict sanitization, and structural
fingerprinting. Cross-workbook formulas are inventoried but never followed,
refreshed, or evaluated.

Excel-to-PowerPoint coverage states its population explicitly: claims that were
read, and surfaces that could not be. A figure rendered into a picture or an
embedded object is counted as unavailable and its slides are named, rather than
being dropped from the denominator. Embedded image bytes are also hashed without
decoding, so additions, removals, and byte changes are reported structurally.
Rendered-visual comparison and OCR remain deliberately unavailable: byte
identity does not prove pixel, crop, layout, or text equivalence.

The review queue can also collapse one logical time series that the engine
correctly split across several decisions. A related-series row is a navigation
lens, never a new decision: canonical decisions, group identities, counts,
reports, JSON, carry-forward, and signed evidence are unchanged, and promoting a
series can only increase the number of top-level rows. Parent labels are
structural only — sheet, measure column letter or row number, and how many
periods changed — and `Confirm visible` records each visible, unreviewed finding
at its own current severity while stating exactly what a filter is hiding.
A period populated for the first time this cycle joins its column as a new-period
segment, and a historical cell cleared to blank joins as a cleared-period segment
while sibling measures still prove that period exists; both sort at their period
position so a parent reads in time order. A wiped period row stays together as
one decision rather than being scattered across column parents. Continuity comes
from producer-authored, digest-bound private evidence, so runs recorded before
this feature keep their existing grouping until a verified Re-QC.

Named reporting contracts have typed Core and Advanced Excel/PowerPoint editors
for every profile field, including repeatable controls, mappings, waivers, and
availability rules. Canonical YAML remains available as an advanced view over
the same lossless draft. Saves are statically linted, atomic, and protected by
an exact source-byte concurrency check; optional selected-file validation uses
the same bounded local loaders as QC.

Modern Excel spill references (`B2#` / `ANCHORARRAY`) use only declared
array-formula extents from the OOXML anchor. Implicit intersection (`@`) is
resolved only when the host cell makes one result unambiguous. Missing or
unprovable spill metadata degrades the affected reference coverage rather than
guessing; XLSB spill extents remain degraded even when formula text is enriched.

OOXML loading streams worksheet content, verifies physical cells independently
of declared dimensions, and reports workload evidence from uncompressed XML,
shared strings, styles, cell counts, and sheet bounds. XLSB structural scanning
reports retained cells, formulas, binary worksheet bytes, shared strings,
styles, and sheet extents before cached values are materialized. Pathological
packages are refused before expensive parsing unless the operator deliberately
enables the per-run override. Every finding is stored and reported in full; run
history tracks the storage each run occupies and prompts for cleanup when the
total grows large. Low-confidence alignment is always disclosed through
summary findings and degraded coverage rather than silently guessing.

</details>

## Installation and launch reference

```bash
# Stable release
pip install cadence-diff
cadence-diff                    # web UI → http://127.0.0.1:8080
cadence-diff --port 9000 --data-dir ~/qc-data
qc-tool                         # compatibility alias
python -m qc_tool               # PATH-independent equivalent
python -m qc_tool launch        # start quietly, or reuse an authenticated instance
python -m qc_tool --desktop-focus  # focus enabled for this launch only
```

`1.2.0` is the current stable release. Existing `1.1.0`, `1.0.0`, `0.1.x`,
and `0.2.0a1` artifacts remain immutable; an unqualified install selects the
latest stable version.

The web UI includes a packaged **Guide** page at `/guide`. It covers mode
selection, files, profiles and controls, coverage/severity, finding review,
Excel↔PPT mappings, Re-QC, privacy artifacts, CLI automation, network access,
troubleshooting, and the sign-off checklist. Context links from the run page
jump directly to the relevant guide section.

### Headless / scripting

```bash
# CI-style QC: exit 0 clean, 2 when findings at/above --fail-on exist
cadence-diff run --baseline-excel last.xlsx --current-excel this.xlsx \
                 --profile monthly --json findings.json --progress
cadence-diff run --current-excel this.xlsx            # preflight (inferred)
cadence-diff run --current-excel x.xlsx --current-ppt d.pptx   # package QC
cadence-diff run --current-excel core.xlsx \
  --current-workbook ops=ops.xlsx --current-ppt deck.pptx
cadence-diff run --current-excel this.xlsx --individual-findings # raw CLI rows
cadence-diff run --current-excel this.xlsx --json findings.json \
  --json-review-summary  # v2: pattern groups, stories, alignment trust

# Only after reviewing every workload refusal and confirming sufficient memory
cadence-diff run --current-excel unusually-large.xlsx \
  --allow-large-workbooks

# Optional analyst acceptance threshold: visible Info, never hidden
cadence-diff run --baseline-excel last.xlsx --current-excel this.xlsx \
  --accept-absolute 1 --accept-percent 0.1

# Optional validated scope: files load fully; selected/total counts are reported
cadence-diff run --baseline-excel last.xlsx --current-excel this.xlsx \
  --sheets "Dashboard,Data" --slides 1,3-5

# Local numeric scrambling only — NOT privacy-safe or shareable
cadence-diff sanitize client_pack.xlsx --seed 7

# Strict, fail-closed redaction + recursive OOXML privacy verification
cadence-diff sanitize client_pack.xlsx --seed 7 --redact-text \
  --output client_pack.sanitized.xlsx
cadence-diff verify-sanitized client_pack.sanitized.xlsx \
  --forbid "Client Name" --json privacy-report.json

# Structural-only evidence: no values, text, formulas, paths, or identifiers
cadence-diff fingerprint client_pack.xlsx --output client_pack.fingerprint.json
cadence-diff fingerprint --current-excel core.xlsx \
  --current-workbook ops=ops.xlsx --current-ppt deck.pptx \
  --output package.fingerprint.json

# Sanitize Excel member(s)/PPT together and reverify confirmed mappings
cadence-diff sanitize-package --excel current.xlsx --ppt current.pptx \
  --profile monthly --output-dir sanitized-package --forbid "Client Name"
cadence-diff sanitize-package --excel core.xlsx --workbook ops=ops.xlsx \
  --ppt deck.pptx --profile monthly --output-dir sanitized-package

# Signed evidence bundle; verify every member and manifest signature
cadence-diff run --current-excel this.xlsx --fail-on never \
  --attestation run.qca
cadence-diff verify-attestation run.qca

# Validate a profile — optionally against the real files it targets
cadence-diff lint monthly.yaml --against-excel this.xlsx --against-ppt d.pptx
```

`qc-tool` remains a compatibility alias. Headless runs record into the same
history the web UI shows. `--progress` writes phase updates to stderr, leaving
normal stdout and JSON files unchanged. The web UI shows the same phases and a
Cancel control; cancellation stops at the next safe boundary, removes partial
reports, and does not record a successful run.

Profiles can define legacy additive tie-outs with `components`, or mutually
exclusive signed `terms` using `operation: add|subtract`. Targets and terms may
use bounded A1, workbook named, or supported structured references. Profile
text is never executed as Python or a free-form expression language.

JSON exports omit raw cell neighborhoods and mapping-candidate values by
default. `--json-context` includes them for private diagnostics and must not be
treated as a shareable artifact. For encrypted files, prefer
`--password-env`, a mode-600 `--password-file`, or `--password-prompt`; inline
`--password ROLE=PW` is retained for compatibility but exposes secrets in shell
history and process listings. File-open passwords are supplied per input role,
so baseline and current files may use different passwords. Standard Excel
worksheet/workbook protection and locked, unlocked, or formula-hidden cell
flags control editing and do not prevent read-only QC; those protection settings
are not themselves audited. IRM, sensitivity-label encryption, and filesystem
read restrictions may still prevent access. Use distinct filenames when two
encrypted inputs require different passwords because runtime password lookup is
currently keyed by filename.

### Network exposure

The server defaults to loopback. Temporary network access is explicit,
time-bounded, and **has no built-in authentication or TLS**:

```bash
qc-tool network status
qc-tool network lan --minutes 60
qc-tool --port 8001                 # reads the persisted temporary config
qc-tool network local              # running LAN server shuts down within 2s

# one-process override (also persisted so status/UI agree)
qc-tool --port 8001 --network lan --expose-for 60
```

LAN mode binds `0.0.0.0`, displays an exposure warning in the UI, shuts the
server down at expiry, and resets the persisted config to local. Router port
forwarding and host firewalls are separate; public Internet exposure is not
recommended without an authenticated TLS reverse proxy.

Runs, uploads, profiles, and history live in a per-user data directory
(`--data-dir` to override). Supported inputs: `.xlsx`, `.xlsm`, `.xlsb`
and `.pptx`, including password-protected files.

### XLSB formula enrichment

The original XLSB remains read-only and `pyxlsb` remains authoritative for its
saved values. Formula text is enriched through a platform spreadsheet engine
only after an independent structural scan of the binary file identifies every
formula coordinate:

- **Windows:** a licensed desktop Excel installation in an interactive user
  session. The conditional `pywin32` dependency is installed with the package.
  Excel runs in a bounded worker with links, events, calculation, and VBA
  automation disabled; only `Formula2` text is accepted.
- **Linux:** `libreoffice` and `bwrap` must be installed on the host. Conversion
  runs in a private bubblewrap sandbox with no network and a read-only host
  filesystem; only formulas from the temporary OOXML copy are accepted.

The source and extracted formula-coordinate sets must match exactly. Active
content, external data, missing platform tools, timeouts, malformed packages,
or any mismatch fail closed to **presence-only** formula coverage. That fallback
still checks formula hardcodes, removals, missing fill formulas, and saved error
values, but does not claim logic, consistency, or formula-text error checks.
Temporary decrypted/conversion files are private and removed after each load.
The Windows worker is implemented and protocol-tested, but must still pass the
restricted corporate Windows/Excel pilot before production sign-off.

Strict sanitization supports bounded current `.xlsx` workbook members plus one
`.pptx`. Macro-enabled `.xlsm` is refused because VBA can retain credentials
and client identifiers; create and review a macro-free `.xlsx` copy first.
Multi-member outputs use structural aliases rather than source filenames or
analyst member IDs. The redaction manifest and privacy verifier report what was
removed, transformed, skipped, or unverifiable.

## Development

The development environment is Conda/Miniforge:

```bash
conda env create -f environment.yml
conda run -n cadence-diff-dev python main.py        # dev server, repo-local ./data
conda run -n cadence-diff-dev pytest -x -q          # fixture-driven contract suite
conda run -n cadence-diff-dev ruff check .
conda run -n cadence-diff-dev pyright
```

Tests are contract-driven: `tests/fixtures/generate.py` builds synthetic
deliverable pairs with a ground-truth defect manifest, and the E2E suite
asserts both directions — every seeded defect detected, and no finding
without a seeded cause.

### Optional native XLSB kernel

`native/xlsbkernel/` is an optional Rust/PyO3 accelerator that speeds up XLSB
value and formula decoding. It is its **own, separately versioned package**
(`xlsbkernel`, built with maturin) — it is never a build- or install-time
dependency of `cadence-diff`, which stays on Hatchling and installs and runs
identically whether or not it is present. Build and install it locally with:

```bash
conda run -n cadence-diff-dev maturin build --release -m native/xlsbkernel/Cargo.toml
conda run -n cadence-diff-dev pip install native/xlsbkernel/target/wheels/xlsbkernel-*.whl
```

or, for iterative development, `maturin develop --release` from inside
`native/xlsbkernel/`. Set a profile's `formula_engine` to `native` (or leave
it `auto`, the default, which prefers native when importable) to use it; when
absent, XLSB formula/value handling falls back to the existing Excel-COM
(Windows) or LibreOffice (Linux) adapters exactly as before, with the
fallback disclosed in run coverage.

### Layout

```
qc_tool/
  io/          read-only loading, decryption, XLSB scan/enrichment, snapshots
  excel/       regions, alignment, value/structure/formula QC, deps, preflight
  ppt/         extraction, fuzzy slide matching, diff, preflight
  crosscheck/  figure parsing, source tracing, final-package reconciliation
  triage/      severity rules (analyst overrides layered on top)
  report/      annotated Excel + standalone HTML exports
  history/     SQLite run history, annotations, re-QC lineage
  privacy.py   fail-closed OOXML privacy verification
  fingerprint.py structural-only artifact/package evidence export
  package.py   stable bounded package/member identity
  package_sanitize.py strict member-aware package redaction
  attestation.py signed QC evidence bundles
  progress.py  progress events and cooperative cancellation
  server_config.py fail-safe local/temporary-LAN configuration
  ui/          NiceGUI app (loopback only) + theme
               packaged operator guide at /guide
  engine.py    run orchestration
native/
  xlsbkernel/  optional Rust/PyO3 XLSB decoder accelerator (own package)
```

## Publishing (maintainers)

Releases are built and validated by GitHub Actions, then published through PyPI
Trusted Publishing. Create an annotated version tag matching
`qc_tool.__version__`, publish the corresponding GitHub Release, and promote the
exact workflow-built artifacts. Do not upload manually or rebuild between
validation and publication.

## License

Apache-2.0 — see [LICENSE](LICENSE).
