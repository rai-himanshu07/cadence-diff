# cadence-diff (QC Tool)

Local, read-only QC for recurring ("cadence") reporting deliverables —
Excel workbooks and PowerPoint decks. Everything runs on your machine,
binds to `127.0.0.1` only, never modifies source files, and uses no LLMs
or external services.

## Three QC modes

| Mode | Inputs | What it answers |
|---|---|---|
| **Current-file preflight** | latest Excel and/or PPT | Is this file internally sound? (error literals, cleared/inconsistent formulas, period sequence, draft tokens, blanks, broken links…) |
| **Cycle comparison** | baseline + current | What changed vs last cycle — with expected cadence growth isolated from real errors (shift-aware formula compare, growth-aware alignment). |
| **Final-package QC** | current Excel + current PPT | Do the deck's figures reconcile to the workbook? (suggest → confirm → persist source mappings; coverage reporting.) |

Every run reports explicit coverage — checked / degraded / unavailable —
so a check that could not run is never silently treated as passed.
Findings support analyst severity overrides and comments; exports
(annotated Excel workbook, self-contained HTML) regenerate from the
reviewed state.

Analyst-facing views collapse compatible, edge-adjacent Excel cell findings
into deterministic **review groups**. Review-item counts answer how many
decisions remain; affected-finding counts preserve the complete cell-level
evidence. Open a group for paged atomic detail or switch to Individual findings.
JSON, attestations, Re-QC identity, coverage counts, and CI failure thresholds
remain atomic and backward-compatible.

The structural engine models Excel tables and structured references, mixed
cadence bands, every plot and named series in combo charts, chart axes/legends/
labels/geometry, data validation, conditional formatting, bounded and symbolic
formula dependencies, and explicit future-blank availability rules. PowerPoint
uses collision-safe shape identities, semantic table/chart/plot/series matching,
separate speaker notes, and visible chart-label anchors for Excel reconciliation.
Unsupported representations degrade the relevant check instead of becoming a
false pass.

Modern Excel spill references (`B2#` / `ANCHORARRAY`) use only declared
array-formula extents from the OOXML anchor. Implicit intersection (`@`) is
resolved only when the host cell makes one result unambiguous. Missing or
unprovable spill metadata degrades the affected reference coverage rather than
guessing; XLSB spill extents remain degraded even when formula text is enriched.

OOXML loading streams worksheet content, verifies physical cells independently
of declared dimensions, and reports workload evidence from uncompressed XML,
shared strings, styles, cell counts, and sheet bounds. Pathological packages
are refused before expensive parsing unless the operator deliberately enables
the per-run override. Findings budgets and low-confidence alignment are always
disclosed through summary findings and degraded coverage rather than silently
truncating or guessing.

## Install & run

```bash
pip install cadence-diff
cadence-diff                    # web UI → http://127.0.0.1:8080
cadence-diff --port 9000 --data-dir ~/qc-data
qc-tool                         # compatibility alias
python -m qc_tool               # equivalent
```

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
cadence-diff run --current-excel this.xlsx --individual-findings # raw CLI rows

# Only after reviewing workload refusal and confirming sufficient local memory
cadence-diff run --current-excel unusually-large.xlsx \
  --allow-large-workbooks

# Local numeric scrambling only — NOT privacy-safe or shareable
cadence-diff sanitize client_pack.xlsx --seed 7

# Strict, fail-closed redaction + recursive OOXML privacy verification
cadence-diff sanitize client_pack.xlsx --seed 7 --redact-text \
  --output client_pack.sanitized.xlsx
cadence-diff verify-sanitized client_pack.sanitized.xlsx \
  --forbid "Client Name" --json privacy-report.json

# Structural-only evidence: no values, text, formulas, paths, or identifiers
cadence-diff fingerprint client_pack.xlsx --output client_pack.fingerprint.json

# Sanitize an Excel/PPT package together and reverify confirmed mappings
cadence-diff sanitize-package --excel current.xlsx --ppt current.pptx \
  --profile monthly --output-dir sanitized-package --forbid "Client Name"

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
only after an independent BIFF12 scan identifies every formula coordinate:

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

Strict sanitization supports `.xlsx` and `.pptx`. Macro-enabled `.xlsm` is
refused because VBA can retain credentials and client identifiers; create and
review a macro-free `.xlsx` copy first. The redaction manifest and privacy
verifier report what was removed, transformed, skipped, or unverifiable.

## Development

The development environment is Conda/Miniforge:

```bash
conda env create -f environment.yml
conda run -n cadence-diff-dev python main.py        # dev server, repo-local ./data
conda run -n cadence-diff-dev pytest -x -q          # 400+ tests, fixture-driven
conda run -n cadence-diff-dev ruff check .
conda run -n cadence-diff-dev pyright
```

Tests are contract-driven: `tests/fixtures/generate.py` builds synthetic
deliverable pairs with a ground-truth defect manifest, and the E2E suite
asserts both directions — every seeded defect detected, and no finding
without a seeded cause.

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
  fingerprint.py structural-only evidence export
  attestation.py signed QC evidence bundles
  progress.py  progress events and cooperative cancellation
  server_config.py fail-safe local/temporary-LAN configuration
  ui/          NiceGUI app (loopback only) + theme
               packaged operator guide at /guide
  engine.py    run orchestration
```

## Publishing (maintainers)

Releases are built and validated by GitHub Actions, then published through PyPI
Trusted Publishing. Create an annotated version tag matching
`qc_tool.__version__`, publish the corresponding GitHub Release, and promote the
exact workflow-built artifacts. Do not upload manually or rebuild between
validation and publication.

## License

Apache-2.0 — see [LICENSE](LICENSE).
