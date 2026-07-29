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

The structural engine models Excel tables and structured references, mixed
cadence bands, every plot and named series in combo charts, chart axes/legends/
labels/geometry, data validation, conditional formatting, bounded and symbolic
formula dependencies, and explicit future-blank availability rules. PowerPoint
uses collision-safe shape identities, semantic table/chart/plot/series matching,
separate speaker notes, and visible chart-label anchors for Excel reconciliation.
Unsupported representations degrade the relevant check instead of becoming a
false pass.

## Install & run

```bash
pip install cadence-diff        # not yet published — see Development
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
qc-tool run --baseline-excel last.xlsx --current-excel this.xlsx \
            --profile monthly --json findings.json
qc-tool run --current-excel this.xlsx            # preflight (inferred)
qc-tool run --current-excel x.xlsx --current-ppt d.pptx   # package QC

# Local numeric scrambling only — NOT privacy-safe or shareable
qc-tool sanitize client_pack.xlsx --seed 7

# Strict, fail-closed redaction + recursive OOXML privacy verification
qc-tool sanitize client_pack.xlsx --seed 7 --redact-text \
  --output client_pack.sanitized.xlsx
qc-tool verify-sanitized client_pack.sanitized.xlsx \
  --forbid "Client Name" --json privacy-report.json

# Structural-only evidence: no values, text, formulas, paths, or identifiers
qc-tool fingerprint client_pack.xlsx --output client_pack.fingerprint.json

# Sanitize an Excel/PPT package together and reverify confirmed mappings
qc-tool sanitize-package --excel current.xlsx --ppt current.pptx \
  --profile monthly --output-dir sanitized-package --forbid "Client Name"

# Signed evidence bundle; verify every member and manifest signature
qc-tool run --current-excel this.xlsx --fail-on never \
  --attestation run.qca
qc-tool verify-attestation run.qca

# Validate a profile — optionally against the real files it targets
qc-tool lint monthly.yaml --against-excel this.xlsx --against-ppt d.pptx
```

Headless runs record into the same history the web UI shows.

JSON exports omit raw cell neighborhoods and mapping-candidate values by
default. `--json-context` includes them for private diagnostics and must not be
treated as a shareable artifact. For encrypted files, prefer
`--password-env`, a mode-600 `--password-file`, or `--password-prompt`; inline
`--password ROLE=PW` is retained for compatibility but exposes secrets in shell
history and process listings.

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
conda run -n cadence-diff-dev pytest -x -q          # 380+ tests, fixture-driven
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
  server_config.py fail-safe local/temporary-LAN configuration
  ui/          NiceGUI app (loopback only) + theme
               packaged operator guide at /guide
  engine.py    run orchestration
docs/QC_MODES.md   operator guide
```

## Publishing (maintainers)

The package carries the `Private :: Do Not Upload` classifier while under
development — PyPI rejects it, so accidental uploads fail. To release:
remove that classifier, bump `qc_tool.__version__`, then
`python -m build && twine check dist/*` and upload deliberately.

## License

Apache-2.0 — see [LICENSE](LICENSE).
