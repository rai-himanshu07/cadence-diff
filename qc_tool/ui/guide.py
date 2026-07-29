# ruff: noqa: E501
"""Packaged in-app operator guide for QC Tool."""

from collections.abc import Iterator
from contextlib import contextmanager
from html import escape

from nicegui import ui

GUIDE_SECTIONS = (
    ("start", "Start here"),
    ("modes", "Choose a mode"),
    ("files", "Files and formats"),
    ("profiles", "Profiles and controls"),
    ("coverage", "Coverage and severity"),
    ("review", "Review findings"),
    ("mappings", "Excel to PowerPoint mappings"),
    ("reqc", "Re-QC and history"),
    ("privacy", "Privacy and sharing"),
    ("cli", "CLI and automation"),
    ("network", "Network access"),
    ("troubleshooting", "Troubleshooting"),
    ("signoff", "Before sign-off"),
)


@contextmanager
def _guide_section(anchor: str, title: str) -> Iterator[None]:
    with ui.element("section").classes("guide-section").props(f'id="{anchor}"'):
        ui.label(title).classes("guide-section-title")
        yield


def _paragraph(text: str) -> None:
    ui.label(text).classes("guide-copy")


def _list(items: list[str]) -> None:
    content = "".join(f"<li>{item}</li>" for item in items)
    ui.html(f'<ul class="guide-list">{content}</ul>')


def _code(content: str) -> None:
    ui.html(f"<pre class=\"guide-code\"><code>{escape(content)}</code></pre>")


def _table(headers: list[str], rows: list[list[str]]) -> None:
    head = "".join(f"<th>{header}</th>" for header in headers)
    body = "".join(
        "<tr>" + "".join(f"<td>{cell}</td>" for cell in row) + "</tr>"
        for row in rows
    )
    ui.html(
        '<div class="guide-table-wrap"><table class="guide-table">'
        f"<thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>"
    )


def _callout(title: str, text: str, *, warning: bool = False) -> None:
    with ui.element("div").classes(
        "guide-callout guide-callout-warning" if warning else "guide-callout"
    ):
        ui.label(title).classes("guide-callout-title")
        ui.label(text).classes("guide-callout-copy")


def render_guide() -> None:
    """Render the complete installed-user guide."""
    ui.label("QC Tool guide").classes("guide-page-title")
    ui.label(
        "Operational guidance for selecting a mode, configuring repeatable checks, "
        "reviewing findings, sharing evidence safely, and automating runs."
    ).classes("guide-page-lede")

    with ui.element("div").classes("guide-layout"):
        with ui.element("aside").classes("guide-toc"):
            ui.label("On this page").classes("guide-toc-title")
            for anchor, title in GUIDE_SECTIONS:
                ui.link(title, f"/guide#{anchor}").classes("guide-toc-link")

        with ui.element("article").classes("guide-content"):
            with _guide_section("start", "Start here"):
                _paragraph(
                    "A defensible run has four parts: select the mode that answers the "
                    "actual QC question, supply only the files that mode needs, review "
                    "what was and was not checked, then disposition every material finding."
                )
                _list(
                    [
                        "Select a <strong>QC mode</strong> before choosing files.",
                        "Use a <strong>named profile</strong> for recurring deliverables and mappings.",
                        "Run QC, then review <strong>coverage before counts</strong>.",
                        "Expand findings for baseline/current evidence, cell context, and impacts.",
                        "Record analyst comments or severity overrides before exporting or attesting.",
                    ]
                )
                _callout(
                    "Sources are read-only",
                    "QC Tool copies uploaded files into managed local storage and verifies source "
                    "integrity in tests. It does not repair or overwrite source deliverables.",
                )

            with _guide_section("modes", "Choose the right QC mode"):
                _table(
                    ["Mode", "Required inputs", "Use it to answer", "Cannot establish"],
                    [
                        [
                            "Current-file preflight",
                            "Latest Excel and/or latest PPT",
                            "Is this file internally sound now?",
                            "Whether historical content changed",
                        ],
                        [
                            "Cycle comparison",
                            "Baseline/current Excel pair and/or PPT pair",
                            "What changed, excluding expected cadence growth?",
                            "Upstream-data correctness or recalculated formulas",
                        ],
                        [
                            "Final-package QC",
                            "Latest Excel plus latest PPT",
                            "Does the final deck reconcile to its workbook?",
                            "Whether either artifact changed from last cycle",
                        ],
                    ],
                )
                _callout(
                    "Do not substitute preflight for comparison",
                    "A standalone run can detect intrinsic defects but cannot prove that a prior "
                    "value, row, formula, slide, or layout was not changed or removed.",
                    warning=True,
                )

            with _guide_section("files", "Files, passwords, and formats"):
                _list(
                    [
                        "Excel: <code>.xlsx</code>, <code>.xlsm</code>, and <code>.xlsb</code> for QC input.",
                        "PowerPoint: <code>.pptx</code>. Legacy <code>.xls</code>/<code>.ppt</code> are unsupported.",
                        "For encrypted files, expand <strong>Passwords</strong> and enter the open password for that file role.",
                        "A comparison requires both sides of an artifact pair; incomplete pairs are rejected.",
                        "Re-QC reuses a stored upload only after its hash still matches the recorded run.",
                    ]
                )
                _callout(
                    "XLSB formula coverage",
                    "The original XLSB supplies saved values. On Windows, installed desktop "
                    "Excel can supply Formula2 text; on Linux, LibreOffice runs inside a "
                    "networkless bubblewrap sandbox. Formula text is trusted only when its "
                    "coordinates exactly match an independent BIFF12 scan. Missing tools, "
                    "active/external content, timeouts, or mismatches fall back to "
                    "presence-only checks and are reported as degraded.",
                    warning=True,
                )

            with _guide_section("profiles", "Profiles, controls, and waivers"):
                _paragraph(
                    "The default profile runs general checks. Create a named profile when a "
                    "deliverable needs stable mappings, tolerances, region overrides, controls, "
                    "required slides, severity rules, or waivers. The in-app editor validates YAML."
                )
                _table(
                    ["Profile feature", "Typical use"],
                    [
                        ["Tolerance", "Suppress immaterial historical numeric differences"],
                        ["Sheet regions", "Pin long, wide, or block layouts when detection needs help"],
                        ["Cadence bands", "Separate adjacent monthly, quarterly, weekly, or dated axes"],
                        ["Chart windows", "Override rolling versus full history for one chart or series"],
                        ["Availability rules", "Allow blanks only after an explicit required-through period"],
                        ["Required ranges", "Require current cells to be populated"],
                        ["Unique ranges", "Detect duplicate business-key tuples"],
                        ["Numeric bounds", "Enforce valid ranges such as 0 to 1 for margins"],
                        ["Tie-outs", "Compare one target cell to sums of component ranges"],
                        ["Required slides / draft tokens", "Enforce final-deck completeness"],
                        ["Waivers", "Retain approved evidence as Expected until an expiry date"],
                    ],
                )
                _code(
                    """name: monthly-pack
excel:
  controls:
    required_ranges:
      - name: Current KPIs
        sheet: Dashboard
        range: B2:B8
    numeric_bounds:
      - name: Margin bounds
        sheet: Dashboard
        range: B4
        minimum: 0
        maximum: 1
    tie_outs:
      - name: Regional revenue total
        target: Dashboard!B2
        components: [Data!C2:C5]
        absolute_tolerance: 1
ppt:
  required_slides: [Executive Summary, Revenue Trend]
"""
                )
                _callout(
                    "Availability controls blankness only",
                    "Ignore ranges win first. Availability decides whether a blank is "
                    "required or allowed after the declared period. Refresh ranges decide "
                    "whether a nonblank value change is expected. Actual, forecast, and "
                    "target roles are descriptive and never allow blanks by themselves.",
                    warning=True,
                )
                _callout(
                    "Waivers are evidence, not deletion",
                    "A waiver requires a reason and ISO expiry date. The finding remains in the "
                    "record as Expected. Expired waivers produce warnings. Avoid broad class-only waivers.",
                    warning=True,
                )

            with _guide_section("coverage", "Coverage and severity"):
                _paragraph(
                    "Read coverage before interpreting a low finding count. Coverage describes "
                    "whether each check actually ran for the supplied files and format capabilities."
                )
                _table(
                    ["Coverage state", "Meaning"],
                    [
                        ["checked", "The check ran against the supplied artifact"],
                        ["degraded", "The check ran with a disclosed capability limitation"],
                        ["unavailable", "Required input or representation was absent"],
                    ],
                )
                _table(
                    ["Severity", "Default interpretation"],
                    [
                        ["Critical", "Material data, formula, deletion, or reconciliation failure"],
                        ["Warning", "Formula/structure/content drift requiring analyst review"],
                        ["Info", "Additive or presentation-level change"],
                        ["Expected", "Cadence growth, approved waiver, or reviewed expected change"],
                    ],
                )
                _callout(
                    "Zero findings is not automatically a full pass",
                    "A run with unavailable or degraded checks has a narrower conclusion. Record "
                    "those limitations in the review or attestation.",
                    warning=True,
                )
                _callout(
                    "Safeguards are visible",
                    "Excel workload warnings, accepted large-workbook overrides, findings caps, "
                    "and low-confidence alignment all degrade coverage and explain the affected "
                    "scope. They never silently truncate or guess.",
                )
                _paragraph(
                    "Tables, structured references, combo charts, interaction rules, "
                    "dependencies, availability, and speaker notes have explicit coverage. "
                    "Symbolic aggregate dependencies remain checked when membership and "
                    "transitive impacts are complete; unsupported formulas, chart parts, "
                    "rule families, or theme styles are disclosed separately."
                )

            with _guide_section("review", "Review findings in one place"):
                _list(
                    [
                        "Filter by severity, but inspect the coverage table first.",
                        "Expand a finding to see baseline/current values, element details, impacts, and nearby cells.",
                        "A shared <strong>root cause</strong> key groups multiple truthful symptoms at one location.",
                        "Use the severity selector only for an analyst disposition; it does not rewrite engine logic.",
                        "Add a comment explaining evidence, approval, source, or required follow-up.",
                        "Exports are regenerated from the reviewed state so comments and overrides are included.",
                    ]
                )
                _callout(
                    "Expected findings are hidden by default",
                    "Enable Expected in the severity filter when auditing cadence growth or waiver application.",
                )

            with _guide_section("mappings", "Excel to PowerPoint mappings"):
                _paragraph(
                    "Final-package QC extracts eligible figures from the current deck and ranks "
                    "candidate workbook cells by display match, label affinity, and numeric distance."
                )
                _list(
                    [
                        "Use a <strong>named profile</strong>; the default profile cannot persist confirmations.",
                        "Expand a suggestion and compare slide context, source labels, value, and match type.",
                        "Confirm only when the workbook cell is the intended source, not merely the same number.",
                        "A <strong>near match</strong> is useful for locating drift but is not a successful reconciliation.",
                        "Coverage distinguishes eligible, mapped, verified, mismatched, unresolved, and unmapped figures.",
                        "Visible native chart labels use chart, series, and category anchors and can be confirmed like text or table figures.",
                        "Re-run Final-package QC after confirmations to verify persisted mappings independently.",
                    ]
                )

            with _guide_section("reqc", "History and Re-QC"):
                _paragraph(
                    "Run history stores hashes, coverage, findings, reports, comments, mapping "
                    "review state, and Re-QC lineage. Use Re-QC after corrected deliverables arrive."
                )
                _list(
                    [
                        "Open <strong>Run history</strong>, select a run, then choose <strong>Re-QC</strong>.",
                        "Unchanged stored files are reused only after hash verification.",
                        "Missing, renamed, or changed files must be selected again in the correct role.",
                        "The delta reports resolved, new, and persisting non-Expected finding identities.",
                        "Review new coverage as well as the delta; a missing capability can change conclusions.",
                    ]
                )

            with _guide_section("privacy", "Privacy, sharing, and attestations"):
                _table(
                    ["Artifact", "Contains source information?", "Use"],
                    [
                        ["Numeric sanitize", "Yes", "Local testing only; not safe to share"],
                        ["Strict sanitize", "Transformed figures and period grammar", "Verified share candidate"],
                        ["Fingerprint", "No values or visible text", "Safest structural evidence"],
                        ["JSON", "Findings and compared values", "Private CI integration"],
                        ["JSON context", "Includes neighborhoods/candidates", "Private diagnostics only"],
                        [".qca attestation", "Full findings and reports", "Private audit/archive"],
                    ],
                )
                _callout(
                    "Safest sharing path",
                    "Start with a structural fingerprint. Use strict sanitize only when the Office "
                    "files are necessary, supply forbidden client tokens, inspect the redaction "
                    "manifest, and visually review the output before sharing.",
                    warning=True,
                )
                _callout(
                    "Macro files are refused",
                    "Strict sanitization does not accept XLSM because VBA can retain credentials "
                    "and client identifiers. Produce and review a macro-free XLSX copy first.",
                )

            with _guide_section("cli", "CLI and automation"):
                _code(
                    """# Headless comparison; exit 2 when critical findings exist
cadence-diff run --baseline-excel last.xlsx --current-excel this.xlsx \\
    --profile monthly --json findings.json --progress

# Deliberate local override after reviewing workload refusal
cadence-diff run --current-excel unusually-large.xlsx --allow-large-workbooks

# Structural-only evidence
cadence-diff fingerprint current.xlsx -o current.fingerprint.json

# Strict verified package redaction
cadence-diff sanitize-package --excel current.xlsx --ppt current.pptx \\
  --profile monthly --output-dir sanitized --forbid "Client Name"

# Signed audit evidence
cadence-diff run --current-excel current.xlsx --fail-on never \\
  --attestation run.qca
cadence-diff verify-attestation run.qca

# Validate a profile against actual artifacts
cadence-diff lint monthly.yaml --against-excel current.xlsx --against-ppt current.pptx"""
                )
                _list(
                    [
                        "Exit <code>0</code>: command completed without threshold/lint failures.",
                        "Exit <code>1</code>: usage, file, password, privacy, or runtime error.",
                        "Exit <code>2</code>: QC threshold reached, lint errors, unsafe verification, or invalid attestation.",
                        "Prefer <code>--password-env</code>, a mode-600 <code>--password-file</code>, or <code>--password-prompt</code>.",
                        "Use <code>--progress</code> for stderr phase updates; the web UI shows the same phases and supports cooperative cancellation.",
                        "<code>qc-tool</code> remains a compatibility alias for <code>cadence-diff</code>.",
                    ]
                )

            with _guide_section("network", "Local and temporary network access"):
                _paragraph(
                    "The default server binds to 127.0.0.1. Temporary LAN mode binds to all "
                    "interfaces, displays a warning on every page, and requires an expiry."
                )
                _code(
                    """qc-tool network status
qc-tool network lan --minutes 60 --data-dir data
qc-tool --port 8001 --data-dir data

# Return to local mode; a running LAN server stops within two seconds
qc-tool network local --data-dir data"""
                )
                _callout(
                    "No built-in authentication or TLS",
                    "LAN mode is unauthenticated. Router forwarding and host firewalls are "
                    "separate. Do not expose it publicly without an authenticated TLS reverse proxy.",
                    warning=True,
                )

            with _guide_section("troubleshooting", "Troubleshooting"):
                _table(
                    ["Symptom", "Action"],
                    [
                        ["Password required / invalid", "Enter the open password for the exact file role"],
                        ["Formula cache missing", "Open and recalculate in Excel, save, then rerun"],
                        [
                            "Workbook workload refused",
                            "Review the reported XML, shared-string, style, cell, and sheet-area metrics. "
                            "Use the per-run override only when sufficient local memory is confirmed.",
                        ],
                        [
                            "Run cancelled",
                            "Wait for the next safe boundary. Partial reports are removed and no successful run is recorded.",
                        ],
                        [
                            "XLSB formula checks degraded",
                            "Review coverage detail. Windows needs desktop Excel; Linux needs "
                            "LibreOffice and bubblewrap. Active/external content remains "
                            "presence-only by design.",
                        ],
                        ["Mapping unresolved", "Check slide wording/source movement and reconfirm the mapping"],
                        [
                            "Chart/table structure changed",
                            "Inspect the matched element, plot/series identity, source ranges, "
                            "axes, labels, and geometry before approving the change.",
                        ],
                        [
                            "Validation or conditional format changed",
                            "Confirm target coverage, formulas/operators, priority, stop-if-true, "
                            "and whether style coverage is degraded by theme/indexed colors.",
                        ],
                        [
                            "Availability finding looks unexpected",
                            "Lint the profile and verify required-through, period mapping, element/"
                            "series anchors, allow_blank_after, ignore ranges, and refresh ranges.",
                        ],
                        [
                            "Dependency coverage degraded",
                            "Review unsupported, invalid, or unparseable reference counts. Symbolic "
                            "aggregate references remain queryable and are disclosed separately.",
                        ],
                        ["Re-QC file not found", "Select the renamed or replacement artifact again"],
                        ["Profile will not save", "Correct the validation message or run qc-tool lint"],
                        ["Strict privacy verification fails", "Resolve each reported package issue or share only a fingerprint"],
                        ["Network page stops responding", "The temporary exposure may have expired; check qc-tool network status"],
                        ["Few or no findings", "Review coverage for unavailable/degraded checks before concluding clean"],
                    ],
                )

            with _guide_section("signoff", "Before sign-off"):
                _list(
                    [
                        "The selected mode matches the conclusion being made.",
                        "Every coverage row was reviewed, including degraded and unavailable checks.",
                        "Critical and Warning findings have a disposition or documented follow-up.",
                        "Expected findings and active waivers were sampled for correctness.",
                        "Final-package mapping coverage is acceptable and unresolved figures are documented.",
                        "Re-QC deltas are reviewed after corrected files arrive.",
                        "Exports or attestations were generated only after analyst comments and overrides were saved.",
                        "A human still performs any visual or domain checks outside the reported coverage.",
                    ]
                )
                _callout(
                    "QC evidence supports judgment",
                    "The tool makes checks reproducible and reviewable; it does not replace the "
                    "analyst's responsibility for unsupported visual, business, or upstream-data assertions.",
                )
