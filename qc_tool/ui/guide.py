# ruff: noqa: E501
"""Packaged in-app operator guide for QC Tool."""

from collections.abc import Iterator
from contextlib import contextmanager
from html import escape

from nicegui import ui

GUIDE_SECTIONS = (
    ("launch", "Launch and first run"),
    ("start", "Start here"),
    ("example", "Worked example"),
    ("modes", "Choose a mode"),
    ("files", "Files and formats"),
    ("profiles", "Profiles and controls"),
    ("coverage", "Coverage and severity"),
    ("review", "Review findings"),
    ("focus", "Desktop Office focus"),
    ("mappings", "Excel to PowerPoint mappings"),
    ("reqc", "Re-QC and history"),
    ("privacy", "Privacy and sharing"),
    ("cli", "CLI and automation"),
    ("network", "Network access"),
    ("troubleshooting", "Troubleshooting"),
    ("signoff", "Before sign-off"),
    ("reference", "Advanced capability reference"),
    ("glossary", "Glossary"),
)

COMMON_TASKS = (
    ("Launch QC Tool", "launch"),
    ("Run a first preflight", "launch"),
    ("See a worked example", "example"),
    ("Review a reporting cycle", "review"),
    ("Use Desktop Office focus", "focus"),
    ("Handle a workload refusal", "coverage"),
    ("Re-QC after a correction", "reqc"),
    ("Share evidence safely", "privacy"),
    ("Interpret a capability-limited run", "coverage"),
)

#: A deliberately tiny weekly tracker used to show what each check reports.
EXAMPLE_BASELINE = (
    ("Week", "Units", "Price", "Revenue"),
    ("2026-06-05", "120", "9.50", "=B2*C2"),
    ("2026-06-12", "135", "9.50", "=B3*C3"),
    ("2026-06-19", "128", "9.50", "=B4*C4"),
    ("2026-06-26", "141", "9.50", "=B5*C5"),
)

EXAMPLE_CURRENT = (
    ("Week", "Units", "Price", "Revenue"),
    ("2026-06-05", "120", "9.50", "=B2*C2"),
    ("2026-06-12", "135", "9.50", "=B3*C3"),
    ("2026-06-19", "128", "9.50", "1216"),
    ("2026-06-26", "152", "9.50", "=B5*C5"),
    ("2026-07-03", "147", "9.50", "=B6*C6"),
)

EXAMPLE_FINDINGS = (
    (
        "Critical",
        "formula replaced by a constant",
        "Revenue!D4",
        "<code>=B4*C4</code> became the typed literal <code>1216</code>. The cell "
        "no longer recalculates, so next week's edit will silently go stale.",
    ),
    (
        "Critical",
        "historical value changed",
        "Revenue!B4",
        "A closed week moved from 141 to 152. History should not move; this is "
        "reported as material because it is outside the recent window.",
    ),
    (
        "Expected",
        "cadence growth",
        "Revenue!A6:D6",
        "The new 2026-07-03 row continues the weekly cadence and reuses the same "
        "formula pattern, so it is recorded as expected rather than flagged.",
    ),
)

GLOSSARY = (
    (
        "pattern group",
        "One analyst decision covering findings that repeat the same semantic "
        "pattern. Counted as a review item.",
    ),
    (
        "atomic finding",
        "One individual difference or control result in Atomic output. In "
        "Decision output, a population finding record can represent many cell "
        "changes while preserving exact membership evidence.",
    ),
    (
        "story",
        "A plain-language summary linking findings through confirmed structural "
        "relationships. Stories never change counts or severities.",
    ),
    (
        "scope",
        "The sheets and slides the run compared. Everything is compared unless "
        "a narrower scope was selected and disclosed.",
    ),
    (
        "coverage",
        "Whether each check was checked, degraded, unavailable, or not included "
        "because that artifact or workflow was not part of the run. A low finding "
        "count with unavailable checks is not a clean result.",
    ),
    (
        "Expected reason",
        "The typed justification for an Expected finding, such as cadence "
        "growth or a rolling chart window.",
    ),
    (
        "workbook workload refusal",
        "A safety stop raised when workbook size or formula-link complexity may "
        "use too much memory or time. It is not a QC finding.",
    ),
    (
        "projected finding volume",
        "An estimate of how many differences the changed sheets could produce. "
        "It helps choose scope, but it is not a memory-safety check.",
    ),
)

#: Client-side only: no network, no external assets. Runs after NiceGUI mounts.
GUIDE_SCRIPT = r"""
(function () {
  const box = document.getElementById('guide-q');
  if (!box || box.dataset.wired) return;
  box.dataset.wired = '1';
  const count = document.getElementById('guide-q-count');
  const sections = Array.from(document.querySelectorAll('.guide-section'));
  const links = Array.from(document.querySelectorAll('.guide-toc-link'));
    const aliases = {
        launch: 'install windows path command not found shortcut start browser first run data directory storage localappdata xdg_data_home',
        coverage: 'capability limited unavailable degraded not included omitted input not checked workload refusal override memory projected findings safety',
        mappings: 'mapping unavailable opaque screenshot raster image claim',
        review: 'unreviewed only replace all confirm severity reviewed note',
        reqc: 'rerun repeat carry forward history',
        files: 'xlsb formula encrypted password same file workload memory refusal',
        focus: 'desktop office excel powerpoint bind confirm open cell shape',
        reference: 'vba power query connection structured spill circular parser',
    };
  const mobile = () => window.matchMedia('(max-width: 640px)').matches;

  sections.forEach(function (s) {
    const title = s.querySelector('.guide-section-title');
    if (!title) return;
    title.setAttribute('role', 'button');
    title.setAttribute('tabindex', '0');
    const toggle = function () {
      s.classList.toggle('collapsed');
      title.setAttribute('aria-expanded', String(!s.classList.contains('collapsed')));
    };
    title.addEventListener('click', toggle);
    title.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); }
    });
    if (mobile() && s.id !== 'launch') s.classList.add('collapsed');
    title.setAttribute('aria-expanded', String(!s.classList.contains('collapsed')));
  });

  const openTarget = function () {
    const id = (location.hash || '').slice(1);
    if (!id) return;
    const s = document.getElementById(id);
    if (s && s.classList.contains('collapsed')) {
      s.classList.remove('collapsed');
      s.querySelector('.guide-section-title')?.setAttribute('aria-expanded', 'true');
    }
    s?.scrollIntoView({ block: 'start' });
  };
  window.addEventListener('hashchange', openTarget);
  openTarget();

  box.addEventListener('input', function () {
    const q = box.value.trim().toLowerCase();
        const tokens = q.split(/\s+/).filter(Boolean);
    let hits = 0;
    sections.forEach(function (s) {
            const searchable = (s.innerText + ' ' + (aliases[s.id] || '')).toLowerCase();
            const match = !tokens.length || tokens.every(function (token) {
                return searchable.indexOf(token) !== -1;
            });
      s.hidden = !match;
      if (q && match) { hits += 1; s.classList.remove('collapsed'); }
    });
    links.forEach(function (a) {
      const id = (a.getAttribute('href') || '').split('#')[1];
      const s = document.getElementById(id);
      a.hidden = Boolean(q) && Boolean(s) && s.hidden;
    });
    count.textContent = q
      ? hits + ' matching topic' + (hits === 1 ? '' : 's')
      : '';
  });

  const more = document.getElementById('guide-toc-more');
  if (more) {
    more.addEventListener('click', function () {
      const toc = document.querySelector('.guide-toc');
      toc.classList.toggle('expanded');
      more.textContent = toc.classList.contains('expanded')
        ? 'Fewer topics'
        : 'All topics';
    });
  }
})();
"""

PROFILE_CONTROLS_EXAMPLE = "\n".join(
    (
        "name: monthly-pack",
        "excel:",
        "  controls:",
        "    required_ranges:",
        "      - name: Current KPIs",
        "        sheet: Dashboard",
        "        range: B2:B8",
        "    numeric_bounds:",
        "      - name: Margin bounds",
        "        sheet: Dashboard",
        "        range: B4",
        "        minimum: 0",
        "        maximum: 1",
        "    tie_outs:",
        "      - name: Regional revenue total",
        "        target: Dashboard!B2",
        "        components: [Data!C2:C5]",
        "        absolute_tolerance: 1",
        "      - name: Net contribution",
        "        target: NetContribution",
        "        terms:",
        "          - reference: RevenueData[Revenue]",
        "            operation: add",
        "          - reference: RevenueData[Cost]",
        "            operation: subtract",
        "ppt:",
        "  required_slides: [Executive Summary, Revenue Trend]",
        "",
    )
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

    ui.html(
        '<div class="guide-search">'
        '<input id="guide-q" type="search" autocomplete="off" '
        'placeholder="Search the guide" aria-label="Search the guide">'
        '<span id="guide-q-count" class="guide-search-count"></span>'
        "</div>"
    )
    with ui.element("div").classes("guide-tasks"):
        ui.label("Common tasks").classes("guide-tasks-title")
        for label, anchor in COMMON_TASKS:
            ui.link(label, f"/guide#{anchor}").classes("guide-task-link")

    with ui.element("div").classes("guide-layout"):
        with ui.element("aside").classes("guide-toc"):
            ui.label("On this page").classes("guide-toc-title")
            for anchor, title in GUIDE_SECTIONS:
                ui.link(title, f"/guide#{anchor}").classes("guide-toc-link")
            ui.html('<button id="guide-toc-more" type="button">All topics</button>')

        with ui.element("article").classes("guide-content"):
            with _guide_section("launch", "Launch and first run"):
                _paragraph(
                    "After installation, start the local app with the Python environment "
                    "that contains cadence-diff. The module command is the most reliable "
                    "choice on Windows because it does not depend on the Scripts directory "
                    "being listed in PATH."
                )
                _code(
                    """# Works from an activated pip or Conda environment
python -m qc_tool

# Windows: create a per-user Desktop shortcut for this exact environment
python -m qc_tool shortcut install
python -m qc_tool shortcut status
python -m qc_tool shortcut remove

# Compatibility: enable Desktop Office focus for this launch only
python -m qc_tool --desktop-focus"""
                )
                _paragraph(
                    "QC Tool keeps uploads, profiles, run history, generated reports, "
                    "and local server state together in one data directory."
                )
                _table(
                    ["Launch context", "Default data directory"],
                    [
                        [
                            "Installed on Windows",
                            "<code>%LOCALAPPDATA%\\qc-tool</code> (normally "
                            "<code>C:\\Users\\&lt;you&gt;\\AppData\\Local\\qc-tool</code>)",
                        ],
                        [
                            "Installed on Linux",
                            "<code>$XDG_DATA_HOME/qc-tool</code> when "
                            "<code>XDG_DATA_HOME</code> is set; otherwise "
                            "<code>~/.local/share/qc-tool</code>",
                        ],
                        [
                            "Source checkout: python main.py",
                            "<code>&lt;repository&gt;/data</code> (developer entry point only)",
                        ],
                    ],
                )
                _code(
                    """# Show the resolved default for this account and operating system
python -m qc_tool --help

# Select a different data directory before the server starts
python -m qc_tool --data-dir "/path/to/qc-data"
cadence-diff --data-dir "/path/to/qc-data"
"""
                )
                _list(
                    [
                        "The data directory contains managed upload copies, profile YAML, SQLite run history, generated reports, and private server or launcher state. Treat the whole directory as sensitive local working data.",
                        "QC Tool restricts managed files to the current user where the operating system supports it. Choose a trusted local location; do not use a team share as an access-control substitute.",
                        "The directory is selected before startup and cannot be switched from Local app settings while the server is running. Stop QC Tool before changing <code>--data-dir</code>.",
                        "Each path is an independent store. Starting with a new path shows a separate history and profile set; it does not migrate or delete the old directory.",
                        "Quote a custom path that contains spaces. Use the same <code>--data-dir</code> on later launches to return to that store.",
                    ]
                )
                _callout(
                    "Changing the path does not migrate history",
                    "To keep using existing runs and profiles, continue pointing QC Tool at "
                    "their original data directory. Relocating an existing store requires a "
                    "separate, verified migration rather than copying it while the server runs.",
                    warning=True,
                )
                _list(
                    [
                        "The browser opens at <code>http://127.0.0.1:8080</code>. The server remains local and unauthenticated network access stays off.",
                        "If Windows says <code>cadence-diff is not recognized</code>, use <code>python -m qc_tool</code>. Do not search for or copy a random executable from another environment.",
                        "A fresh installation opens in <strong>Current-file preflight</strong>. Upload one current Excel workbook and/or PowerPoint deck; the built-in default profile is enough for a first run.",
                        "After you choose another mode, QC Tool remembers it on this machine. A Re-QC link always restores the original run mode.",
                        "Use the settings icon in the header to create or remove the Windows shortcut and to manage Desktop Office focus.",
                        "Use the power icon in the header to stop the local server when finished.",
                    ]
                )
                _callout(
                    "Shortcut lifecycle",
                    "The shortcut points to the absolute Python environment that created it, "
                    "so PATH and Conda activation are not required. If that environment is "
                    "deleted or moved, recreate the shortcut from the replacement environment. "
                    "Package installation never changes the Desktop automatically.",
                )

            with _guide_section("start", "Start here"):
                _callout(
                    "Check prerequisites before you compare",
                    "If the profile pins a scenario or selector cell (for example a "
                    "dropdown-driven forecast case), baseline and current must show "
                    "the exact same, non-blank value there. A mismatch blocks the run "
                    "before any comparison, report, or history entry is produced: "
                    "select the same scenario in both files, fully recalculate, save, "
                    "then run again. Configure pins in a named profile's "
                    "<strong>Comparison prerequisites</strong> (Excel, advanced "
                    "section) as a sheet and cell for each selector that must match. "
                    "Prerequisite cells are manually pinned for both OOXML and XLSB "
                    "workbooks today.",
                    warning=True,
                )
                _callout(
                    "Ranked or sorted tables",
                    "When a large positional block looks like the same records in a "
                    "different order, QC pauses before creating mass cell-to-cell "
                    "noise. Review row matching: pick the columns that match rows "
                    "by (and, optionally, columns whose order-only values should "
                    "be ignored), choose how duplicate identities are handled, and "
                    "either run once without saving or save the rule to an existing "
                    "or new named profile. Detected header labels appear beside stable "
                    "column letters, including when a table has preamble rows. QC then "
                    "re-runs automatically. "
                    "Confirmed row order and rank ordinals are ignored; formulas, "
                    "styles, structure, and other business values remain checked. "
                    "If no safe identity can be proved automatically, add a "
                    "<strong>Row identity rule</strong> under the sheet in Manage "
                    "profiles or leave positional comparison unchanged.",
                    warning=True,
                )
                _paragraph(
                    "A defensible run has four parts: select the mode that answers the "
                    "actual QC question, supply only the files that mode needs, review "
                    "what was and was not checked, then disposition every material finding."
                )
                _list(
                    [
                        "Select a <strong>QC mode</strong> before choosing files.",
                        "Use the built-in <strong>default profile</strong> for a first or one-off run. Create a named profile only for recurring controls, mappings, tolerances, or waivers.",
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

            with _guide_section("example", "Worked example"):
                _paragraph(
                    "A four-week revenue tracker, refreshed for a new week. The "
                    "sample is deliberately tiny so every reported finding can be "
                    "traced by eye. Upload the previous file as Baseline and the "
                    "refreshed file as Current, then run a cycle comparison."
                )
                _paragraph("Baseline — last week's file:")
                _table(
                    list(EXAMPLE_BASELINE[0]),
                    [list(row) for row in EXAMPLE_BASELINE[1:]],
                )
                _paragraph("Current — this week's file, with three differences:")
                _table(
                    list(EXAMPLE_CURRENT[0]),
                    [list(row) for row in EXAMPLE_CURRENT[1:]],
                )
                _paragraph("QC Tool reports exactly this:")
                _table(
                    ["Severity", "Finding", "Where", "Why"],
                    [list(row) for row in EXAMPLE_FINDINGS],
                )
                _paragraph(
                    "Two decisions need an analyst and one does not. That split is "
                    "the whole point: growth is recognised as growth, so the "
                    "hardcoded formula and the moved history stay visible instead "
                    "of drowning in a diff of every changed cell."
                )
                _list(
                    [
                        "Open the <strong>Review queue</strong>: the two Critical items are the top rows.",
                        "Select a row to see baseline and current values side by side with the surrounding cells.",
                        "Set <strong>Show severities</strong> to include Expected when you want to audit the growth row itself.",
                        "Add a specific comment such as <code>Revenue!D4 confirmed by J. Smith (data owner), 2026-06-30</code>, then export — the comment travels with the report.",
                        "Fix the source file, then use <strong>Re-QC</strong> to prove the two Critical items are resolved.",
                    ]
                )
                _callout(
                    "Why the new week is not a finding",
                    "Cadence growth is inferred from the period column and the reused "
                    "formula pattern. If the new row had broken the pattern, or landed "
                    "in the middle of history, it would be reported instead of accepted.",
                )
                _callout(
                    "Same file twice proves nothing",
                    "If Baseline and Current are the same file, the run reports almost "
                    "nothing and that clean result is meaningless. The Compare page "
                    "warns when the two look identical.",
                    warning=True,
                )

            with _guide_section("modes", "Choose the right QC mode"):
                _paragraph(
                    "A fresh installation selects Current-file preflight. The Compare page "
                    "then remembers your last manual choice; Re-QC always uses the recorded mode."
                )
                _table(
                    ["Mode", "Required inputs", "Use it to answer", "Cannot establish"],
                    [
                        [
                            "Current-file preflight",
                            "Latest Excel member(s) and/or latest PPT",
                            "Is this package internally sound now?",
                            "Whether historical content changed",
                        ],
                        [
                            "Cycle comparison",
                            "Baseline/current Excel members and/or PPT pair",
                            "What changed, excluding expected cadence growth?",
                            "Whether saved formula results are numerically correct, or what Excel would recalculate from upstream data",
                        ],
                        [
                            "Final-package QC",
                            "One to eight current Excel members plus latest PPT",
                            "Does the final deck reconcile to its nominated workbooks?",
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
                _paragraph(
                    "Alongside QC mode, choose a <strong>Finding output</strong>: "
                    "Decision (recommended) groups related findings into compact "
                    "populations for large runs, using your profile's own population "
                    "settings when it already enables them, else a conservative "
                    "built-in default. Profile reproduces exactly what your saved "
                    "profile's own review policy already produces. Atomic (advanced) "
                    "lists every individual finding, ungrouped -- forensic detail, "
                    "outside the interactive review target for very large runs. "
                    "Re-QC preselects the source run's own recorded output; changing "
                    "it is an explicit, disclosed representation change, not silent."
                )

            with _guide_section("files", "Files, passwords, and formats"):
                _list(
                    [
                        "Excel: <code>.xlsx</code>, <code>.xlsm</code>, and <code>.xlsb</code> for QC input.",
                        "PowerPoint: <code>.pptx</code>. Legacy <code>.xls</code>/<code>.ppt</code> are unsupported.",
                        "For encrypted files, expand <strong>Passwords</strong> and enter the open password for that file role.",
                        "Baseline and current files may use different open passwords. Use distinct filenames when encrypted inputs require different passwords.",
                        "Excel sheet/workbook protection and locked, unlocked, or formula-hidden cells control editing; they do not block read-only QC, and protection settings are not themselves audited.",
                        "IRM, sensitivity-label encryption, or missing filesystem read permission can still prevent automated access.",
                        "A comparison requires both sides of an artifact pair; incomplete pairs are rejected.",
                        "Add up to eight Excel workbook members per side. Stable member IDs pair baseline and current workbooks; added or removed members are reported explicitly.",
                        "Duplicate sheet names remain qualified by workbook member. Ambiguous unscoped profile rules are refused instead of being applied to several workbooks.",
                        "Re-QC reuses a stored upload only after its hash still matches the recorded run.",
                    ]
                )
                _callout(
                    "XLSB values, formulas, and workload",
                    "QC always reads saved values from the original XLSB. Before loading those "
                    "values, it counts retained cells, formulas, and binary worksheet bytes; "
                    "an oversized file is refused unless the workload override is deliberately "
                    "enabled. Formula text is an additional capability, read by one of three "
                    "engines: a bundled native decoder (fastest, no external application), "
                    "installed desktop Excel on Windows, or LibreOffice in a Linux networkless "
                    "sandbox. A profile's <code>formula_engine</code> setting picks "
                    "<code>native</code>, <code>excel</code>, <code>libreoffice</code>, or the "
                    "default <code>auto</code> (native when available, otherwise the existing "
                    "platform default). Formula text is accepted only when its cells exactly "
                    "match an independent structural scan of the binary file. A missing engine, "
                    "active or external content, timeouts, or mismatches fall back to "
                    "formula-presence checks and are clearly reported as degraded -- a run never "
                    "fails because of a formula-engine choice the current install cannot satisfy. "
                    "Saved values are still read. A private, bounded cache remembers a "
                    "workbook's extracted formula text so an unchanged file skips repeat "
                    "extraction on its next QC run; the cache key includes which engine "
                    "produced the entry, so switching engines never reuses a stale one, and a "
                    "cache hit is revalidated against a fresh structural scan every time and "
                    "never changes a finding. Report or clear it from Local app settings or "
                    "<code>qc-tool formula-cache status|clear</code>.",
                    warning=True,
                )

            with _guide_section("profiles", "Profiles, controls, and waivers"):
                _paragraph(
                    "The default profile runs general checks. Create a named profile when a "
                    "deliverable needs stable mappings, tolerances, region overrides, controls, "
                    "required slides, severity rules, or waivers. Manage profiles exposes typed "
                    "Core and Advanced Excel/PPT views over the complete contract. Lists and "
                    "mappings can be added, edited, removed, and reordered without dropping "
                    "fields that are not currently expanded."
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
                        ["Tie-outs", "Compare one target to additive components or signed named/table terms"],
                        ["Required slides / draft tokens", "Enforce final-deck completeness"],
                        ["Waivers", "Retain approved evidence as Expected until an expiry date"],
                    ],
                )
                _code(PROFILE_CONTROLS_EXAMPLE)
                _callout(
                    "Form and YAML are one draft",
                    "Advanced YAML shows the complete canonical profile. Apply parses and "
                    "previews YAML in the typed form; Reset discards unapplied YAML text. "
                    "Save runs static lint, refuses a changed source file by its exact byte "
                    "hash, and atomically replaces only a valid named profile. The built-in "
                    "default remains read-only.",
                )
                _callout(
                    "Optional file-backed validation",
                    "Select files on Compare, then validate from Manage profiles to check "
                    "sheet, range, slide, and mapping references with the same local read-only "
                    "loaders and passwords. Current files take precedence over baseline files. "
                    "Workbook workload refusal and cancellation remain active. XLSB reference "
                    "lint stays bounded; formula-text adapter availability and degradation are "
                    "reported by the subsequent QC run.",
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
                _callout(
                    "PowerPoint media has two separate claims",
                    "ppt-media-structural hashes embedded media bytes without decoding them. "
                    "A Warning means an embedded image was added, removed, or its bytes changed; "
                    "the digest, pixels, source path, and image text are never reported. Linked "
                    "or malformed image relationships degrade that check instead of guessing. "
                    "ppt-media-visual remains unavailable because byte identity cannot prove "
                    "rendered appearance, crop/layout equivalence, or OCR text.",
                    warning=True,
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
                _paragraph(
                    "Numeric changes are judged on two separate questions: how large is the "
                    "change, and how old is the period? Sub-display-precision rounding noise "
                    "is Info. A profile can define tolerance bands for particular ranges; an "
                    "accepted change remains visible as within-tolerance Info. A current or "
                    "recent change is normally Warning, but a sign flip, crossing zero, or a "
                    "10x-or-larger jump remains Critical. Material changes to older periods "
                    "stay Critical. A profile can define which recent periods may be restated, "
                    "which ranges are expected refreshes, and how each materiality level maps "
                    "to severity."
                )
                _paragraph(
                    "An inherited error is one that was already present in the baseline; it "
                    "was not introduced by the current file. It may be Info or Warning only "
                    "when the evidence shows the same non-structural condition. Structural "
                    "errors, and every new or changed error, remain Critical. A row or column "
                    "label produced by a formula points reviewers to the upstream cause; a "
                    "manually rewritten historical key remains Critical."
                )
                _paragraph(
                    "Change stories are one of the result views, alongside the review "
                    "queue, coverage, and atomic evidence. "
                    "They link structural drivers (new table columns, named ranges, data "
                    "validations) to the formula rollouts referencing them, isolate "
                    "export noise, in-window restatements, derived-label churn, and "
                    "inherited conditions, and leave everything unexplained in a residual "
                    "story — the true review queue. Stories never alter severities or "
                    "counts."
                )
                _paragraph(
                    "Two run-time controls exist on the run form and CLI: an analyst "
                    "acceptance threshold (absolute value and/or percentage; either "
                    "bound accepts; in-band changes stay visible as within-tolerance "
                    "Info and the threshold is disclosed with the run) and a "
                    "comparison scope (pick sheets and slides; the validated scope and "
                    "selected/total counts are disclosed). In a multi-workbook package, "
                    "scope is set per workbook member; a member with no sheet selection "
                    "remains fully compared. Scope narrows only Excel and PowerPoint "
                    "findings. Package mappings, source suggestions, readable-claim counts, "
                    "and period reconciliation still use the whole package. Every selected "
                    "file still loads fully, so scope does not reduce physical loading risk "
                    "or bypass a workload refusal. Both controls default off."
                )
                _paragraph(
                    "A large block of the same saved error may be shown as one review "
                    "decision instead of hundreds of repeated decisions: this happens at "
                    "200 cells by population alone, or from 20 cells when they form a clear "
                    "cluster. Grouping never makes an error safer. Only data-state errors "
                    "(#N/A, #DIV/0!, #VALUE!, #NUM!) can become Warning when formula and "
                    "clustering evidence support it. Sparse errors, structural breakage "
                    "(#REF!, #NAME?, #NULL!), and every new or changed error stay Critical. "
                    "External links and active content also produce Critical findings in all "
                    "run modes."
                )
                _callout(
                    "Zero findings is not automatically a full pass",
                    "A run with unavailable or degraded checks has a narrower conclusion. Record "
                    "those limitations in the review or attestation.",
                    warning=True,
                )
                _callout(
                    "Safeguards are visible",
                    "Excel workload warnings, accepted workload overrides, and low-confidence "
                    "alignment all degrade coverage and explain the affected scope. Every "
                    "finding is retained; safeguards never silently truncate or guess.",
                )
                _callout(
                    "Workbook workload override: use it only after a refusal",
                    "The normal run checks two safety areas: physical workbook load (OOXML or "
                    "XLSB size, cells, and package bytes) and formula-link complexity (how many "
                    "formulas, references, projected dependencies, and interaction rules must "
                    "be analysed). Override workbook workload refusals applies to every workbook "
                    "in that run and bypasses both safety stops. First read the exact refusal "
                    "reason. If you cannot confirm enough memory and time, stop and ask a senior "
                    "reviewer. Do not tick the override in advance, and do not treat sheet scope "
                    "as a memory workaround because files still load fully. The separate "
                    "projected-finding dialog estimates review volume; it is not a safety check "
                    "and does not make a workbook cheaper to load.",
                    warning=True,
                )
                _callout(
                    "Dependency indexing has its own, separate size policy",
                    "Above a documented formula-cell or projected-dependency threshold, "
                    "dependency indexing is skipped for that run instead of attempted -- "
                    "circular-reference detection, formula impacts, chart impacts, and "
                    "PowerPoint chart impacts all degrade together, and every affected "
                    "coverage row (and the exported reports) states the reason as "
                    "\"skipped by size policy\". This is separate from, and does not require, "
                    "Override workbook workload refusals: a workbook that already needed that "
                    "override just to load and compare can still have its dependency indexing "
                    "skipped, or forced back on, independently. Force full dependency indexing "
                    "only when you have confirmed enough local memory and time -- it can be the "
                    "most expensive phase of a large run.",
                    warning=True,
                )
                _paragraph(
                    "When alignment cannot deterministically pair cells, the run records an "
                    "alignment trust manifest listing per-region paired counts, low-confidence "
                    "regions, and unpaired regions. Use Coverage → Alignment trust per region "
                    "to inspect these factual counts when interpreting degraded comparisons."
                )
                _paragraph(
                    "Tables, structured references, combo charts, interaction rules, "
                    "dependencies, availability, and speaker notes have explicit coverage. "
                    "Symbolic aggregate dependencies remain checked when membership and "
                    "transitive impacts are complete. Proven XLSX/XLSM spill references and "
                    "unambiguous implicit intersections are resolved; missing spill extents, "
                    "including XLSB extents, degrade reference coverage. Unsupported formulas, chart parts, "
                    "rule families, or theme styles are disclosed separately."
                )
                _paragraph(
                    "Circular-reference detection builds a separate compact graph over "
                    "formula cells only. It reports exact self and multi-cell cycles; "
                    "unparseable, unsupported, invalid, or edge-budgeted references "
                    "degrade the circular-reference coverage row instead of producing a "
                    "false clean result."
                )
                _paragraph(
                    "Workbook metadata beyond the grid is compared too, each with its own "
                    "coverage row so a limitation is never silent."
                )
                _table(
                    ["Check", "What it compares"],
                    [
                        [
                            "excel-defined-name-scope",
                            "Workbook-scoped and sheet-scoped defined names. The same name "
                            "in two sheets stays two separate identities, shown as "
                            "Sheet!Name. XLSB reports unavailable because its workbook part "
                            "is binary.",
                        ],
                        [
                            "excel-vba",
                            "VBA module inventory and module source text, for XLSM, XLSB, "
                            "and any XLSX carrying a project. Findings carry the module "
                            "name, line counts, changed line ranges, and a digest \u2014 never "
                            "the macro source. No macro is ever executed.",
                        ],
                        [
                            "excel-comments",
                            "Cell comments and threaded notes, by sheet and cell.",
                        ],
                        [
                            "excel-power-query",
                            "Power Query definitions per named query. If the mashup cannot "
                            "be opened, coverage degrades and comparison falls back to a "
                            "definition digest, so a changed query is still detected.",
                        ],
                        [
                            "excel-connections",
                            "Data connections by name, kind, and target digest. The "
                            "connection string, URL, and command text are never read into a "
                            "finding, report, or log. Nothing is fetched and no query runs.",
                        ],
                    ],
                )
                _callout(
                    "A locked VBA project is still compared",
                    "Project protection restricts the Excel editor; it does not encrypt the "
                    "module streams. The run reports the project as protected and still "
                    "compares its text.",
                )

            with _guide_section("review", "Review findings in one place"):
                _paragraph(
                    "Results open on the Review queue. The run header keeps the outcome, "
                    "capability state, decision counts, scope, exports, and Re-QC visible "
                    "in every view, and Stories, Coverage, and Atomic evidence are one "
                    "click away."
                )
                _paragraph("For a first review, follow four steps:")
                _list(
                    [
                        "Read the run outcome and capability status before interpreting the finding count.",
                        "Select the first open review item and compare its baseline/current evidence and nearby cells.",
                        "Confirm or change severity and add a specific evidence-based comment.",
                        "Continue until material items are dispositioned, then generate exports or finalize the review.",
                    ]
                )
                _paragraph("Review controls and advanced workflow:")
                _list(
                    [
                        "Filter by severity, class, sheet or slide, or free text, and read the capability status before concluding a run is clean. The sheet/slide list comes from this run's findings; <em>whole file</em> covers findings that belong to no single sheet.",
                        "The queue pages under its own footer: choose 10 to 100 rows per page (remembered on this machine), pick an order — <em>priority</em> (the evidence default), severity, location, or findings — with a direction toggle, and drag a column boundary in the header to resize; double-click the boundary to restore the defaults. Ordering moves a related series and its decisions together, never splitting them.",
                        "The <strong>review time</strong> box at the top right is an explicit Start/Pause timer with a four-hour cap per session; finalizing the run freezes it, and the recorded minutes feed the longitudinal dossier.",
                        "The queue is ordered by evidence, not by sheet position. The detail panel says <em>prioritized because</em> and names the counts it scored on: severity, materiality, historical position, provenance, downstream impacts, population size, and whether a story explains it.",
                        "Ordering only reorders. Every review item and finding record stays reachable; nothing is hidden.",
                        "Waived, already-reviewed, and expected-growth decisions sink below everything still open, because they need no new judgement.",
                        "Pattern review-item counts are analyst decisions; finding-record counts are stored evidence rows; represented-change counts include every cell summarized by a population. Spatial review counts remain a compatibility metric.",
                        "Select a review item to see its evidence axes, baseline/current values, impacts, and nearby cells in the detail panel.",
                        "For formula logic changes, expand <strong>Formula token diff</strong> to see a normalized token-level comparison; added tokens are underlined and removed tokens are struck through.",
                        "Open a group for paged members, or use the Atomic evidence view for every stored finding record and each population's exact membership.",
                        "Use Unreviewed only to preserve prior member decisions. Replace all overwrites every existing member decision in that group, so use it only when you intend to replace prior analyst work.",
                        "Use the Review selector to show All, Needs review, or Reviewed groups. Keyboard triage is available: <code>j</code>/<code>k</code> or Arrow keys move between visible review items; <code>1</code>/<code>2</code>/<code>3</code>/<code>4</code> set Critical/Warning/Info/Expected for the current unreviewed group and advance; <code>c</code> confirms the current severity as reviewed.",
                        "What-if preview uses private typed numeric evidence to show how temporary acceptance bounds or a materiality review floor would change atomics and decisions. It ignores analyst overrides, never parses display strings, never writes the run or profile, and keeps accepted changes visible as Info.",
                        "A shared <strong>root cause</strong> key groups multiple truthful symptoms at one location.",
                        "Use the severity selector only for an analyst disposition; it does not rewrite engine logic.",
                        "Add a specific comment naming the evidence, approver or source, date, and required follow-up. Avoid comments such as <em>looks fine</em> or <em>checked</em> with no support.",
                        "Exports are regenerated from the reviewed state so comments and overrides are included.",
                        "Excel and HTML exports lead with semantic pattern groups while retaining every finding record and population membership; Excel links stay inside the report workbook.",
                        "On very large runs (over 50,000 findings) report files are not written at run time — the run becomes reviewable sooner, and <strong>Generate Excel/HTML report</strong> on the run page builds the file on demand, stores it with the run, and downloads it. Expect several minutes for a million-finding report; reviewing continues meanwhile.",
                        "When group-first output is enabled on a profile, a large run of identical formula-logic or number-format changes is shown as one <strong>population</strong> item instead of one row per cell. Its detail shows the member count, the current-side cell ranges, and how baseline cells map to them; sample excerpts and the full paged member list remain available. A smaller or non-uniform group of the same class still reviews as individual finding records — nothing is silently combined.",
                    ]
                )
                _callout(
                    "Expected findings are hidden by default",
                    "Enable Expected in the severity filter when auditing cadence growth or waiver application.",
                )
                _table(
                    ["Group severity", "What is recorded"],
                    [
                        [
                            "keep",
                            "Severity is left exactly as the engine set it. With a comment "
                            "this is a note, not a re-classification: counts and exported "
                            "severities do not move.",
                        ],
                        [
                            "any severity",
                            "An explicit analyst disposition. Use Confirm severity to "
                            "record the current severity as reviewed without changing it.",
                        ],
                    ],
                )
                _list(
                    [
                        "A reviewed group row shows <strong>reviewed N/M</strong>: how many members carry an override or a note. Member and atomic rows show <strong>analyst</strong> for a severity override and <strong>note</strong> for a comment alone.",
                        "The toast reports how many members actually changed. Under Unreviewed only a smaller number means members you had already decided were skipped.",
                        "Neither mode erases anything: keep never clears an override, and a blank comment never clears a comment. Clear a note from the row editor in Atomic evidence.",
                        "Severity is part of what groups findings, so changing it regroups them. Under Unreviewed only a group can split, leaving members you already decided in their own group.",
                        "Annotations belong to the run. A Re-QC starts a new run with no annotations; the delta reports resolved, new, and persisting.",
                    ]
                )

                _paragraph(
                    "<strong>Related series</strong> rows collapse one logical time series "
                    "that the engine correctly split across several decisions, so you can "
                    "read a measure column in one place."
                )
                _list(
                    [
                        "A related-series row is a <strong>navigation lens, not a decision</strong>. Decision and atomic counts in the run header keep counting canonical decisions, so promoting a series can make the queue show more top-level rows, never fewer.",
                        "Parent labels are <strong>structural only</strong>: sheet, measure column letter or row number, and how many periods changed. No metric name, header text, or cell value is captured, stored, or shown in the label.",
                        "Parents start collapsed. Click, or press Enter or Space, to expand. <code>j</code>/<code>k</code> move focus over a parent without expanding it, and the <code>1</code>-<code>4</code> and <code>c</code> shortcuts are deliberately unavailable on a parent row.",
                        "Each child segment is one canonical decision intersected with that series: it keeps that decision's identity, severity, and temporal context.",
                        "<strong>Confirm visible</strong> records every visible, still-unreviewed finding at <em>its own current severity</em>. It is a confirmation, never a bulk override or a replacement, and it lists the exact counts before you apply it.",
                        "Under an active filter, Confirm visible touches only the visible child segments and states how many segments and findings stay hidden and untouched.",
                        "Parent context merges the neighbouring cells this run stored around each finding. When some related cells were not stored, the panel says <strong>N of M related finding cells shown</strong>; the parent still lists every related finding.",
                        "A period populated for the first time this cycle joins its column as a <strong>new period</strong> segment, and a historical cell cleared to blank joins as a <strong>cleared period</strong> segment, provided sibling measures still prove that period exists. Both sort at their period position so a parent reads in time order, carry no materiality tier or temporal context (there is no baseline/current pair to measure), and keep their own severity exactly like any other member.",
                        "A parent that contains a new or cleared period is labelled <strong>periods affected</strong> rather than <em>changed periods</em>, because a period that appeared or disappeared was not changed.",
                        "A <strong>wiped period row</strong> is one event, not one finding per column: when no measure retains data at that period, its cleared cells stay together in their own decision instead of being scattered across column parents.",
                        "Runs recorded before this feature keep today's grouping. Re-QC that deliverable to get producer-authored series evidence; nothing is guessed from coordinates.",
                    ]
                )

            with _guide_section("focus", "Desktop Office focus"):
                _paragraph(
                    "On Windows, a finding can navigate to the same saved cell or shape in an "
                    "already-open Excel or PowerPoint document. This is an optional local "
                    "review aid, not part of the QC verdict or signed evidence."
                )
                _list(
                    [
                        "Keep the server in local mode. Desktop focus is unavailable in LAN mode even when your preference is remembered.",
                        "Open the exact workbook or presentation in desktop Office, with AutoSave off. The saved bytes must match the run input exactly.",
                        "Open local app settings and enable Desktop Office focus. Read and accept the Office side-effect notice; the setting is initially off.",
                        "Open an atomic finding. Choose <strong>Bind</strong> for the current or baseline role. QC Tool offers a document only when exactly one eligible open document has byte-identical saved content.",
                        "Review the filename, role, folder label, and unsaved-change warning, then choose <strong>Confirm binding</strong>.",
                        "Choose <strong>Focus</strong>. QC Tool revalidates the process, window, path identity, saved bytes, AutoSave state, and target immediately before navigation.",
                        "Disable focus from local app settings to revoke every token, pending offer, acknowledgement, and live binding immediately.",
                    ]
                )
                _callout(
                    "What focus never does",
                    "It never opens, saves, recalculates, edits, closes, or quits an Office "
                    "document. Active content, Protected View, AutoSave, ambiguous copies, "
                    "multiple windows, unreadable identity, or changed saved bytes refuse the action.",
                    warning=True,
                )

            with _guide_section("mappings", "Excel to PowerPoint mappings"):
                _paragraph(
                    "Final-package QC extracts eligible figures from the current deck and ranks "
                    "candidate workbook cells by display match, label affinity, and numeric distance."
                )
                _table(
                    ["Stage", "What QC Tool does", "What the analyst must do"],
                    [
                        [
                            "Read the deck",
                            "Extracts figures from readable text, table cells, and visible native-chart labels.",
                            "Check coverage for pictures, embedded objects, or unavailable chart labels.",
                        ],
                        [
                            "Find candidates",
                            "Normalizes displayed forms such as $1.2M, 12%, and (1,234), then searches numeric workbook cells for display matches and near matches within 5%.",
                            "Treat the list as a search aid, not as proof of source.",
                        ],
                        [
                            "Rank candidates",
                            "Compares slide wording with the workbook member, sheet name, and nearest text to the left and above each cell. Exact display matches rank before near matches.",
                            "Confirm the business meaning, unit, workbook member, sheet, and cell.",
                        ],
                        [
                            "Recheck later",
                            "Stores the confirmed slide wording pattern and exact source cell, then compares them deterministically on later runs.",
                            "Re-confirm when wording, figure order, sheet layout, or source location changes.",
                        ],
                    ],
                )
                _list(
                    [
                        "Use a <strong>named profile</strong>; the default profile cannot persist confirmations.",
                        "Expand a suggestion and compare slide context, source labels, value, and match type.",
                        "Confirm only when the workbook cell is the intended source, not merely the same number.",
                        "For a multi-workbook package, confirm the nominated workbook member as well as the sheet and cell.",
                        "A <strong>near match</strong> is useful for locating drift but is not a successful reconciliation.",
                        "Coverage distinguishes eligible, mapped, verified, mismatched, unresolved, and unmapped figures.",
                        "Coverage also counts <strong>unavailable</strong> surfaces: a figure rendered into a picture or an embedded object is a real claim nobody here can read, so it is counted and named rather than dropped from the denominator.",
                        "Visible native chart labels use chart, series, and category anchors and can be confirmed like text or table figures.",
                        "Re-run Final-package QC after confirmations to verify persisted mappings independently.",
                        "QC Tool also detects semantically identical claims repeated across slides and reports mismatches (ppt-internal-repetition).",
                    ]
                )
                _callout(
                    "Suggestions are not proof",
                    "Real Excel and PowerPoint files often name the same measure differently: "
                    "abbreviations, renamed sheets, merged or multi-row headers, transposed "
                    "tables, hidden units, and month or region labels can all weaken the text "
                    "score. Formatting can differ too: the deck may show $1.2M while Excel "
                    "stores 1,249,999, or several unrelated cells may display the same rounded "
                    "number. QC Tool has no LLM or business dictionary, does not prove currency "
                    "or unit meaning from Excel formatting, and returns at most the configured "
                    "top candidate list. The intended source can therefore rank lower or be "
                    "absent. Confirm against the workbook's business logic; if no single saved "
                    "cell is the defensible source, leave the claim unmapped and document it.",
                    warning=True,
                )
                _callout(
                    "A clean mapping result still has a population",
                    "If any slide carries a rasterized or embedded surface, crosscheck coverage "
                    "degrades and names the affected slides. Read that before treating the deck "
                    "as fully reconciled.",
                    warning=True,
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
                _paragraph(
                    "History is also a shelf you curate. Tick one or more runs to "
                    "act on them together. Every run is stored in full, so the "
                    "list shows the storage each run occupies; when the total "
                    "grows large the app suggests exporting what you need and "
                    "archiving or deleting the rest."
                )
                _paragraph(
                    "Run IDs are permanent audit identities, not a count of stored runs. "
                    "Deleting run 12 does not make a later run reuse ID 12, so gaps are "
                    "normal and old reports, links, and Re-QC references never become ambiguous."
                )
                _paragraph(
                    "Dossier and recurrence: when a finding can be matched to a"
                    " stable anchor across an unbroken rerun chain, the UI shows a"
                    " per-run dossier (exact/changed/ambiguous/no stored observation)."
                )
                _paragraph(
                    "Recurrence is an advisory detection: it requires three consecutive"
                    " finalized matching observations, the same analyst disposition,"
                    " and at least two manual confirmations. The UI may nudge when"
                    " recurrence is eligible; it never auto-applies promotions."
                )
                _table(
                    ["Action", "Effect"],
                    [
                        [
                            "Archive",
                            "Hides the run from the default list and keeps every record, finding, comment, and report. Switch the shelf filter to Archived to see or restore them.",
                        ],
                        [
                            "Restore",
                            "Returns archived runs to the active list.",
                        ],
                        [
                            "Export selected (.zip)",
                            "One private zip with each selected run's Excel and HTML reports plus a manifest of run ids, modes, profiles, display filenames, and counts. No filesystem paths are included.",
                        ],
                        [
                            "Delete",
                            "Permanently removes the history records, analyst annotations, and stored report files after an explicit confirmation. Source deliverables are never touched.",
                        ],
                    ],
                )
                _callout(
                    "Archive before you delete",
                    "Deletion also removes the evidence a later audit or Re-QC delta "
                    "would rely on. Archiving retires a run from view without losing "
                    "anything, so prefer it unless the record must genuinely be gone.",
                    warning=True,
                )

            with _guide_section("privacy", "Privacy, sharing, and attestations"):
                _table(
                    ["Artifact", "Contains source information?", "Use"],
                    [
                        ["Numeric sanitize", "Yes", "Local testing only; not safe to share"],
                        ["Strict sanitize", "Transformed figures; date and category labels are kept only in structural form; package members use aliases", "Verified share candidate"],
                        ["Fingerprint", "No values, visible text, paths, filenames, or member IDs", "Safest structural evidence"],
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

# Semantic JSON review summary (pattern groups, stories, alignment trust) — v2
cadence-diff run --current-excel current.xlsx --json findings.json \\
    --json-review-summary

# Analyst acceptance threshold (visible-Info, never suppressed)
cadence-diff run --baseline-excel last.xlsx --current-excel this.xlsx \\
    --accept-absolute 1 --accept-percent 0.1

# Narrow the comparison scope (files still load fully; disclosed)
cadence-diff run --baseline-excel last.xlsx --current-excel this.xlsx \\
    --sheets "Dashboard,Data" --slides 1,3-5

# Overrides physical-size and formula-link safety refusals for this run
cadence-diff run --current-excel unusually-large.xlsx --allow-large-workbooks

# Forces full dependency indexing above the separate size policy that
# otherwise skips circular/formula/chart/PowerPoint-chart impacts
cadence-diff run --current-excel unusually-large.xlsx \\
    --allow-dependency-indexing

# Structural-only evidence
cadence-diff fingerprint current.xlsx -o current.fingerprint.json

# Structural-only package evidence
cadence-diff fingerprint --current-excel core.xlsx \\
    --current-workbook ops=ops.xlsx --current-ppt deck.pptx \\
    -o package.fingerprint.json

# Strict verified package redaction
cadence-diff sanitize-package --excel current.xlsx --ppt current.pptx \\
  --profile monthly --output-dir sanitized --forbid "Client Name"

# Multi-workbook package redaction
cadence-diff sanitize-package --excel core.xlsx --workbook ops=ops.xlsx \\
    --ppt deck.pptx --profile monthly --output-dir sanitized

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
                        [
                            "cadence-diff is not recognized on Windows",
                            "Activate the environment and run python -m qc_tool. Then create a Desktop shortcut from local app settings or python -m qc_tool shortcut install.",
                        ],
                        [
                            "I have only one current file",
                            "Choose Current-file preflight; no baseline or named profile is required.",
                        ],
                        [
                            "Desktop shortcut is stale",
                            "The Python environment moved or was removed. Run python -m qc_tool shortcut install from the environment you want to use.",
                        ],
                        [
                            "Desktop focus controls are missing or unavailable",
                            "Focus requires Windows and local network mode. Enable it in local app settings, open the exact saved file in desktop Office with AutoSave off, then use Bind and Confirm binding on an atomic finding.",
                        ],
                        ["Formula cache missing", "Open and recalculate in Excel, save, then rerun"],
                        [
                            "Workbook workload refused",
                            "Read the reason shown: it will name workbook size/cell counts, "
                            "XLSB cells/formulas/binary bytes, or formula-link complexity. "
                            "If you cannot confirm enough local memory and time, do not "
                            "override; ask a senior reviewer. Sheet scope is not a workaround "
                            "because the complete workbook still loads. When approved, enable "
                            "Override workbook workload refusals for that run and record why.",
                        ],
                        [
                            "Run cancelled",
                            "Wait for the next safe boundary. Partial reports are removed and no successful run is recorded.",
                        ],
                        [
                            "Run queued behind another run",
                            "One run executes at a time in an owned worker process. Queued "
                            "requests keep their position across a refresh; cancel one to "
                            "stop it before it starts.",
                        ],
                        [
                            "Run shown as orphaned",
                            "The server stopped before the run finished. Nothing is resumed "
                            "automatically \u2014 submit the run again.",
                        ],
                        [
                            "Password-protected run refused",
                            "Passwords are never stored in the queue. Start that run when "
                            "the worker is free instead of queueing it.",
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
                            "Review unsupported, invalid, or unparseable reference counts. A dynamic spill with no declared extent is unsupported, not guessed. Symbolic "
                            "aggregate references remain queryable and are disclosed separately.",
                        ],
                        [
                            "Dependency indexing skipped by size policy",
                            "Formula-cell count or projected dependencies crossed a documented "
                            "threshold; circular detection, formula/chart/PowerPoint-chart "
                            "impacts and their report evidence degrade together with this "
                            "reason. Enable Force full dependency indexing only after confirming "
                            "enough local memory and time -- distinct from Override workbook "
                            "workload refusals.",
                        ],
                        ["Re-QC file not found", "Select the renamed or replacement artifact again"],                        ["Profile will not save", "Correct the validation message or run qc-tool lint"],
                        ["Strict privacy verification fails", "Resolve each reported package issue or share only a fingerprint"],
                        ["Network page stops responding", "The temporary exposure may have expired; check qc-tool network status"],
                        ["Few or no findings", "Review coverage for unavailable/degraded checks before concluding clean"],
                        ["Finished for the day", "Use the power button in the header to stop the server; any queued or running job is recorded as unfinished and must be submitted again"],
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
                        "Any non-zero acceptance threshold or workbook workload override has a documented reason and approval.",
                        "Exports or attestations were generated only after analyst comments and overrides were saved.",
                        "A human still performs any visual or domain checks outside the reported coverage.",
                    ]
                )
                _callout(
                    "QC evidence supports judgment",
                    "The tool makes checks reproducible and reviewable; it does not replace the "
                    "analyst's responsibility for unsupported visual, business, or upstream-data assertions.",
                )

            with _guide_section("reference", "Advanced capability reference"):
                _paragraph(
                    "Use this section when a coverage detail names a specific parser or "
                    "representation. These capabilities are important, but they are not "
                    "prerequisites for a first run."
                )
                _table(
                    ["Capability", "Boundary"],
                    [
                        ["Defined names", "Workbook and sheet scope are compared separately; XLSB scope is unavailable."],
                        ["VBA", "Module inventory and source changes are compared without execution; source text never enters findings."],
                        ["Comments, Power Query, connections", "Content is compared locally; queries/connections are never executed and targets are represented by digests."],
                        ["Structured and spill references", "Proven extents are resolved; missing or ambiguous metadata degrades coverage instead of guessing."],
                        ["Circular references", "A compact formula-only graph reports exact cycles; unsupported or budgeted edges degrade coverage."],
                        ["PowerPoint images", "Embedded-byte changes are checked; rendered pixels, crop/layout equivalence, and OCR remain unavailable."],
                    ],
                )

            with _guide_section("glossary", "Glossary"):
                _paragraph(
                    "The vocabulary used across the results workbench, exports, "
                    "history, and JSON."
                )
                _table(
                    ["Term", "Meaning"],
                    [[term, meaning] for term, meaning in GLOSSARY],
                )

    ui.run_javascript(GUIDE_SCRIPT)
