"""Visual theme for the QC tool: quiet, print-inspired, instrument-panel.

Design rules: warm paper background, near-black ink, hairline borders
instead of shadows, system fonts only (the tool runs offline), monospace
for cell references and file names, and color reserved for severity —
the one thing that must pop. No gradients, no decoration.
"""

from collections.abc import Callable, Iterator
from contextlib import contextmanager

from nicegui import app, ui

from qc_tool import __version__

CSS = """
:root {
  --paper: #f4f2ed;
  --panel: #fdfcfa;
  --ink: #1d2025;
  --ink-soft: #4d5560;
  --line: #d8d4ca;
  --line-soft: #e7e4dc;
  --surface1: #f6f4ef;
  --surface2: #eceae3;
  --hit-bg: #fff6e2;
  --header-bg: #1d2025;
  --header-line: transparent;
  --btn-bg: #1d2025;
  --btn-fg: #f4f2ed;
  --btn-bg-hover: #30353c;
  --critical: #b42318;
  --warning: #b54708;
  --info: #175cd3;
  --expected: #067647;
  --font-sans: -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  --font-mono: ui-monospace, "Cascadia Mono", Consolas, Menlo, monospace;
  /* one fluid gutter for header, content, and footer */
  --gutter: clamp(0.75rem, 1.6vw, 2.25rem);
  /* type scale: metadata never drops below 12px */
  --fs-page: 1.45rem;
  --fs-run: 1.15rem;
  --fs-section: 0.9375rem;
  --fs-body: 0.875rem;
  --fs-meta: 0.75rem;
}
body.body--dark {
  --paper: #15171b;
  --panel: #1d2025;
  --ink: #e6e3dc;
  --ink-soft: #9aa1aa;
  --line: #363b42;
  --line-soft: #2a2e34;
  --surface1: #22262b;
  --surface2: #262b31;
  --hit-bg: #3d3419;
  --header-bg: #101216;
  --header-line: #2a2e34;
  --btn-bg: #e6e3dc;
  --btn-fg: #15171b;
  --btn-bg-hover: #cfccc4;
  --critical: #e0604d;
  --warning: #dd9a4a;
  --info: #6fa5ef;
  --expected: #56b184;
  /* Quasar tints flat buttons, dropdown highlights, chips etc. with primary;
     the light-mode ink primary is invisible on dark surfaces. NiceGUI sets
     --q-primary as an inline style on <body>, so !important is required. */
  --q-primary: #e6e3dc !important;
}
body { background: var(--paper); color: var(--ink); font-family: var(--font-sans); }
body.body--dark { background: var(--paper) !important; color: var(--ink) !important; }

/* header */
.appheader { background: var(--header-bg) !important; padding: 0;
  border-bottom: 1px solid var(--header-line); }
.appheader .inner { width: 100%; display: flex; align-items: baseline;
  gap: 1rem; padding: 0.65rem var(--gutter); }
.wordmark { color: #f4f2ed; font-weight: 650; font-size: 1.02rem;
  letter-spacing: 0.01em; }
.wordtag { font-family: var(--font-mono); font-size: var(--fs-meta); color: #b6bdc5;
  border: 1px solid #3a4048; border-radius: 3px; padding: 0.1rem 0.4rem; }
.headnav { margin-left: auto; display: flex; gap: 1.25rem; align-items: baseline; }
.headnav a { color: #c8cdd3; font-size: 0.82rem; text-decoration: none;
  padding-bottom: 2px; border-bottom: 1px solid transparent; }
.headnav a + a { margin-left: 1.25rem; }
.headnav a:hover { color: #fff; }
.headnav a.active { color: #fff; border-bottom-color: #8d959e; }
/* #969ea8 on the ink header is 5.9:1 — the old #6d757e was 3.7:1 */
.headnote { font-family: var(--font-mono); font-size: var(--fs-meta); color: #969ea8; }
/* the toggle sits on the always-dark band: force light icon over Quasar's
   text-primary (which is near-black and vanishes against the header) */
.themebtn { border: 1px solid #3a4048; }
.themebtn, .themebtn .q-icon { color: #cfd4da !important; }
.themebtn:hover { border-color: #8d959e; }
.themebtn:hover, .themebtn:hover .q-icon { color: #ffffff !important; }
.quitbtn:hover { border-color: #e0604d; }
.quitbtn:hover, .quitbtn:hover .q-icon { color: #e0604d !important; }

/* the server has stopped and the socket is gone: plain, theme-independent,
   and layered over the app so NiceGUI's disconnect handler keeps its DOM */
.stopped-page { position: fixed; inset: 0; z-index: 99999;
  font-family: -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
  display: flex; flex-direction: column; gap: 0.4rem;
  align-items: center; justify-content: center; background: #f4f2ed;
  color: #1d2025; text-align: center; padding: 2rem; }
.stopped-page .t { font-size: 1.25rem; font-weight: 650; }
.stopped-page .d { font-size: 0.875rem; color: #4d5560; }

/* layout: fluid, never a fixed pixel column — the window is the canvas and
   prose keeps its own reading measure, so any display scale stays usable */
.page-wrap { width: 100%; max-width: none; margin: 0;
  padding: 1.6rem var(--gutter) 3rem; gap: 0; }
.appfooter { width: 100%; max-width: none; margin: 0;
  border-top: 1px solid var(--line); padding: 0.7rem var(--gutter) 2rem;
  font-family: var(--font-mono); font-size: var(--fs-meta); color: var(--ink-soft);
  display: flex; align-items: baseline; justify-content: space-between;
  gap: 0.75rem 1.5rem; flex-wrap: wrap; }
/* muted like the rest of the footer — subtle by placement, not by low contrast */
.colophon { white-space: nowrap; }

/* section headings */
.section { display: flex; align-items: center; gap: 0.75rem;
  margin: 1.9rem 0 0.9rem; width: 100%; }
.section:first-child { margin-top: 0.4rem; }
.section .kicker { font-size: var(--fs-section); letter-spacing: 0.07em;
  font-weight: 650; color: var(--ink); text-transform: uppercase;
  white-space: nowrap; }
.section .rule { flex: 1; border-top: 1px solid var(--line); }

.pagetitle { font-size: var(--fs-page); font-weight: 680; line-height: 1.2; }
.hint { font-size: var(--fs-meta); color: var(--ink-soft); }
/* `.note` was referenced by the app but never defined, so supporting copy
   rendered at default browser size with no muted color. */
.note { font-size: var(--fs-meta); color: var(--ink-soft); line-height: 1.5;
  max-width: 62rem; }
.lede { color: var(--ink-soft); font-size: var(--fs-body); max-width: 56rem; }
.mode-select { width: 100%; max-width: 52rem; border: 1px solid var(--line);
  border-radius: 5px; overflow: hidden; }
.mode-select .q-btn { min-height: 2.5rem; background: var(--panel); }
.modecopy { color: var(--ink-soft); font-size: 0.8125rem; margin-bottom: 0.55rem; }
.guide-jump { color: var(--info) !important; font-size: var(--fs-meta);
  text-decoration: none; margin: -0.35rem 0 0.55rem; }
.guide-jump:hover { text-decoration: underline; }

/* guide */
.guide-page-title { font-size: var(--fs-page); font-weight: 680; line-height: 1.2; }
.guide-page-lede { color: var(--ink-soft); font-size: var(--fs-body); max-width: 74rem;
  margin-top: 0.35rem; }
.guide-search { display: flex; align-items: center; gap: 0.7rem; width: 100%;
  margin-top: 1rem; }
.guide-search input { flex: 1; max-width: 26rem; font-family: var(--font-sans);
  font-size: var(--fs-body); color: var(--ink); background: var(--panel);
  border: 1px solid var(--line); border-radius: 5px; padding: 0.4rem 0.6rem; }
.guide-search input:focus { outline: 2px solid var(--info); outline-offset: -1px; }
.guide-search-count { font-size: var(--fs-meta); color: var(--ink-soft); }
.guide-tasks { display: flex; align-items: baseline; gap: 0.9rem; flex-wrap: wrap;
  width: 100%; margin-top: 0.8rem; }
.guide-tasks-title { font-size: var(--fs-meta); letter-spacing: 0.08em;
  text-transform: uppercase; font-weight: 700; color: var(--ink-soft); }
.guide-task-link { color: var(--info) !important; font-size: var(--fs-meta);
  text-decoration: none; }
.guide-task-link:hover { text-decoration: underline; }
.guide-layout { display: grid; grid-template-columns: 13rem minmax(0, 1fr);
  gap: 2.5rem; width: 100%; align-items: start; margin-top: 1.5rem; }
.guide-toc { position: sticky; top: 1rem; border-left: 2px solid var(--line);
  padding-left: 0.85rem; gap: 0.25rem; }
.guide-toc-title { color: var(--ink-soft); font-size: var(--fs-meta);
  letter-spacing: 0.1em; text-transform: uppercase; font-weight: 700;
  margin-bottom: 0.35rem; }
.guide-toc-link { display: block; color: var(--ink-soft) !important;
  font-size: var(--fs-meta); line-height: 1.4; text-decoration: none;
  padding: 0.18rem 0; }
.guide-toc-link:hover { color: var(--ink) !important; }
/* one measure for the whole column: prose, lists, callouts, tables, and code
   then share exactly the same left and right edge */
.guide-content { min-width: 0; max-width: 74rem; }
.guide-section { scroll-margin-top: 1rem; padding: 0 0 1.8rem;
  border-bottom: 1px solid var(--line); margin-bottom: 1.8rem; }
.guide-section:last-child { border-bottom: 0; }
.guide-section-title { font-size: 1.12rem; font-weight: 680; margin-bottom: 0.65rem;
  cursor: pointer; }
.guide-section-title:focus-visible { outline: 2px solid var(--info);
  outline-offset: 2px; }
.guide-section-title::after { content: "-"; color: var(--ink-soft);
  font-weight: 500; margin-left: 0.5rem; }
.guide-section.collapsed { padding-bottom: 0.9rem; margin-bottom: 0.9rem; }
.guide-section.collapsed > *:not(.guide-section-title) { display: none; }
.guide-section.collapsed .guide-section-title { margin-bottom: 0; }
.guide-section.collapsed .guide-section-title::after { content: "+"; }
#guide-toc-more { display: none; }
.guide-copy { color: var(--ink); font-size: var(--fs-body); line-height: 1.62;
  white-space: normal; }
.guide-list { color: var(--ink); font-size: var(--fs-body); line-height: 1.55;
  padding-left: 1.2rem; margin: 0.65rem 0; }
.guide-list li { margin: 0.3rem 0; }
.guide-list code, .guide-copy code { font-family: var(--font-mono);
  background: var(--surface2); border: 1px solid var(--line-soft);
  border-radius: 3px; padding: 0.05rem 0.25rem; font-size: var(--fs-meta); }
.guide-code { background: var(--surface1); border: 1px solid var(--line);
  border-radius: 5px; padding: 0.85rem 1rem; overflow-x: auto; color: var(--ink);
  font-family: var(--font-mono); font-size: var(--fs-meta); line-height: 1.55;
  margin: 0.8rem 0; white-space: pre; }
.guide-table-wrap { width: 100%; overflow-x: auto; margin: 0.75rem 0; }
.guide-table { width: 100%; min-width: 38rem; border-collapse: collapse;
  font-size: 0.8125rem; }
.guide-table th { background: var(--surface2); color: var(--ink); text-align: left;
  font-size: var(--fs-meta); letter-spacing: 0.06em; text-transform: uppercase;
  padding: 0.45rem 0.55rem; border: 1px solid var(--line); }
.guide-table td { color: var(--ink); padding: 0.45rem 0.55rem;
  border: 1px solid var(--line); vertical-align: top; line-height: 1.4; }
.guide-callout { border-left: 3px solid var(--info); background: var(--surface1);
  padding: 0.65rem 0.8rem; margin: 0.8rem 0; }
.guide-callout-warning { border-left-color: var(--warning); }
.guide-callout-title { font-size: 0.8125rem; font-weight: 700; }
.guide-callout-copy { color: var(--ink-soft); font-size: 0.8125rem;
  line-height: 1.5; margin-top: 0.15rem; white-space: normal; }

/* panels + quasar surfaces */
.q-card { background: var(--panel); border: 1px solid var(--line);
  border-radius: 6px; box-shadow: none; }
.q-uploader { background: var(--panel); border: 1px solid var(--line);
  border-radius: 6px; box-shadow: none; width: 100%; }
.q-uploader__header { background: var(--surface2); color: var(--ink); }
.q-uploader__title { font-size: 0.8125rem; font-weight: 600; }
.q-uploader__subtitle { font-family: var(--font-mono); font-size: var(--fs-meta); }
.q-uploader__list { min-height: 1.4rem; padding: 0.35rem 0.6rem; }
.q-expansion-item { background: var(--panel); border: 1px solid var(--line);
  border-radius: 6px; }
.profile-section { margin-bottom: 0.4rem; }
.profile-section .profile-section { border-color: var(--line-soft);
  margin: 0.25rem 0; }
.profile-section .q-item { min-height: 2.35rem; }
.profile-section .q-expansion-item__content { padding: 0.35rem 0.55rem 0.55rem; }
.profile-icon-button { width: 2rem; height: 2rem; flex: 0 0 2rem; }
.profile-yaml textarea { font-family: var(--font-mono); font-size: var(--fs-meta);
  line-height: 1.45; }
.profile-editor-actions { position: sticky; bottom: -1px; z-index: 2;
  background: var(--panel); border-top: 1px solid var(--line);
  padding: 0.65rem 0 0.2rem; width: 100%; }
.preline { white-space: pre-line; overflow-wrap: anywhere; }

/* inputs: one panel per cycle so baseline and current can never be scanned
   as a single left-to-right list of four look-alike slots */
.rolegroups { display: grid; grid-template-columns: repeat(auto-fit, minmax(30rem, 1fr));
  gap: 1rem; width: 100%; margin-top: 0.6rem; }
.rolegroup { border: 1px solid var(--line); border-radius: 6px;
  padding: 0.7rem 0.9rem 0.9rem; background: var(--panel); min-width: 0; }
.rolegroup-baseline { background: var(--surface1); }
.rolegroup-current { border-left: 3px solid var(--ink); }
.rolegrouphead { align-items: baseline; gap: 0.6rem; width: 100%; flex-wrap: wrap;
  margin-bottom: 0.6rem; padding-bottom: 0.45rem;
  border-bottom: 1px solid var(--line-soft); }
.rolegrouptitle { font-size: var(--fs-section); font-weight: 680; }
.rolegroupnote { font-size: var(--fs-meta); color: var(--ink-soft); }

.upgrid { display: grid; grid-template-columns: repeat(auto-fit, minmax(22rem, 1fr));
  gap: 1rem 1.5rem; width: 100%; }
.rolelabel { font-size: var(--fs-meta); letter-spacing: 0.09em; font-weight: 650;
  text-transform: uppercase; color: var(--ink-soft); margin-bottom: 0.25rem; }
.filestate { font-family: var(--font-mono); font-size: var(--fs-meta);
  color: var(--ink-soft); margin-top: 0.3rem; }
.filestate.ok { color: var(--expected); }
.filestate.err { color: var(--critical); font-weight: 650; }

/* file-role rows */
.rolerow { gap: 0; width: 100%; min-width: 0; }
.rolehead { align-items: baseline; gap: 0.5rem; width: 100%; flex-wrap: nowrap; }
.rolebadge { font-size: var(--fs-meta); letter-spacing: 0.05em;
  text-transform: uppercase; color: var(--ink-soft); border: 1px solid var(--line);
  border-radius: 3px; padding: 0 0.35rem; white-space: nowrap; }
.rolebadge.satisfied { color: var(--expected); border-color: var(--expected); }
.rolefoot { align-items: center; gap: 0.5rem; width: 100%; min-width: 0; }
.rolefoot .filestate { flex: 1; min-width: 0; overflow-wrap: anywhere; }
/* the uploader duplicates the filename we already print in .filestate */
.rolerow .q-uploader__list { display: none; }
/* "0.0B / 0.00%" is implementation-facing; .filestate carries name and size */
.rolerow .q-uploader__subtitle { display: none; }
.rolerow .q-uploader__header { min-height: 2.1rem; }
.linkbtn { color: var(--info) !important; font-size: var(--fs-meta);
  padding: 0 0.35rem; min-height: 1.4rem; }
.linkbtn .q-btn__content { color: var(--info) !important; }

/* run policy */
.policyrow { display: flex; align-items: flex-end; gap: 1rem; flex-wrap: wrap;
  width: 100%; }
.policytokens { display: flex; gap: 0.4rem; flex-wrap: wrap; width: 100%;
  margin-top: 0.5rem; }
.policytoken { font-size: var(--fs-meta); color: var(--ink); background: var(--surface1);
  border: 1px solid var(--line); border-radius: 3px; padding: 0.15rem 0.5rem; }

/* sticky run readiness bar */
.readybar { position: sticky; bottom: 0; z-index: 5; width: 100%;
  margin-top: 1.5rem; background: var(--panel); border: 1px solid var(--line);
  border-radius: 6px; padding: 0.65rem 0.9rem; display: flex;
  align-items: center; gap: 0.9rem; flex-wrap: wrap; }
.readysummary { display: flex; flex-direction: column; gap: 0.15rem;
  flex: 1; min-width: 12rem; }
.readysummary .r1 { font-size: var(--fs-body); font-weight: 650; }
.readysummary .r2 { font-size: var(--fs-meta); color: var(--ink-soft);
  overflow-wrap: anywhere; }
.readysummary .r2.caution { color: var(--warning); }
.readysummary .blocked { color: var(--critical); }
.readyactions { display: flex; align-items: center; gap: 0.6rem; flex-wrap: wrap; }
.readyqueue { width: 100%; display: flex; align-items: center; gap: 0.6rem;
  flex-wrap: wrap; border-top: 1px solid var(--line-soft); padding-top: 0.5rem; }

/* buttons */
.q-btn { border-radius: 4px; text-transform: none; font-weight: 550;
  letter-spacing: 0.01em; box-shadow: none; }
/* content-level color rules outrank Quasar's .text-primary on flat buttons */
.runbtn { background: var(--btn-bg) !important;
  padding: 0.55rem 2.2rem; font-size: 0.92rem; }
.runbtn, .runbtn .q-btn__content { color: var(--btn-fg) !important; }
.runbtn:hover { background: var(--btn-bg-hover) !important; }
.ghostbtn { border: 1px solid var(--line); background: transparent; }
.ghostbtn, .ghostbtn .q-btn__content { color: var(--ink) !important; }

/* Quasar paints the pressed toggle button bg-primary + text-white, and
   --q-primary is light in dark mode, so white on light would be unreadable. */
.q-btn-toggle .q-btn, .q-btn-toggle .q-btn .q-btn__content {
  color: var(--ink) !important; }
.q-btn-toggle .q-btn[aria-pressed="true"] { background: var(--btn-bg) !important; }
.q-btn-toggle .q-btn[aria-pressed="true"],
.q-btn-toggle .q-btn[aria-pressed="true"] .q-btn__content {
  color: var(--btn-fg) !important; }

/* quasar form surfaces in dark mode */
body.body--dark .q-field--outlined .q-field__control { background: var(--panel); }
body.body--dark .q-field__native, body.body--dark .q-field__input,
body.body--dark .q-field__label { color: var(--ink); }
body.body--dark .q-chip { background: var(--surface2); color: var(--ink); }
body.body--dark .q-select__dropdown-icon { color: var(--ink-soft); }

/* severity stat strip */
.statstrip { display: flex; gap: 0.7rem; width: 100%; flex-wrap: wrap; }
.stat { background: var(--panel); border: 1px solid var(--line);
  border-top-width: 3px; border-radius: 4px; padding: 0.35rem 0.85rem 0.4rem;
  min-width: 7.5rem; gap: 0; }
.stat .n { font-size: 1.35rem; font-weight: 650; line-height: 1.15;
  font-variant-numeric: tabular-nums; }
.stat .l { font-size: var(--fs-meta); letter-spacing: 0.07em; text-transform: uppercase;
  color: var(--ink-soft); }
.stat .a { font-family: var(--font-mono); font-size: var(--fs-meta);
  color: var(--ink-soft); margin-top: 0.15rem; }
.stat-critical { border-top-color: var(--critical); }
.stat-warning { border-top-color: var(--warning); }
.stat-info { border-top-color: var(--info); }
.stat-expected { border-top-color: var(--expected); }

/* findings table */
.findings-table { border: 1px solid var(--line); border-radius: 6px;
  box-shadow: none; background: var(--panel); width: 100%; }
.findings-table .q-table__top { display: none; }
.findings-table thead th { background: var(--surface2); color: var(--ink);
  font-size: var(--fs-meta); letter-spacing: 0.07em; text-transform: uppercase;
  font-weight: 650; border-bottom: 1px solid var(--line); }
.findings-table td { border-color: var(--line-soft) !important;
  font-size: 0.8125rem; vertical-align: top; white-space: normal;
  overflow-wrap: break-word; }
.findings-table td:first-child { white-space: nowrap; }
/* let the message column absorb the slack instead of wrapping every word */
.findings-table:not(.history-table) th:last-child,
.findings-table:not(.history-table) td:last-child { width: 60%; min-width: 15rem; }
/* auto layout treats a cell width as a hint, so the review item column only
   holds its share under fixed layout; metadata keeps a title tooltip */
.review-groups-table .q-table { table-layout: fixed; }
.review-groups-table th:nth-child(1), .review-groups-table td:nth-child(1) { width: 8%; }
.review-groups-table th:nth-child(2), .review-groups-table td:nth-child(2) { width: 12%; }
.review-groups-table th:nth-child(3), .review-groups-table td:nth-child(3) { width: 16%; }
.review-groups-table th:nth-child(4), .review-groups-table td:nth-child(4) { width: 4%; }
.review-groups-table th:nth-child(5), .review-groups-table td:nth-child(5) {
  width: 60%; min-width: 0; }
/* fixed layout leaves no slack, so headers must wrap instead of colliding */
.findings-table thead th { white-space: normal; }
.findings-table .mono { font-family: var(--font-mono); font-size: var(--fs-meta); }
.review-groups-table tbody tr { cursor: pointer; }
.review-groups-table tbody tr:focus-visible,
.members-table tbody tr:focus-visible { outline: 2px solid var(--ink);
  outline-offset: -2px; }
.review-groups-table tbody tr.selrow td { background: var(--surface1); }
.review-groups-table tbody tr.selrow td:first-child {
  box-shadow: inset 3px 0 0 var(--ink); }
.review-toggle { border: 1px solid var(--line); border-radius: 4px; }
.review-groups-table .groupcount { font-family: var(--font-mono);
  font-variant-numeric: tabular-nums; text-align: right; }
.capbadge { color: var(--warning); font-size: var(--fs-meta); display: block;
  margin-top: 0.15rem; }
.coverage-table { width: 100%; border: 1px solid var(--line); border-radius: 6px;
  background: var(--panel); }
.coverage-table thead th { background: var(--surface2); color: var(--ink);
  font-size: var(--fs-meta); letter-spacing: 0.07em; text-transform: uppercase; }
.coverage-table td { border-color: var(--line-soft) !important; font-size: 0.8125rem; }
.mappingstats { display: flex; gap: 0.7rem; flex-wrap: wrap; width: 100%;
  margin-bottom: 0.6rem; }
.mappingstat { min-width: 6.8rem; gap: 0; border-left: 2px solid var(--line);
  padding: 0.2rem 0.65rem; }
.mappingstat .n { font-size: 1.25rem; font-weight: 650; line-height: 1.1; }
.mappingstat .l { color: var(--ink-soft); font-size: var(--fs-meta);
  letter-spacing: 0.07em; text-transform: uppercase; }
.mappingitem { width: 100%; }
.mappingcontext { font-family: var(--font-mono); font-size: var(--fs-meta);
  color: var(--ink-soft); padding: 0.25rem 0.75rem 0.5rem; }
.candidate-row { width: 100%; align-items: center; gap: 0.8rem;
  border-top: 1px solid var(--line-soft); padding: 0.35rem 0.75rem; }
.candidate-ref { font-family: var(--font-mono); font-size: var(--fs-meta);
  min-width: 12rem; }
.candidate-value { font-family: var(--font-mono); font-size: var(--fs-meta); }
.candidate-match, .candidate-near { margin-left: auto; font-size: var(--fs-meta);
  text-transform: uppercase; letter-spacing: 0.06em; }
.candidate-match { color: var(--expected); }
.candidate-near { color: var(--warning); }
.detailrow td { background: var(--surface1); }
.detailgrid { display: grid; grid-template-columns: 6.5rem 1fr;
  gap: 0.2rem 1rem; padding: 0.25rem 0 0.35rem; max-width: 60rem; }
.detailgrid .dk { color: var(--ink-soft); text-transform: uppercase;
  font-size: var(--fs-meta); letter-spacing: 0.07em; padding-top: 3px; }
.detailgrid .dv { font-size: 0.8125rem; overflow-wrap: anywhere; }
.detailgrid .dv.mono { font-family: var(--font-mono); font-size: var(--fs-meta); }
.ctxpair { display: flex; gap: 1.5rem; flex-wrap: wrap; }
/* wide evidence grids scroll inside their own block, never the whole panel */
.ctxblock { max-width: 100%; overflow-x: auto; margin-top: 0.45rem; }
.ctxlabel { font-size: var(--fs-meta); letter-spacing: 0.07em; text-transform: uppercase;
  color: var(--ink-soft); margin-bottom: 2px; }
.ctxgrid { border-collapse: collapse; font-family: var(--font-mono);
  font-size: var(--fs-meta); }
.ctxgrid th { background: var(--surface2); color: var(--ink-soft); font-weight: 500;
  padding: 1px 6px; border: 1px solid var(--line-soft); font-size: var(--fs-meta); }
.ctxgrid td { border: 1px solid var(--line-soft); padding: 1px 6px;
  background: var(--panel); max-width: 9rem; overflow: hidden; text-overflow: ellipsis;
  white-space: nowrap; }
.ctxgrid td.hit { outline: 2px solid var(--warning); outline-offset: -2px;
  background: var(--hit-bg); font-weight: 650; }
.deltaline { font-family: var(--font-mono); font-size: 0.8125rem;
  color: var(--ink-soft); margin-top: 0.5rem; }
.deltaline .good { color: var(--expected); font-weight: 650; }
.deltaline .bad { color: var(--critical); font-weight: 650; }
.annotrow { display: flex; gap: 0.6rem; align-items: center; max-width: 44rem;
  margin-top: 0.15rem; }
.annotsev { min-width: 9.5rem; }
.annotcomment { flex: 1; }
.overridden { font-size: var(--fs-meta); letter-spacing: 0.06em; color: var(--info);
  text-transform: uppercase; margin-left: 0.4rem; }
/* An analyst note without a severity change still counts as a decision. */
.reviewed { font-size: var(--fs-meta); letter-spacing: 0.06em; color: var(--ink-soft);
  text-transform: uppercase; margin-left: 0.4rem; white-space: nowrap; }
.rerunbanner { border-left: 3px solid var(--info); background: var(--panel);
  border-top: 1px solid var(--line); border-right: 1px solid var(--line);
  border-bottom: 1px solid var(--line); border-radius: 0 4px 4px 0;
  padding: 0.5rem 0.85rem; font-size: var(--fs-body); width: 100%; }
.sevdot { display: inline-block; width: 8px; height: 8px; border-radius: 50%;
  margin-right: 0.45rem; vertical-align: baseline; }
.sev-critical { background: var(--critical); }
.sev-warning { background: var(--warning); }
.sev-info { background: var(--info); }
.sev-expected { background: var(--expected); }
.sevtext { font-size: var(--fs-meta); letter-spacing: 0.04em; }

/* run outcome + capability status */
.statusrow { display: flex; align-items: center; gap: 0.6rem; flex-wrap: wrap;
  width: 100%; margin: 0.2rem 0 0.6rem; }
.statuschip { display: inline-flex; align-items: baseline; gap: 0.45rem;
  border: 1px solid var(--line); border-left-width: 3px; border-radius: 0 4px 4px 0;
  background: var(--panel); padding: 0.32rem 0.7rem; }
.statuschip .t { font-size: var(--fs-body); font-weight: 650; }
.statuschip .d { font-size: var(--fs-meta); color: var(--ink-soft); }
.status-attention { border-left-color: var(--critical); }
.status-attention .t { color: var(--critical); }
.status-limited { border-left-color: var(--warning); }
.status-limited .t { color: var(--warning); }
.status-ok { border-left-color: var(--expected); }
.status-ok .t { color: var(--expected); }
.status-neutral { border-left-color: var(--line); }

/* results workbench */
.runheader { width: 100%; display: flex; flex-direction: column; gap: 0.55rem;
  border-bottom: 1px solid var(--line); padding-bottom: 0.9rem;
  margin-bottom: 0.9rem; }
.runtitle { font-size: var(--fs-run); font-weight: 680; line-height: 1.25; }
.runhead { font-weight: 650; font-size: var(--fs-run); }
.runmeta { font-family: var(--font-mono); font-size: var(--fs-meta);
  color: var(--ink-soft); overflow-wrap: anywhere; }
.headeractions { display: flex; align-items: center; gap: 0.6rem; flex-wrap: wrap;
  width: 100%; }
.storyline { font-size: var(--fs-meta); color: var(--ink-soft); width: 100%;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.resulttabs { width: 100%; border-bottom: 1px solid var(--line); color: var(--ink); }
.resulttabs .q-tab { min-height: 2.3rem; }
.resulttabs .q-tab__label { font-size: var(--fs-body); font-weight: 600; }
.resultpanels { width: 100%; background: transparent !important; }
.resultpanels .q-tab-panel { padding: 0.9rem 0 0; }
.reviewtoolbar { display: flex; align-items: center; gap: 0.7rem; flex-wrap: wrap;
  width: 100%; margin-bottom: 0.7rem; }
.reviewsplit { display: grid;
  grid-template-columns: minmax(0, 1fr) clamp(24rem, 38%, 48rem);
  gap: 1rem; align-items: start; width: 100%; }
.detailpanel { border: 1px solid var(--line); border-radius: 6px;
  background: var(--panel); padding: 0.75rem 0.85rem; min-height: 6rem;
  position: sticky; top: 0.75rem; max-height: calc(100vh - 6rem);
  overflow-y: auto; overflow-x: hidden; }
.detailpanel:empty::before { content: "Select a review item to see its evidence.";
  font-size: var(--fs-meta); color: var(--ink-soft); }
.detail-id { font-family: var(--font-mono); font-size: var(--fs-meta);
  color: var(--ink-soft); }
.detail-msg { font-size: var(--fs-body); font-weight: 650; line-height: 1.35;
  margin: 0.2rem 0 0.35rem; overflow-wrap: anywhere; }
.detailpanel .detailgrid { grid-template-columns: 8.5rem minmax(0, 1fr);
  max-width: 100%; }
/* a long axis (impacts on a wide rollout) must not bury the grids below it */
.detailpanel .detailgrid .dv { max-height: 11rem; overflow-y: auto; }
.covlimit { border-left: 3px solid var(--warning); background: var(--panel);
  border-top: 1px solid var(--line); border-right: 1px solid var(--line);
  border-bottom: 1px solid var(--line); border-radius: 0 4px 4px 0;
  padding: 0.45rem 0.8rem; width: 100%; margin-top: 0.5rem; }
.covlimit .t { font-size: var(--fs-body); font-weight: 650; }
.covlimit .d { font-size: var(--fs-meta); color: var(--ink-soft); }
.story-item { width: 100%; }

/* affected-findings dialog: the same master-detail surface, never a wide
   expandable row that pushes the table into a horizontal scroll */
.memberscard { width: min(104rem, 96vw); max-width: 96vw; max-height: 88vh;
  display: flex; flex-direction: column; gap: 0.5rem; }
.memberssplit { grid-template-columns: minmax(0, 1fr) clamp(24rem, 42%, 46rem);
  flex: 1; min-height: 0; }
.memberssplit .members-table { max-height: 66vh; overflow: auto; }
.memberssplit .detailpanel { position: static; max-height: 66vh; }
.members-table td:first-child { white-space: normal; }

.notecard { border-left: 3px solid var(--critical); background: var(--panel);
  border-top: 1px solid var(--line); border-right: 1px solid var(--line);
  border-bottom: 1px solid var(--line); border-radius: 0 4px 4px 0;
  padding: 0.5rem 0.85rem; font-size: var(--fs-body); color: var(--ink-soft);
  width: 100%; }
.exposurebanner { border: 2px solid var(--critical); background: var(--panel);
  color: var(--critical); font-weight: 700; padding: 0.65rem 0.85rem;
  border-radius: 4px; width: 100%; margin-bottom: 0.8rem; }

/* history */
.runcard { width: 100%; padding: 0.9rem 1.1rem; gap: 0.15rem; }
.runcounts { display: flex; gap: 0.9rem; font-size: var(--fs-meta); margin-top: 0.2rem;
  font-variant-numeric: tabular-nums; }
.historytoolbar { display: flex; align-items: center; gap: 0.6rem; flex-wrap: wrap;
  width: 100%; margin: 0.9rem 0 0.8rem; position: sticky; top: 0; z-index: 4;
  background: var(--paper); padding: 0.35rem 0; }
.history-table td { white-space: nowrap; }
.history-table .c-files { font-family: var(--font-mono); max-width: 26rem;
  overflow: hidden; text-overflow: ellipsis; }
.history-table tr.archivedrow td { opacity: 0.62; }
.archivedtag { font-size: var(--fs-meta); color: var(--ink-soft);
  border: 1px solid var(--line); border-radius: 3px; padding: 0 0.3rem;
  margin-left: 0.4rem; }
.bulkbar { display: flex; align-items: center; gap: 0.6rem; flex-wrap: wrap;
  width: 100%; margin-bottom: 0.7rem; padding: 0.5rem 0.75rem;
  border: 1px solid var(--line); border-left: 3px solid var(--ink);
  border-radius: 0 5px 5px 0; background: var(--panel); }
.bulkcount { font-size: var(--fs-body); font-weight: 650; margin-right: 0.4rem; }
.dangerbtn, .dangerbtn .q-btn__content { color: var(--critical) !important; }
.dangerbtn { border-color: var(--critical); }
.decisioncount { font-variant-numeric: tabular-nums; margin-right: 0.55rem; }
.atomicnote { color: var(--ink-soft); font-size: var(--fs-meta); }
.caplimited { color: var(--warning); font-weight: 650; }
.capfull { color: var(--ink-soft); }

.reviewclass-inline { display: none; color: var(--ink-soft);
  font-family: var(--font-mono); font-size: var(--fs-meta); margin-right: 0.45rem; }

@media (max-width: 1350px) {
  .review-groups-table th:nth-child(2),
  .review-groups-table td:nth-child(2) { display: none; }
  .review-groups-table th:nth-child(1), .review-groups-table td:nth-child(1) { width: 11%; }
  .review-groups-table th:nth-child(3), .review-groups-table td:nth-child(3) { width: 20%; }
  .review-groups-table th:nth-child(4), .review-groups-table td:nth-child(4) { width: 5%; }
  .review-groups-table th:nth-child(5), .review-groups-table td:nth-child(5) { width: 64%; }
  .reviewclass-inline { display: inline; }
  .review-groups-table .sevtext { white-space: nowrap; }
}

@media (max-width: 640px) {
  .appheader .inner { align-items: center; flex-wrap: wrap; gap: 0.45rem 0.7rem;
    padding: 0.55rem 0.75rem; }
  .wordmark, .wordtag { white-space: nowrap; }
  .headnav { order: 3; margin-left: 0; width: 100%; gap: 0.75rem;
    justify-content: space-between; align-items: center; }
  .headnav a + a { margin-left: 0.7rem; }
  .page-wrap { padding: 1.1rem 0.75rem 2rem; min-width: 0; }
  .section { align-items: flex-start; min-width: 0; gap: 0.45rem; }
  .section .kicker { white-space: normal; min-width: 0; overflow-wrap: anywhere;
    line-height: 1.45; }
  .section .rule { min-width: 1rem; margin-top: 0.45rem; }
  .mode-select { display: grid; grid-template-columns: 1fr; }
  .guide-layout { grid-template-columns: 1fr; gap: 1.2rem; }
  .guide-toc { position: static; display: grid; grid-template-columns: repeat(2, 1fr);
    border-left: 0; border-bottom: 1px solid var(--line); padding: 0 0 0.8rem; }
  .guide-toc-title { grid-column: 1 / -1; }
  /* six topics, then an explicit toggle instead of a long list */
  .guide-toc:not(.expanded) .guide-toc-link:nth-of-type(n + 7) { display: none; }
  #guide-toc-more { display: block; grid-column: 1 / -1; justify-self: start;
    margin-top: 0.4rem; background: transparent; border: 1px solid var(--line);
    border-radius: 4px; color: var(--ink); font-size: var(--fs-meta);
    padding: 0.2rem 0.6rem; }
  .guide-search input { max-width: none; }
  .guide-section { scroll-margin-top: 0.5rem; }
  .upgrid { grid-template-columns: 1fr; }
  .readybar { padding: 0.55rem 0.7rem; gap: 0.5rem; }
  .readyactions { width: 100%; }
  .readyactions .runbtn { flex: 1; padding: 0.55rem 1rem; }
  .statstrip { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .stat { min-width: 0; }
  .coverage-table, .findings-table { min-width: 0; max-width: 100%; }
  /* the detail panel carries the class on small screens */
  .review-groups-table thead th:nth-child(2),
  .findings-table .c-class { display: none; }
  .findings-table:not(.history-table) th:last-child,
  .findings-table:not(.history-table) td:last-child { min-width: 0; }
  .reviewsplit { grid-template-columns: 1fr; }
  .detailpanel { position: static; max-height: none; }
  .resulttabs .q-tab__label { font-size: var(--fs-meta); }
  .mappingstats { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); }
  .mappingstat { min-width: 0; }
  .candidate-row { flex-wrap: wrap; }
  .candidate-ref { min-width: 0; flex: 1 1 10rem; }
  .candidate-match, .candidate-near { margin-left: 0; }
  .detailgrid { grid-template-columns: 5rem minmax(0, 1fr); }
  .annotrow { flex-direction: column; align-items: stretch; }
  .annotsev { width: 100%; }
  .runcounts { flex-wrap: wrap; }
}
"""

SEVERITY_CELL_SLOT = """
<q-td :props="props">
  <span :class="'sevdot sev-' + props.value"></span><span class="sevtext">{{ props.value }}</span>
</q-td>
"""

HISTORY_BODY_SLOT = """
<q-tr :props="props" :class="{archivedrow: props.row.archived}">
  <q-td auto-width>
    <q-checkbox dense :model-value="props.row.sel"
      :aria-label="'Select run ' + props.row.id"
      @update:model-value="v => $parent.$emit('select', {id: props.row.id, value: v})" />
  </q-td>
  <q-td key="id" :props="props" class="mono">#{{ props.row.id }}
    <span v-if="props.row.archived" class="archivedtag">archived</span>
    <span v-if="props.row.finalized" class="reviewed">finalized</span>
  </q-td>
  <q-td key="when" :props="props">
    <span :title="props.row.started">{{ props.row.when }}</span>
  </q-td>
  <q-td key="mode" :props="props">{{ props.row.mode }}</q-td>
  <q-td key="profile" :props="props">{{ props.row.profile }}</q-td>
  <q-td key="capability" :props="props">
    <span v-if="props.row.capability === 'limited'" class="caplimited">limited</span>
    <span v-else class="capfull">complete</span>
  </q-td>
  <q-td key="decisions" :props="props">
    <span v-for="d in props.row.decisions" :key="d.k" class="decisioncount"
      :title="d.k + ' ' + props.row.metric">
      <span :class="'sevdot sev-' + d.k"></span>{{ d.n }}
    </span>
    <span class="atomicnote">{{ props.row.atomic }} atomic</span>
  </q-td>
  <q-td key="files" :props="props" class="c-files">{{ props.row.files }}</q-td>
  <q-td auto-width>
    <q-btn dense flat no-caps class="linkbtn" label="Open"
      @click="$parent.$emit('open', {id: props.row.id})" />
    <q-btn dense flat round size="sm" icon="more_horiz" aria-label="More run actions">
      <q-menu>
        <q-list dense style="min-width: 10rem">
          <q-item clickable v-close-popup
            @click="$parent.$emit('rerun', {id: props.row.id})">
            <q-item-section>Re-QC</q-item-section>
          </q-item>
          <q-item v-for="k in props.row.exports" :key="k" clickable v-close-popup
            @click="$parent.$emit('export', {id: props.row.id, kind: k})">
            <q-item-section>Export {{ k }}</q-item-section>
          </q-item>
        </q-list>
      </q-menu>
    </q-btn>
  </q-td>
</q-tr>
"""

REVIEW_GROUPS_BODY_SLOT = """
<q-tr :props="props" :class="{selrow: props.row.sel}"
  role="button" tabindex="0" :data-review-id="props.row.id"
  :aria-label="'Review ' + props.row.class.replace(/_/g, ' ') +
    ' at ' + props.row.where + ' ' + props.row.location"
  @click="$parent.$emit('select', {id: props.row.id})"
  @keydown.enter.prevent="$parent.$emit('select', {id: props.row.id})"
  @keydown.space.prevent="$parent.$emit('select', {id: props.row.id})">
  <q-td key="severity" :props="props">
    <span :class="'sevdot sev-' + props.row.severity"></span
    ><span class="sevtext">{{ props.row.severity }}</span>
  </q-td>
  <q-td key="class" :props="props" class="mono c-class" :title="props.row.class">{{
    props.row.class.replace(/_/g, ' ') }}</q-td>
  <q-td key="where" :props="props" class="c-loc"
    :title="props.row.where + ' ' + props.row.location">{{
    props.row.where }}<span v-if="props.row.where && props.row.location"> · </span
    ><span class="mono">{{ props.row.location }}</span></q-td>
  <q-td key="members" :props="props" class="groupcount">{{ props.row.members }}</q-td>
  <q-td key="message" :props="props"><span class="reviewclass-inline">{{
    props.row.class.replace(/_/g, ' ') }}</span>{{ props.row.message }}
    <span v-if="props.row.reviewed" class="reviewed"
      :title="props.row.reviewed + ' of ' + props.row.members + ' reviewed'"
      >reviewed {{ props.row.reviewed }}/{{ props.row.members }}</span>
    <span v-if="props.row.cap_degraded" class="capbadge">
      retained details; coverage degraded
    </span>
  </q-td>
</q-tr>
"""

REVIEW_MEMBER_ROWS_SLOT = """
<q-tr :props="props" :class="{selrow: props.row.sel}"
  role="button" tabindex="0" :data-member-id="props.row.id"
  :aria-label="'Review finding ' + props.row.id + ' at ' + props.row.location"
  @click="$parent.$emit('select', {id: props.row.id})"
  @keydown.enter.prevent="$parent.$emit('select', {id: props.row.id})"
  @keydown.space.prevent="$parent.$emit('select', {id: props.row.id})">
  <q-td key="id" :props="props" class="mono">{{ props.row.id }}</q-td>
  <q-td key="severity" :props="props">
    <span :class="'sevdot sev-' + props.row.severity"></span
    ><span class="sevtext">{{ props.row.severity }}</span
    ><span v-if="props.row.overridden" class="overridden">analyst</span
    ><span v-else-if="props.row.comment" class="reviewed">note</span>
  </q-td>
  <q-td key="class" :props="props" class="mono c-class">{{
    props.row.class.replace(/_/g, ' ') }}</q-td>
  <q-td key="location" :props="props" class="mono">{{ props.row.location }}</q-td>
  <q-td key="message" :props="props">{{ props.row.message }}</q-td>
</q-tr>
"""

#: Full body slot: main row + expandable detail row (baseline/current/impacts).
FINDINGS_BODY_SLOT = """
<q-tr :props="props">
  <q-td auto-width>
    <q-btn size="sm" color="grey-7" flat dense round
      :icon="props.expand ? 'expand_less' : 'expand_more'"
      @click="props.expand = !props.expand" />
  </q-td>
  <q-td key="id" :props="props" class="mono">{{ props.row.id }}</q-td>
  <q-td key="severity" :props="props">
    <span :class="'sevdot sev-' + props.row.severity"></span
    ><span class="sevtext">{{ props.row.severity }}</span
    ><span v-if="props.row.overridden" class="overridden">analyst</span
    ><span v-else-if="props.row.comment" class="reviewed">note</span>
  </q-td>
  <q-td key="class" :props="props" class="mono">{{ props.row.class }}</q-td>
  <q-td key="where" :props="props">{{ props.row.where }}</q-td>
  <q-td key="location" :props="props" class="mono">{{ props.row.location }}</q-td>
  <q-td key="message" :props="props">{{ props.row.message }}</q-td>
</q-tr>
<q-tr v-show="props.expand" :props="props" class="detailrow">
  <q-td colspan="100%">
    <div class="detailgrid">
      <template v-if="props.row.baseline">
        <div class="dk">baseline</div><div class="dv mono">{{ props.row.baseline }}</div>
      </template>
      <template v-if="props.row.current">
        <div class="dk">current</div><div class="dv mono">{{ props.row.current }}</div>
      </template>
      <template v-if="props.row.element">
        <div class="dk">element</div><div class="dv">{{ props.row.element }}</div>
      </template>
      <template v-if="props.row.impacts">
        <div class="dk">impacts</div><div class="dv mono">{{ props.row.impacts }}</div>
      </template>
      <template v-if="props.row.root">
        <div class="dk">root cause</div><div class="dv mono">{{ props.row.root }}</div>
      </template>
      <template v-if="props.row.waiver">
        <div class="dk">waiver</div><div class="dv">{{ props.row.waiver }}</div>
      </template>
      <template v-if="props.row.provenance">
        <div class="dk">provenance</div><div class="dv mono">{{ props.row.provenance }}</div>
      </template>
      <template v-if="props.row.subtype">
        <div class="dk">subtype</div><div class="dv mono">{{ props.row.subtype }}</div>
      </template>
      <template v-if="props.row.materiality">
        <div class="dk">materiality</div><div class="dv mono">{{ props.row.materiality }}</div>
      </template>
      <template v-if="props.row.temporal_context">
        <div class="dk">temporal context</div>
        <div class="dv mono">{{ props.row.temporal_context }}</div>
      </template>
      <template v-if="props.row.expected_reason">
        <div class="dk">expected reason</div>
        <div class="dv mono">{{ props.row.expected_reason }}</div>
      </template>
      <template v-if="props.row.evidence_tags">
        <div class="dk">evidence</div><div class="dv mono">{{ props.row.evidence_tags }}</div>
      </template>
      <div class="dk">artifact</div><div class="dv">{{ props.row.artifact }}</div>
      <div class="dk">review</div>
      <div class="dv">
        <div v-if="props.row.mutable" class="annotrow">
          <q-select dense outlined options-dense class="annotsev" label="severity"
            :model-value="props.row.severity"
            :options="['critical','warning','info','expected']"
            @update:model-value="v => { props.row.severity = v; props.row.overridden = true;
              $parent.$emit('sev', {id: props.row.id, value: v}) }" />
          <q-btn dense flat no-caps class="confirmsev" label="Confirm severity"
            :aria-label="'Confirm current severity for ' + props.row.id"
            @click="props.row.overridden = true;
              $parent.$emit('sev', {id: props.row.id, value: props.row.severity})" />
          <q-input dense outlined class="annotcomment" label="analyst comment"
            :model-value="props.row.comment"
            @update:model-value="v => props.row.comment = v"
            @blur="() => $parent.$emit('note', {id: props.row.id, value: props.row.comment})" />
        </div>
        <span v-else class="reviewed">finalized — review state locked</span>
      </div>
      <template v-if="props.row.bx || props.row.cx">
        <div class="dk">context</div>
        <div class="dv">
          <div class="ctxpair">
            <div v-if="props.row.bx" class="ctxblock">
              <div class="ctxlabel">baseline</div>
              <table class="ctxgrid"><tbody>
                <tr><th></th><th v-for="c in props.row.bx.cols" :key="c">{{ c }}</th></tr>
                <tr v-for="(r, ri) in props.row.bx.cells" :key="ri">
                  <th>{{ props.row.bx.rows[ri] }}</th>
                  <td v-for="(v, ci) in r" :key="ci"
                    :class="{hit: ri===props.row.bx.hit_row && ci===props.row.bx.hit_col}"
                  >{{ v }}</td>
                </tr>
              </tbody></table>
            </div>
            <div v-if="props.row.cx" class="ctxblock">
              <div class="ctxlabel">current</div>
              <table class="ctxgrid"><tbody>
                <tr><th></th><th v-for="c in props.row.cx.cols" :key="c">{{ c }}</th></tr>
                <tr v-for="(r, ri) in props.row.cx.cells" :key="ri">
                  <th>{{ props.row.cx.rows[ri] }}</th>
                  <td v-for="(v, ci) in r" :key="ci"
                    :class="{hit: ri===props.row.cx.hit_row && ci===props.row.cx.hit_col}"
                  >{{ v }}</td>
                </tr>
              </tbody></table>
            </div>
          </div>
        </div>
      </template>
    </div>
  </q-td>
</q-tr>
"""


def section(label: str) -> None:
    """A numbered micro-heading with a hairline rule."""
    with ui.element("div").classes("section"):
        ui.label(label).classes("kicker")
        ui.element("div").classes("rule")


def status_chip(tone: str, title: str, detail: str = "") -> None:
    """One compact semantic status: `attention`, `limited`, `ok`, or `neutral`."""
    with ui.element("div").classes(f"statuschip status-{tone}"):
        ui.label(title).classes("t")
        if detail:
            ui.label(detail).classes("d")


@contextmanager
def page_frame(
    active: str,
    *,
    network_mode: str = "local",
    expires_at: str | None = None,
    on_shutdown: Callable[[], None] | None = None,
    colophon: str | None = None,
) -> Iterator[None]:
    """Shared chrome: ink header, paper content column, mono footer."""
    ui.colors(
        primary="#1d2025",
        secondary="#4d5560",
        accent="#175cd3",
        positive="#067647",
        negative="#b42318",
        warning="#b54708",
        info="#175cd3",
    )
    ui.add_head_html(f"<style>{CSS}</style>")
    dark = ui.dark_mode(value=bool(app.storage.general.get("dark_mode", False)))
    with ui.header(elevated=False).classes("appheader"), ui.element("div").classes("inner"):
        ui.label("QC Tool").classes("wordmark")
        ui.label(f"cadence-diff v{__version__}").classes("wordtag")
        with ui.element("nav").classes("headnav"):
            ui.html(
                f'<a href="/" class="{"active" if active == "run" else ""}">Compare</a>'
                f'<a href="/history" class="{"active" if active == "history" else ""}">'
              "Run history</a>"
              f'<a href="/guide" class="{"active" if active == "guide" else ""}">'
              "Guide</a>"
            )
            ui.label(
                "NETWORK EXPOSED · unauthenticated"
                if network_mode == "lan"
                else "local · read-only"
            ).classes("headnote")

            def toggle_dark() -> None:
                dark.value = not dark.value
                app.storage.general["dark_mode"] = dark.value
                theme_button.props(
                    f"icon={'light_mode' if dark.value else 'dark_mode'}"
                )

            theme_button = (
                ui.button(
                    icon="light_mode" if dark.value else "dark_mode",
                    on_click=toggle_dark,
                )
                .classes("themebtn")
                .props('flat round dense aria-label="Toggle dark mode"')
                .mark("dark-toggle")
            )
            if on_shutdown is not None:
                ui.button(
                    icon="power_settings_new", on_click=on_shutdown
                ).classes("themebtn quitbtn").props(
                    'flat round dense aria-label="Stop the QC Tool server"'
                ).mark("quit").tooltip("Stop the QC Tool server")
    with ui.column().classes("page-wrap"):
        if network_mode == "lan":
            expiry = f" until {expires_at}" if expires_at else ""
            ui.label(
                "NETWORK ACCESS ENABLED — no built-in authentication or TLS"
                f"{expiry}. Router forwarding is separate."
            ).classes("exposurebanner")
        yield
    with ui.element("div").classes("appfooter"):
        ui.label("all processing happens on this machine — source files are never modified")
        if colophon:
            ui.label(colophon).classes("colophon")
