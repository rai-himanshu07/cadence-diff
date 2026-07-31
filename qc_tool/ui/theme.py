"""Visual theme for the QC tool: quiet, print-inspired, instrument-panel.

Design rules: warm paper background, near-black ink, hairline borders
instead of shadows, system fonts only (the tool runs offline), monospace
for cell references and file names, and color reserved for severity —
the one thing that must pop. No gradients, no decoration.
"""

from collections.abc import Iterator
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
.appheader .inner { max-width: 1180px; margin: 0 auto; width: 100%;
  display: flex; align-items: baseline; gap: 1rem; padding: 0.65rem 1.25rem; }
.wordmark { color: #f4f2ed; font-weight: 650; font-size: 1.02rem;
  letter-spacing: 0.01em; }
.wordtag { font-family: var(--font-mono); font-size: 0.7rem; color: #9aa3ad;
  border: 1px solid #3a4048; border-radius: 3px; padding: 0.1rem 0.4rem; }
.headnav { margin-left: auto; display: flex; gap: 1.25rem; align-items: baseline; }
.headnav a { color: #c8cdd3; font-size: 0.82rem; text-decoration: none;
  padding-bottom: 2px; border-bottom: 1px solid transparent; }
.headnav a + a { margin-left: 1.25rem; }
.headnav a:hover { color: #fff; }
.headnav a.active { color: #fff; border-bottom-color: #8d959e; }
.headnote { font-family: var(--font-mono); font-size: 0.68rem; color: #6d757e; }
/* the toggle sits on the always-dark band: force light icon over Quasar's
   text-primary (which is near-black and vanishes against the header) */
.themebtn { border: 1px solid #3a4048; }
.themebtn, .themebtn .q-icon { color: #cfd4da !important; }
.themebtn:hover { border-color: #8d959e; }
.themebtn:hover, .themebtn:hover .q-icon { color: #ffffff !important; }

/* layout */
.page-wrap { max-width: 1180px; margin: 0 auto; width: 100%;
  padding: 1.6rem 1.25rem 3rem; gap: 0; }
.appfooter { max-width: 1180px; margin: 0 auto; width: 100%;
  border-top: 1px solid var(--line); padding: 0.7rem 1.25rem 2rem;
  font-family: var(--font-mono); font-size: 0.68rem; color: var(--ink-soft); }

/* section headings */
.section { display: flex; align-items: center; gap: 0.75rem;
  margin: 1.9rem 0 0.9rem; width: 100%; }
.section:first-child { margin-top: 0.4rem; }
.section .kicker { font-size: 0.68rem; letter-spacing: 0.14em; font-weight: 650;
  color: var(--ink-soft); text-transform: uppercase; white-space: nowrap; }
.section .rule { flex: 1; border-top: 1px solid var(--line); }

.lede { color: var(--ink-soft); font-size: 0.86rem; max-width: 46rem; }
.mode-select { width: 100%; max-width: 52rem; border: 1px solid var(--line);
  border-radius: 5px; overflow: hidden; }
.mode-select .q-btn { min-height: 2.5rem; background: var(--panel); color: var(--ink); }
.mode-select .q-btn .q-btn__content { color: var(--ink) !important; }
.mode-select .q-btn[aria-pressed="true"] { background: var(--btn-bg) !important;
  color: var(--btn-fg) !important; }
.mode-select .q-btn[aria-pressed="true"] .q-btn__content {
  color: var(--btn-fg) !important; }
.modecopy { color: var(--ink-soft); font-size: 0.78rem; margin-bottom: 0.55rem; }
.guide-jump { color: var(--info) !important; font-size: 0.72rem;
  text-decoration: none; margin: -0.35rem 0 0.55rem; }
.guide-jump:hover { text-decoration: underline; }

/* guide */
.guide-page-title { font-size: 1.55rem; font-weight: 680; line-height: 1.2; }
.guide-page-lede { color: var(--ink-soft); font-size: 0.88rem; max-width: 48rem;
  margin-top: 0.35rem; }
.guide-layout { display: grid; grid-template-columns: 13rem minmax(0, 1fr);
  gap: 2.5rem; width: 100%; align-items: start; margin-top: 1.5rem; }
.guide-toc { position: sticky; top: 1rem; border-left: 2px solid var(--line);
  padding-left: 0.85rem; gap: 0.25rem; }
.guide-toc-title { color: var(--ink-soft); font-size: 0.64rem;
  letter-spacing: 0.12em; text-transform: uppercase; font-weight: 700;
  margin-bottom: 0.35rem; }
.guide-toc-link { display: block; color: var(--ink-soft) !important;
  font-size: 0.76rem; line-height: 1.35; text-decoration: none; padding: 0.18rem 0; }
.guide-toc-link:hover { color: var(--ink) !important; }
.guide-content { min-width: 0; }
.guide-section { scroll-margin-top: 1rem; padding: 0 0 1.8rem;
  border-bottom: 1px solid var(--line); margin-bottom: 1.8rem; }
.guide-section:last-child { border-bottom: 0; }
.guide-section-title { font-size: 1.12rem; font-weight: 680; margin-bottom: 0.65rem; }
.guide-copy { color: var(--ink); font-size: 0.84rem; line-height: 1.62;
  max-width: 52rem; white-space: normal; }
.guide-list { color: var(--ink); font-size: 0.82rem; line-height: 1.55;
  padding-left: 1.2rem; margin: 0.65rem 0; }
.guide-list li { margin: 0.3rem 0; }
.guide-list code, .guide-copy code { font-family: var(--font-mono);
  background: var(--surface2); border: 1px solid var(--line-soft);
  border-radius: 3px; padding: 0.05rem 0.25rem; font-size: 0.76rem; }
.guide-code { background: var(--surface1); border: 1px solid var(--line);
  border-radius: 5px; padding: 0.85rem 1rem; overflow-x: auto; color: var(--ink);
  font-family: var(--font-mono); font-size: 0.73rem; line-height: 1.5;
  margin: 0.8rem 0; white-space: pre; }
.guide-table-wrap { width: 100%; overflow-x: auto; margin: 0.75rem 0; }
.guide-table { width: 100%; min-width: 38rem; border-collapse: collapse;
  font-size: 0.77rem; }
.guide-table th { background: var(--surface2); color: var(--ink); text-align: left;
  font-size: 0.65rem; letter-spacing: 0.08em; text-transform: uppercase;
  padding: 0.45rem 0.55rem; border: 1px solid var(--line); }
.guide-table td { color: var(--ink); padding: 0.45rem 0.55rem;
  border: 1px solid var(--line); vertical-align: top; line-height: 1.4; }
.guide-callout { border-left: 3px solid var(--info); background: var(--surface1);
  padding: 0.65rem 0.8rem; margin: 0.8rem 0; max-width: 52rem; }
.guide-callout-warning { border-left-color: var(--warning); }
.guide-callout-title { font-size: 0.78rem; font-weight: 700; }
.guide-callout-copy { color: var(--ink-soft); font-size: 0.77rem;
  line-height: 1.5; margin-top: 0.15rem; white-space: normal; }

/* panels + quasar surfaces */
.q-card { background: var(--panel); border: 1px solid var(--line);
  border-radius: 6px; box-shadow: none; }
.q-uploader { background: var(--panel); border: 1px solid var(--line);
  border-radius: 6px; box-shadow: none; width: 100%; }
.q-uploader__header { background: var(--surface2); color: var(--ink); }
.q-uploader__title { font-size: 0.78rem; font-weight: 600; }
.q-uploader__subtitle { font-family: var(--font-mono); font-size: 0.66rem; }
.q-uploader__list { min-height: 1.4rem; padding: 0.35rem 0.6rem; }
.q-expansion-item { background: var(--panel); border: 1px solid var(--line);
  border-radius: 6px; }

.upgrid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 1rem 1.5rem; width: 100%; }
.rolelabel { font-size: 0.68rem; letter-spacing: 0.12em; font-weight: 650;
  text-transform: uppercase; color: var(--ink-soft); margin-bottom: 0.25rem; }
.filestate { font-family: var(--font-mono); font-size: 0.72rem;
  color: var(--ink-soft); margin-top: 0.3rem; }
.filestate.ok { color: var(--expected); }
.filestate.err { color: var(--critical); font-weight: 650; }

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

/* quasar form surfaces in dark mode */
body.body--dark .q-field--outlined .q-field__control { background: var(--panel); }
body.body--dark .q-field__native, body.body--dark .q-field__input,
body.body--dark .q-field__label { color: var(--ink); }
body.body--dark .q-chip { background: var(--surface2); color: var(--ink); }
body.body--dark .q-select__dropdown-icon { color: var(--ink-soft); }

/* severity stat strip */
.statstrip { display: flex; gap: 1rem; width: 100%; flex-wrap: wrap; }
.stat { background: var(--panel); border: 1px solid var(--line);
  border-top-width: 3px; border-radius: 4px; padding: 0.5rem 1.1rem 0.55rem;
  min-width: 8.5rem; gap: 0; }
.stat .n { font-size: 1.7rem; font-weight: 650; line-height: 1.15;
  font-variant-numeric: tabular-nums; }
.stat .l { font-size: 0.7rem; letter-spacing: 0.1em; text-transform: uppercase;
  color: var(--ink-soft); }
.stat .a { font-family: var(--font-mono); font-size: 0.66rem;
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
  font-size: 0.7rem; letter-spacing: 0.09em; text-transform: uppercase;
  font-weight: 650; border-bottom: 1px solid var(--line); }
.findings-table td { border-color: var(--line-soft) !important;
  font-size: 0.8rem; vertical-align: top; }
.findings-table .mono { font-family: var(--font-mono); font-size: 0.75rem; }
.review-toggle { border: 1px solid var(--line); border-radius: 4px; }
.review-groups-table .groupcount { font-family: var(--font-mono);
  font-variant-numeric: tabular-nums; text-align: right; }
.capbadge { color: var(--warning); font-size: 0.65rem; display: block;
  margin-top: 0.15rem; }
.coverage-table { width: 100%; border: 1px solid var(--line); border-radius: 6px;
  background: var(--panel); }
.coverage-table thead th { background: var(--surface2); color: var(--ink);
  font-size: 0.68rem; letter-spacing: 0.08em; text-transform: uppercase; }
.coverage-table td { border-color: var(--line-soft) !important; font-size: 0.76rem; }
.mappingstats { display: flex; gap: 0.7rem; flex-wrap: wrap; width: 100%;
  margin-bottom: 0.6rem; }
.mappingstat { min-width: 6.8rem; gap: 0; border-left: 2px solid var(--line);
  padding: 0.2rem 0.65rem; }
.mappingstat .n { font-size: 1.25rem; font-weight: 650; line-height: 1.1; }
.mappingstat .l { color: var(--ink-soft); font-size: 0.62rem;
  letter-spacing: 0.09em; text-transform: uppercase; }
.mappingitem { width: 100%; }
.mappingcontext { font-family: var(--font-mono); font-size: 0.74rem;
  color: var(--ink-soft); padding: 0.25rem 0.75rem 0.5rem; }
.candidate-row { width: 100%; align-items: center; gap: 0.8rem;
  border-top: 1px solid var(--line-soft); padding: 0.35rem 0.75rem; }
.candidate-ref { font-family: var(--font-mono); font-size: 0.75rem; min-width: 12rem; }
.candidate-value { font-family: var(--font-mono); font-size: 0.75rem; }
.candidate-match, .candidate-near { margin-left: auto; font-size: 0.68rem;
  text-transform: uppercase; letter-spacing: 0.07em; }
.candidate-match { color: var(--expected); }
.candidate-near { color: var(--warning); }
.detailrow td { background: var(--surface1); }
.detailgrid { display: grid; grid-template-columns: 6.5rem 1fr;
  gap: 0.2rem 1rem; padding: 0.25rem 0 0.35rem; max-width: 60rem; }
.detailgrid .dk { color: var(--ink-soft); text-transform: uppercase;
  font-size: 0.62rem; letter-spacing: 0.11em; padding-top: 3px; }
.detailgrid .dv { font-size: 0.78rem; overflow-wrap: anywhere; }
.detailgrid .dv.mono { font-family: var(--font-mono); font-size: 0.74rem; }
.ctxpair { display: flex; gap: 1.5rem; flex-wrap: wrap; }
.ctxlabel { font-size: 0.62rem; letter-spacing: 0.11em; text-transform: uppercase;
  color: var(--ink-soft); margin-bottom: 2px; }
.ctxgrid { border-collapse: collapse; font-family: var(--font-mono);
  font-size: 0.68rem; }
.ctxgrid th { background: var(--surface2); color: var(--ink-soft); font-weight: 500;
  padding: 1px 6px; border: 1px solid var(--line-soft); font-size: 0.62rem; }
.ctxgrid td { border: 1px solid var(--line-soft); padding: 1px 6px;
  background: var(--panel); max-width: 9rem; overflow: hidden; text-overflow: ellipsis;
  white-space: nowrap; }
.ctxgrid td.hit { outline: 2px solid var(--warning); outline-offset: -2px;
  background: var(--hit-bg); font-weight: 650; }
.deltaline { font-family: var(--font-mono); font-size: 0.78rem;
  color: var(--ink-soft); margin-top: 0.5rem; }
.deltaline .good { color: var(--expected); font-weight: 650; }
.deltaline .bad { color: var(--critical); font-weight: 650; }
.annotrow { display: flex; gap: 0.6rem; align-items: center; max-width: 44rem;
  margin-top: 0.15rem; }
.annotsev { min-width: 9.5rem; }
.annotcomment { flex: 1; }
.overridden { font-size: 0.62rem; letter-spacing: 0.08em; color: var(--info);
  text-transform: uppercase; margin-left: 0.4rem; }
.rerunbanner { border-left: 3px solid var(--info); background: var(--panel);
  border-top: 1px solid var(--line); border-right: 1px solid var(--line);
  border-bottom: 1px solid var(--line); border-radius: 0 4px 4px 0;
  padding: 0.5rem 0.85rem; font-size: 0.82rem; width: 100%; }
.sevdot { display: inline-block; width: 8px; height: 8px; border-radius: 50%;
  margin-right: 0.45rem; vertical-align: baseline; }
.sev-critical { background: var(--critical); }
.sev-warning { background: var(--warning); }
.sev-info { background: var(--info); }
.sev-expected { background: var(--expected); }
.sevtext { font-size: 0.75rem; letter-spacing: 0.04em; }

.notecard { border-left: 3px solid var(--critical); background: var(--panel);
  border-top: 1px solid var(--line); border-right: 1px solid var(--line);
  border-bottom: 1px solid var(--line); border-radius: 0 4px 4px 0;
  padding: 0.5rem 0.85rem; font-size: 0.8rem; color: var(--ink-soft);
  width: 100%; }
.exposurebanner { border: 2px solid var(--critical); background: var(--panel);
  color: var(--critical); font-weight: 700; padding: 0.65rem 0.85rem;
  border-radius: 4px; width: 100%; margin-bottom: 0.8rem; }

/* history */
.runcard { width: 100%; padding: 0.9rem 1.1rem; gap: 0.15rem; }
.runcard .runhead { font-weight: 650; font-size: 0.9rem; }
.runcard .runmeta { font-family: var(--font-mono); font-size: 0.72rem;
  color: var(--ink-soft); }
.runcounts { display: flex; gap: 0.9rem; font-size: 0.75rem; margin-top: 0.2rem;
  font-variant-numeric: tabular-nums; }

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
  .guide-section { scroll-margin-top: 0.5rem; }
  .upgrid { grid-template-columns: 1fr; }
  .statstrip { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .stat { min-width: 0; }
  .coverage-table, .findings-table { min-width: 0; max-width: 100%; }
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

REVIEW_GROUPS_BODY_SLOT = """
<q-tr :props="props">
  <q-td auto-width>
    <q-btn size="sm" color="grey-7" flat dense round icon="list_alt"
      aria-label="View affected findings"
      @click="$parent.$emit('members', {id: props.row.id})" />
    <q-btn size="sm" color="grey-7" flat dense round icon="fact_check"
      aria-label="Review affected findings"
      @click="$parent.$emit('groupreview', {id: props.row.id})" />
  </q-td>
  <q-td key="id" :props="props" class="mono">{{ props.row.id }}</q-td>
  <q-td key="severity" :props="props">
    <span :class="'sevdot sev-' + props.row.severity"></span
    ><span class="sevtext">{{ props.row.severity }}</span>
  </q-td>
  <q-td key="class" :props="props" class="mono">{{ props.row.class }}</q-td>
  <q-td key="where" :props="props">{{ props.row.where }}</q-td>
  <q-td key="location" :props="props" class="mono">{{ props.row.location }}</q-td>
  <q-td key="members" :props="props" class="groupcount">{{ props.row.members }}</q-td>
  <q-td key="message" :props="props">{{ props.row.message }}
    <span v-if="props.row.cap_degraded" class="capbadge">
      retained details; coverage degraded
    </span>
  </q-td>
</q-tr>
"""

REVIEW_MEMBERS_BODY_SLOT = """
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
    ><span v-if="props.row.overridden" class="overridden">analyst</span>
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
      <template v-if="props.row.comment">
        <div class="dk">analyst comment</div><div class="dv">{{ props.row.comment }}</div>
      </template>
      <template v-if="props.row.bx || props.row.cx">
        <div class="dk">context</div>
        <div class="dv">
          <div class="ctxpair">
            <div v-if="props.row.bx">
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
            <div v-if="props.row.cx">
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
    ><span v-if="props.row.overridden" class="overridden">analyst</span>
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
      <div class="dk">artifact</div><div class="dv">{{ props.row.artifact }}</div>
      <div class="dk">review</div>
      <div class="dv">
        <div class="annotrow">
          <q-select dense outlined options-dense class="annotsev" label="severity"
            :model-value="props.row.severity"
            :options="['critical','warning','info','expected']"
            @update:model-value="v => { props.row.severity = v; props.row.overridden = true;
              $parent.$emit('sev', {id: props.row.id, value: v}) }" />
          <q-input dense outlined class="annotcomment" label="analyst comment"
            :model-value="props.row.comment"
            @update:model-value="v => props.row.comment = v"
            @blur="() => $parent.$emit('note', {id: props.row.id, value: props.row.comment})" />
        </div>
      </div>
      <template v-if="props.row.bx || props.row.cx">
        <div class="dk">context</div>
        <div class="dv">
          <div class="ctxpair">
            <div v-if="props.row.bx">
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
            <div v-if="props.row.cx">
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


@contextmanager
def page_frame(
    active: str, *, network_mode: str = "local", expires_at: str | None = None
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
