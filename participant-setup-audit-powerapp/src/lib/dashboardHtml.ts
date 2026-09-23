import type { DashboardModel } from "./types";

type DashboardModels = Record<string, DashboardModel>;

interface DashboardHtmlView {
  snapshot: string;
  kpis: string;
  analystTable: string;
  regionBars: string;
  lobBars: string;
  lobTable: string;
}

export function buildDashboardHtml(
  models: DashboardModels,
  selectedRegion: string,
  processingMonth: string,
): string {
  const initialRegion = models[selectedRegion] ? selectedRegion : "All Regions";
  const views = Object.fromEntries(
    Object.entries(models).map(([region, model]) => [region, buildDashboardView(model)]),
  );
  const initialView = views[initialRegion];
  const serializedViews = JSON.stringify(views)
    .replace(/</g, "\\u003c")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
  const serializedRegion = JSON.stringify(initialRegion).replace(/</g, "\\u003c");
  const regionOptions = Object.keys(models)
    .map((region) => `<option value="${escapeHtml(region)}"${region === initialRegion ? " selected" : ""}>${escapeHtml(region)}</option>`)
    .join("");
  const regionEntries = Object.keys(models).map((region, index) => ({ region, index, view: views[region] }));
  const fallbackInputs = regionEntries
    .map(({ region, index }) => `<input class="dashboard-region-toggle" type="radio" name="dashboard-region" id="dashboard-region-${index}"${region === initialRegion ? " checked" : ""}>`)
    .join("");
  const fallbackOptions = regionEntries
    .map(({ region, index }) => `<label class="dashboard-region-option dashboard-region-option-${index}" for="dashboard-region-${index}">${escapeHtml(region)}</label>`)
    .join("");
  const fallbackViews = regionEntries
    .map(({ region, index, view }) => `<div class="dashboard-fallback-view dashboard-fallback-view-${index}"><p class="dashboard-fallback-context">Region: ${escapeHtml(region)}</p>${buildDashboardSections(view)}</div>`)
    .join("");
  const fallbackRules = regionEntries
    .map(({ index }) => `#dashboard-region-${index}:checked ~ .shell .dashboard-fallback-view-${index} { display: grid; } #dashboard-region-${index}:checked ~ .shell .dashboard-region-option-${index} { border-color: #fff; background: #fff; color: #16324f; }`)
    .join("\n    ");

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Participant Setup Dashboard</title>
  <style>
    :root { color-scheme: light; font-family: Aptos, "Segoe UI", sans-serif; color: #172033; background: #eef2f6; }
    * { box-sizing: border-box; }
    body { margin: 0; background: #eef2f6; }
    .shell { display: grid; gap: 18px; width: min(100% - 24px, 1400px); margin: 0 auto; padding: 18px 0 32px; }
    .hero { display: flex; justify-content: space-between; gap: 24px; align-items: start; padding: 28px; border-radius: 14px; color: #f8fafc; background: linear-gradient(120deg, rgba(255,255,255,.06), transparent 46%), #16324f; }
    .eyebrow { margin: 0 0 6px; color: #e86f14; font-size: 11px; font-weight: 800; letter-spacing: .08em; text-transform: uppercase; }
    h1, h2 { margin: 0; }
    h1 { color: #fff; font-size: clamp(24px, 3vw, 34px); }
    .hero-copy { max-width: 760px; margin: 10px 0 0; color: #cbd5e1; }
    .meta { margin: 6px 0 0; color: #cbd5e1; font-size: 12px; }
    .filter { display: grid; gap: 5px; min-width: 210px; }
    .filter label { color: #cbd5e1; font-size: 11px; text-transform: uppercase; }
    select { min-height: 38px; padding: 6px 12px; border: 1px solid #d1d5db; border-radius: 8px; background: #fff; color: #172033; font: inherit; }
    .dashboard-region-toggle { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); clip-path: inset(50%); white-space: nowrap; }
    .dashboard-script-filter, .dashboard-script-view { display: none; }
    .dashboard-fallback-filter { min-width: 0; margin: 0; padding: 0; border: 0; }
    .dashboard-fallback-filter legend { margin-bottom: 6px; padding: 0; color: #cbd5e1; font-size: 11px; text-transform: uppercase; }
    .dashboard-region-options { display: flex; flex-wrap: wrap; gap: 8px; }
    .dashboard-region-option { display: inline-flex; min-height: 40px; align-items: center; padding: 8px 12px; border: 1px solid #94a3b8; border-radius: 999px; color: #f8fafc; font-size: 13px; font-weight: 700; line-height: 1; text-transform: none; cursor: pointer; }
    .dashboard-fallback-view { display: none; gap: 18px; }
    .dashboard-fallback-context { margin: 0; color: #475569; font-size: 12px; font-weight: 700; text-transform: uppercase; }
    .dashboard-js .dashboard-script-filter, .dashboard-js .dashboard-script-view { display: grid; }
    .dashboard-js .dashboard-script-view { gap: 18px; }
    .dashboard-js .dashboard-fallback-filter, .dashboard-js .dashboard-fallback-views { display: none; }
    ${fallbackRules}
    .print-region { display: none; }
    .kpis { display: grid; gap: 10px; }
    .kpi-row { display: grid; gap: 10px; }
    .kpi-row.population { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    .kpi-row.execution { grid-template-columns: repeat(6, minmax(0, 1fr)); }
    .kpi { display: grid; gap: 12px; min-height: 112px; padding: 16px; border: 1px solid #dbe3ec; border-top: 4px solid #64748b; border-radius: 10px; background: #fff; box-shadow: 0 4px 16px rgba(15,23,42,.05); }
    .kpi span { color: #64748b; font-size: 12px; font-weight: 700; text-transform: uppercase; }
    .kpi strong { align-self: end; font-size: clamp(27px, 3vw, 38px); letter-spacing: -.04em; }
    .tone-navy { border-top-color: #16324f; } .tone-ink { border-top-color: #475569; } .tone-green { border-top-color: #15803d; } .tone-amber { border-top-color: #d97706; } .tone-coral { border-top-color: #ea580c; } .tone-blue { border-top-color: #2563eb; }
    .panel { padding: 20px; border: 1px solid #e5e7eb; border-radius: 12px; background: #fff; }
    .panel-head { display: flex; justify-content: space-between; gap: 20px; align-items: start; margin-bottom: 14px; }
    .panel-head h2, .chart h2 { font-size: 17px; }
    .panel-copy { margin: 6px 0 0; color: #64748b; font-size: 13px; }
    .caption { color: #64748b; font-size: 12px; }
    .grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px; }
    .bar-list { display: grid; gap: 12px; margin-top: 18px; }
    .bar-row { display: grid; gap: 5px; }
    .bar-label { display: flex; justify-content: space-between; gap: 12px; color: #475569; font-size: 13px; }
    .bar-track { height: 9px; overflow: hidden; border-radius: 999px; background: #e2e8f0; }
    .bar-track span { display: block; height: 100%; border-radius: inherit; background: linear-gradient(90deg, #16324f, #3b82f6); }
    .table-shell { overflow: auto; border: 1px solid #e2e8f0; border-radius: 8px; }
    table { width: 100%; min-width: 760px; border-collapse: collapse; }
    th, td { padding: 10px 12px; border-bottom: 1px solid #f3f4f6; text-align: left; font-size: 13px; }
    th { color: #64748b; font-size: 11px; text-transform: uppercase; }
    th:not(:first-child), td:not(:first-child) { text-align: right; }
    .empty { padding: 20px; color: #64748b; text-align: center; }
    @media (max-width: 820px) { .hero, .panel-head { flex-direction: column; } .filter { width: 100%; } select { width: 100%; min-height: 44px; } .dashboard-region-option { min-height: 44px; } .kpi-row.population, .kpi-row.execution { grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); } .grid { grid-template-columns: 1fr; } .table-shell { -webkit-overflow-scrolling: touch; } }
    @media (max-width: 520px) { .shell { width: calc(100% - 16px); gap: 12px; padding: 8px 0 20px; } .hero { gap: 18px; padding: 20px; border-radius: 12px; } .hero-copy { font-size: 14px; line-height: 1.45; } .kpi-row.population, .kpi-row.execution { grid-template-columns: repeat(2, minmax(0, 1fr)); } .kpi { min-height: 96px; padding: 12px; } .kpi span { font-size: 10px; } .kpi strong { font-size: 28px; } .panel { padding: 16px; } .panel-head { gap: 8px; } table { min-width: 680px; } th, td { padding: 9px 10px; } }
    @media print {
      @page { size: A4 portrait; margin: 6mm; }
      body { background: #fff; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
      .shell { width: 100%; padding: 0; gap: 6px; }
      .dashboard-script-filter, .dashboard-fallback-filter { display: none !important; } .dashboard-js .print-region { display: block; } .print-region, .dashboard-fallback-context { color: #cbd5e1; font-size: 9px; font-weight: 700; text-transform: uppercase; }
      .dashboard-script-view, .dashboard-fallback-view { gap: 6px; }
      .hero { padding: 10px 12px; } h1 { font-size: 20px; } .hero-copy, .meta, .eyebrow { font-size: 9px; }
      .kpis, .kpi-row { gap: 5px; } .kpi-row.population { grid-template-columns: repeat(3, minmax(0, 1fr)); } .kpi-row.execution { grid-template-columns: repeat(6, minmax(0, 1fr)); }
      .kpi { gap: 4px; min-height: 58px; padding: 7px 8px; box-shadow: none; } .kpi span { font-size: 8px; } .kpi strong { font-size: 21px; }
      .panel { padding: 9px; border-radius: 7px; } .panel-head { flex-direction: row; gap: 6px; margin-bottom: 6px; } .panel-head h2, .chart h2 { font-size: 12px; }
      .panel-copy, .caption { font-size: 8px; } .grid { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 6px; }
      .bar-list, .bar-row { gap: 3px; margin-top: 5px; } .bar-label { font-size: 8px; } .bar-track { height: 5px; }
      .table-shell { overflow: visible; } table { min-width: 0; table-layout: fixed; } th, td { padding: 3px 4px; font-size: 8px; line-height: 1.15; }
      thead { display: table-header-group; } .hero, .kpis, .kpi, .grid, tr { break-inside: avoid; }
    }
  </style>
</head>
<body>
  ${fallbackInputs}
  <main class="shell">
    <section class="hero">
      <div>
        <p class="eyebrow">Manager control tower</p>
        <h1>Commissioned population and setup execution</h1>
        <p class="hero-copy">Employees with Active Status = Yes in the current SCR, paired with the latest verification baseline.</p>
        <p class="meta">Processing month: ${escapeHtml(processingMonth)} <span id="snapshot">${escapeHtml(initialView.snapshot)}</span></p>
      </div>
      <div class="filter dashboard-script-filter"><label for="region-filter">Region</label><select id="region-filter">${regionOptions}</select></div>
      <fieldset class="filter dashboard-fallback-filter"><legend>Region</legend><div class="dashboard-region-options">${fallbackOptions}</div></fieldset>
      <span class="print-region" id="print-region">Region: ${escapeHtml(initialRegion)}</span>
    </section>
    <div class="dashboard-script-view">${buildDashboardSections(initialView, true)}</div>
    <div class="dashboard-fallback-views">${fallbackViews}</div>
  </main>
  <script>
    var views = ${serializedViews};
    var defaultRegion = ${serializedRegion};
    var regionFilter = document.getElementById("region-filter");
    regionFilter.value = defaultRegion;

    function renderDashboard() {
      var region = regionFilter.value;
      var view = views[region] || views[defaultRegion];
      document.title = "Participant Setup Dashboard - " + region;
      document.getElementById("print-region").textContent = "Region: " + region;
      document.getElementById("snapshot").textContent = view.snapshot;
      document.getElementById("kpis").innerHTML = view.kpis;
      document.getElementById("analyst-table").innerHTML = view.analystTable;
      document.getElementById("region-bars").innerHTML = view.regionBars;
      document.getElementById("lob-bars").innerHTML = view.lobBars;
      document.getElementById("lob-table").innerHTML = view.lobTable;
    }

    regionFilter.addEventListener("change", renderDashboard);
    renderDashboard();
    document.documentElement.className += " dashboard-js";
  </script>
</body>
</html>`;
}

export function buildDashboardHtmlFileName(processingMonth: string, now = new Date()): string {
  const stamp = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
    "_",
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
    String(now.getSeconds()).padStart(2, "0"),
  ].join("");
  return `Participant_Setup_Dashboard_${processingMonth}_${stamp}.html`;
}

export const METRIC_DESCRIPTIONS: Record<string, string> = {
  "Commissioned Employees": "Distinct employees with Active Status = Yes in the current SCR for the selected region.",
  "Setup Required": "Completed, Partially Completed, and Pending setup actions; excludes Manager Mismatch Only, Deferred, and Not Verifiable.",
  "Setup Required (%)": "Setup Required divided by Commissioned Employees.",
  Completed: "All verifiable setup fields match the expected values in the follow-up People file.",
  "Partially Completed": "Some verifiable setup fields are complete, while others are still pending.",
  Pending: "None of the verifiable setup fields match the expected values yet.",
  "Term Update Pending": "Termination items awaiting a later mass update; already included in Partially Completed or Pending. Setup Required and Completion Rate are unchanged.",
  "Manager Mismatch Only": "The only outstanding change is the Level 1 Manager; excluded from Setup Required and Completion Rate.",
  "Completion Rate": "Completed divided by Completed, Partially Completed, and Pending setup actions.",
};

function buildDashboardSections(view: DashboardHtmlView, includeIds = false): string {
  const id = (value: string) => includeIds ? ` id="${value}"` : "";
  return `<section class="kpis"${id("kpis")}>${view.kpis}</section>
    <section class="panel">
      <div class="panel-head"><div><p class="eyebrow">Analyst view</p><h2>Analyst setup ownership</h2><p class="panel-copy">Missing assignments use the unique top Analyst for Country + SCR LOB, then Region + SCR LOB. Ties remain Unassigned.</p></div><span class="caption">Inferred assignments are estimates, not Xactly master data.</span></div>
      <div${id("analyst-table")}>${view.analystTable}</div>
    </section>
    <section class="grid"><section class="panel chart"><h2>Commissioned employees by Region</h2><div class="bar-list"${id("region-bars")}>${view.regionBars}</div></section><section class="panel chart"><h2>Commissioned employees by LOB</h2><div class="bar-list"${id("lob-bars")}>${view.lobBars}</div></section></section>
    <section class="panel"><div class="panel-head"><div><p class="eyebrow">Execution</p><h2>Setup progress by LOB</h2></div></div><div${id("lob-table")}>${view.lobTable}</div></section>`;
}

function buildDashboardView(model: DashboardModel): DashboardHtmlView {
  return {
    snapshot: model.latestPeopleDate ? ` · People snapshot: ${model.latestPeopleDate}` : "",
    kpis: buildKpis(model),
    analystTable: buildTable(model.byAnalyst, "Analyst", true),
    regionBars: buildBars(model.byRegion),
    lobBars: buildBars(model.byLob),
    lobTable: buildTable(model.byLob, "LOB", false),
  };
}

function buildKpis(model: DashboardModel): string {
  const metrics = [
    { label: "Commissioned Employees", value: model.commissionedEmployees, tone: "navy" },
    { label: "Setup Required", value: model.setupRequired, tone: "ink" },
    { label: "Setup Required (%)", value: model.setupRequiredRate, tone: "blue", percent: true },
    { label: "Completed", value: model.completed, tone: "green" },
    { label: "Partially Completed", value: model.partiallyCompleted, tone: "amber" },
    { label: "Pending", value: model.pending, tone: "coral" },
    { label: "Term Update Pending", value: model.termUpdatePending, tone: "blue" },
    { label: "Manager Mismatch Only", value: model.managerMismatchOnly, tone: "amber" },
    { label: "Completion Rate", value: model.completionRate, tone: "blue", percent: true },
  ];
  const cards = metrics.map(
    (metric) => `<article class="kpi tone-${metric.tone}" title="${escapeHtml(METRIC_DESCRIPTIONS[metric.label])}"><span>${escapeHtml(metric.label)}</span><strong>${metric.percent ? formatPercent(metric.value) : formatCount(metric.value)}</strong></article>`,
  );
  return `<div class="kpi-row population">${cards.slice(0, 3).join("")}</div><div class="kpi-row execution">${cards.slice(3).join("")}</div><p class="panel-copy">Term Update Pending: ${escapeHtml(METRIC_DESCRIPTIONS["Term Update Pending"])}</p>`;
}

function buildTable(
  rows: DashboardModel["byLob"],
  firstColumn: string,
  showInferred: boolean,
): string {
  if (!rows.length) return '<div class="empty">No data in this selection.</div>';
  const inferredHeader = showInferred ? "<th>Inferred</th>" : "";
  const body = rows
    .map(
      (row) => `<tr><td>${escapeHtml(row.label)}</td><td>${formatCount(row.employees)}</td><td>${formatCount(row.required)}</td><td>${formatCount(row.completed)}</td><td>${formatCount(row.pending)}</td><td>${formatCount(row.overdue)}</td>${showInferred ? `<td>${formatCount(row.inferred)}</td>` : ""}<td>${row.required > 0 ? `${Math.round((row.completed / row.required) * 100)}%` : "—"}</td></tr>`,
    )
    .join("");
  return `<div class="table-shell"><table><thead><tr><th>${escapeHtml(firstColumn)}</th><th>Employees</th><th>Required</th><th>Completed</th><th>Pending</th><th>Overdue</th>${inferredHeader}<th>Completion</th></tr></thead><tbody>${body}</tbody></table></div>`;
}

function buildBars(rows: DashboardModel["byLob"]): string {
  if (!rows.length) return '<div class="empty">No commissioned employees in this selection.</div>';
  const maximum = Math.max(1, ...rows.map((row) => row.employees));
  return rows
    .map(
      (row) => `<div class="bar-row"><div class="bar-label"><span>${escapeHtml(row.label)}</span><strong>${formatCount(row.employees)}</strong></div><div class="bar-track"><span style="width:${Math.max(2, (row.employees / maximum) * 100)}%"></span></div></div>`,
    )
    .join("");
}

function formatCount(value: number): string {
  return Number(value || 0).toLocaleString();
}

function formatPercent(value: number): string {
  return `${Math.round(Number(value || 0) * 100)}%`;
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  })[character] ?? character);
}

export type { DashboardModels };
