#!/usr/bin/env node
/**
 * build_sprint_report_docx.js — turns a _sprint_data.json (produced by sprint_engine.js ingest)
 * into a polished Sprint Report .docx — the document-generation half of the headless pipeline that
 * replaces Sprint_Report_Generator.html's buildReport()/downloadReportPdf() for the Agent.
 *
 * Structure (Confirmation section removed per user request — no longer generated):
 *   Page 1: Title + meta, Sprint Overview table
 *   Page 2: Sprint Summary -> Executive Summary (3 stat cards), Activity by Developer table
 *   Page 3: Sprint Details -> Overdue Tasks / Active Work Items / Completed Work Items tables
 *
 * The summary-only variant (includeFull=false) now also includes Sprint Overview — it only omits
 * Sprint Details (the raw work-item tables).
 *
 * v2: adds an executive-facing layer to Sprint Summary — a Sprint Health KPI row (completion rate /
 * overdue / bug ratio, RAG-colored, each with a delta vs the previous sprint), the work-item-by-type
 * and open/closed donut charts, a workload-by-developer bar chart, and a "Trend Across Sprints"
 * section plotting velocity/completion/bug-ratio/overdue over every sprint folder found alongside
 * this one. Chart PNGs are pre-rendered by generate_charts.py (matplotlib) and embedded via
 * ImageRun, since docx-js has no live charting of its own.
 *
 * Usage:
 *   node build_sprint_report_docx.js <sprint_data.json> <output.docx> [--status Completed|Delayed|Partial]
 *        [--prepared-by "Name"] [--submission-date YYYY-MM-DD] [--sprint-name "Sprint 24"]
 *        [--key-highlight "..."] [--timeline-issue "..."] [--developers "Name1,Name2"]
 *        [--trend <_sprint_trend.json>] [--charts-dir <dir>] [--summary-out <file.docx>]
 *        [--full-only] [--summary-only]
 *
 * Manual-input fields and their defaults when omitted:
 *   --status           one of Completed / Delayed / Partial (case-insensitive) -> default "Completed"
 *   --key-highlight    free text                                              -> default "None"
 *   --timeline-issue   free text                                              -> default "None"
 *   --submission-date  YYYY-MM-DD                                             -> default today
 *   --prepared-by      free text                                              -> default "Gerald Nguyen"
 *   --developers       comma-separated names, matched against activity[].name -> default: everyone
 *                       with activity in the sprint. Only affects the "Activity by Developer" table
 *                       and workload chart (pass the SAME list to generate_charts.py's --developers
 *                       so the pre-rendered chart matches the table) — does not filter sprint-wide
 *                       stats (completion rate, bug ratio, features/bugs completed, etc).
 */
"use strict";
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, ShadingType, BorderStyle, PageOrientation, ImageRun,
} = require("docx");

const NAVY = "1E2761";
const ICE = "CADCFC";
const GREY = "F1EFE8";
const RED = "A32D2D";
const GREEN = "1E6B3C";
const TEXT_SEC = "5F5E5A";
const WHITE = "FFFFFF";
const BORDER_GREY = "D0D7DE";

// RAG (Red/Amber/Green) thresholds — confirmed with the business owner. "Better" direction differs
// per metric: higher completion rate is good; lower overdue count and lower bug ratio are good.
const RAG_COLORS = {
  green: { bg: "E6F4EA", text: "1E6B3C", label: "ON TRACK" },
  amber: { bg: "FFF4E0", text: "946200", label: "NEEDS ATTENTION" },
  red: { bg: "FBE9E7", text: "A32D2D", label: "AT RISK" },
};
const RAG_RANK = { green: 0, amber: 1, red: 2 };

function ragCompletionRate(pct){ return pct >= 80 ? "green" : pct >= 50 ? "amber" : "red"; }
function ragOverdue(n){ return n <= 2 ? "green" : n <= 5 ? "amber" : "red"; }
function ragBugRatio(pct){ return pct < 30 ? "green" : pct <= 50 ? "amber" : "red"; }
function worstRag(...rags){ return rags.reduce((a, b) => (RAG_RANK[b] > RAG_RANK[a] ? b : a)); }

// Delta vs the previous sprint. `higherIsBetter` flips the up/down-is-good color logic per metric.
// Includes the previous sprint's actual value alongside the delta (e.g. "was 70%") so the reader
// gets full context from this one card — without it, understanding a delta like "+6%" requires
// digging up the previous sprint's report just to know what it was a change from.
function deltaText(current, previous, opts = {}){
  const { unit = "", higherIsBetter = true, decimals = 0 } = opts;
  if (previous === undefined || previous === null) return { text: "No prior sprint to compare", color: TEXT_SEC };
  const prevStr = previous.toFixed(decimals) + unit;
  const diff = current - previous;
  if (Math.abs(diff) < Math.pow(10, -decimals) / 2) return { text: "No change vs last sprint (" + prevStr + ")", color: TEXT_SEC };
  const good = higherIsBetter ? diff > 0 : diff < 0;
  const arrow = diff > 0 ? "▲" : "▼";
  const sign = diff > 0 ? "+" : "";
  return { text: arrow + " " + sign + diff.toFixed(decimals) + unit + " vs last sprint (was " + prevStr + ")", color: good ? GREEN : RED };
}

// Auto-generates a short, data-driven Deployed Rate + Bug Ratio narrative from the trend series —
// same kind of read-out a reader would get from eyeballing the trend chart, but written out so they
// don't have to. Only called when trendSeries has >= 3 points (a 2-point "trend" is just one delta,
// not a real shape — see the ingest-workflow discussion this mirrors). Returns null otherwise.
function computeTrendInsights(trendSeries){
  if (!trendSeries || trendSeries.length < 3) return null;
  const first = trendSeries[0];
  const last = trendSeries[trendSeries.length - 1];
  const n = trendSeries.length;

  // --- Deployed Rate: overall direction + how many trailing sprints share the current RAG band ---
  const depDelta = last.deployedRate - first.deployedRate;
  const depDir = depDelta > 0 ? "improved" : depDelta < 0 ? "declined" : "held steady";
  const currentDepRag = ragCompletionRate(last.deployedRate);
  let trailDep = 1;
  for (let i = trendSeries.length - 2; i >= 0; i--){
    if (ragCompletionRate(trendSeries[i].deployedRate) === currentDepRag) trailDep++; else break;
  }
  const depBandLabel = currentDepRag === "green" ? "Green (≥80%)" : currentDepRag === "amber" ? "Amber (50-79%)" : "Red (<50%)";
  let depSentence = "Deployed Rate has " + depDir + " from " + first.deployedRate + "% to " + last.deployedRate + "% over the last " + n + " sprints";
  if (trailDep >= 2) depSentence += ", holding in the " + depBandLabel + " band for " + trailDep + " sprints in a row";
  depSentence += ".";

  // --- Bug Ratio: overall direction + call out the single biggest sprint-over-sprint regression ---
  const bugDelta = last.bugRatio - first.bugRatio; // negative = improvement (lower ratio is better)
  const bugDir = bugDelta < 0 ? "improved" : bugDelta > 0 ? "worsened" : "held steady";
  let worst = null;
  for (let i = 1; i < trendSeries.length; i++){
    const diff = trendSeries[i].bugRatio - trendSeries[i - 1].bugRatio;
    if (diff > 0 && (!worst || diff > worst.diff)) worst = { diff, folder: trendSeries[i].folder, value: trendSeries[i].bugRatio };
  }
  let bugSentence = "Bug Ratio has " + bugDir + " overall, from " + first.bugRatio + "% to " + last.bugRatio + "%";
  if (worst && worst.diff >= 10 && worst.folder !== last.folder){
    bugSentence += ", though it spiked back up to " + worst.value + "% in " + worst.folder + " before improving again";
  }
  bugSentence += ".";

  return { depSentence, bugSentence };
}

const OPEN_STATES = ["New", "Active", "In Review"];
const CLOSED_STATES = ["Resolved", "Closed"];

const PAGE_WIDTH_DXA = 12240; // US Letter portrait
const MARGIN_DXA = 720;
const CONTENT_WIDTH_DXA = PAGE_WIDTH_DXA - MARGIN_DXA * 2;

function cellBorder(){
  return {
    top: { style: BorderStyle.SINGLE, size: 2, color: BORDER_GREY },
    bottom: { style: BorderStyle.SINGLE, size: 2, color: BORDER_GREY },
    left: { style: BorderStyle.SINGLE, size: 2, color: BORDER_GREY },
    right: { style: BorderStyle.SINGLE, size: 2, color: BORDER_GREY },
  };
}

function sectionTitle(text){
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 120 }, children: [new TextRun({ text, bold: true, color: NAVY, size: 30 })] });
}
function subsectionTitle(text){
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 }, children: [new TextRun({ text, bold: true, color: NAVY, size: 24 })] });
}
function bodyText(text, opts = {}){
  return new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, size: 20, color: opts.color || "222222", italics: !!opts.italic })] });
}

// `sub` may be a single string (one plain caption line), an array of strings (one Paragraph per
// line, all plain), or an array where an element is instead `{ runs: [{ text, bold, color }, ...] }`
// for a line that needs mixed formatting (e.g. a label in normal weight followed by bolded figures)
// — used by the Carried Over card's total line and its Pending Deployment / In Progress / Not
// Started breakdown.
function statCard(label, value, sub, widthDxa){
  const subLines = Array.isArray(sub) ? sub : [sub];
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    borders: cellBorder(),
    shading: { type: ShadingType.CLEAR, fill: GREY },
    margins: { top: 200, bottom: 200, left: 150, right: 150 },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: label, bold: true, size: 18, color: TEXT_SEC })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 }, children: [new TextRun({ text: String(value), bold: true, size: 44, color: NAVY })] }),
      ...subLines.map(line => {
        const runs = (line && typeof line === "object" && line.runs)
          ? line.runs.map(r => new TextRun({ text: r.text, size: 16, color: r.color || TEXT_SEC, bold: !!r.bold }))
          : [new TextRun({ text: line, size: 16, color: TEXT_SEC })];
        return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 20 }, children: runs });
      }),
    ],
  });
}

// KPI card colored by RAG status, with an optional delta-vs-last-sprint line and an optional plain
// caption line underneath (e.g. "155 ticket(s) are Closed, deployed to production").
function kpiCard(label, value, rag, deltaInfo, widthDxa, caption){
  const colors = RAG_COLORS[rag];
  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: label, bold: true, size: 18, color: TEXT_SEC })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: (deltaInfo || caption) ? 40 : 80 }, children: [new TextRun({ text: String(value), bold: true, size: 44, color: colors.text })] }),
  ];
  if (deltaInfo){
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: caption ? 40 : 0 }, children: [new TextRun({ text: deltaInfo.text, size: 15, color: deltaInfo.color, bold: true })] }));
  }
  if (caption){
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: caption, size: 16, color: TEXT_SEC })] }));
  }
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    borders: cellBorder(),
    shading: { type: ShadingType.CLEAR, fill: colors.bg },
    margins: { top: 200, bottom: 200, left: 150, right: 150 },
    children,
  });
}

// Full-width banner cell for the overall Sprint Health verdict.
function healthBanner(rag, widthDxa){
  const colors = RAG_COLORS[rag];
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    borders: cellBorder(),
    shading: { type: ShadingType.CLEAR, fill: colors.bg },
    margins: { top: 140, bottom: 140, left: 150, right: 150 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SPRINT HEALTH: " + colors.label, bold: true, size: 24, color: colors.text })] })],
  });
}

// Embeds a PNG at a target display width (inches); height follows from the image's own pixel
// aspect ratio so charts never look stretched. Returns null (renders nothing) if the file is
// missing, so callers can skip a section gracefully rather than crash the whole report.
function imageParagraph(pngPath, widthIn, opts = {}){
  if (!pngPath || !fs.existsSync(pngPath)) return null;
  const dims = readPngDimensions(pngPath);
  const widthPx = Math.round(widthIn * 96);
  const heightPx = dims ? Math.round(widthPx * (dims.height / dims.width)) : Math.round(widthPx * 0.6);
  return new Paragraph({
    alignment: opts.align || AlignmentType.CENTER,
    spacing: { before: opts.before || 100, after: opts.after || 100 },
    children: [new ImageRun({ type: "png", data: fs.readFileSync(pngPath), transformation: { width: widthPx, height: heightPx } })],
  });
}

// Minimal PNG header parser (no extra dependency) — width/height live in the IHDR chunk at a fixed
// offset for every valid PNG.
function readPngDimensions(pngPath){
  try {
    const buf = fs.readFileSync(pngPath);
    if (buf.length < 24 || buf.toString("ascii", 1, 4) !== "PNG") return null;
    return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
  } catch (e){
    return null;
  }
}

function headerCell(text, widthDxa, align){
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    borders: cellBorder(),
    shading: { type: ShadingType.CLEAR, fill: NAVY },
    margins: { top: 100, bottom: 100, left: 100, right: 100 },
    children: [new Paragraph({ alignment: align || AlignmentType.LEFT, children: [new TextRun({ text, bold: true, size: 17, color: WHITE })] })],
  });
}

function dataCell(text, widthDxa, opts = {}){
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    borders: cellBorder(),
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill } : undefined,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    columnSpan: opts.colSpan || undefined,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), size: 17, bold: !!opts.bold, color: opts.color || "1A1A1A" })],
    })],
  });
}

function labelCell(text, widthDxa){
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    borders: cellBorder(),
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 17, color: "1A1A1A" })] })],
  });
}

// ---------------------------------------------------------------
// Work item table (Overdue / Active / Completed) — mirrors renderWorkItemTable()
// ---------------------------------------------------------------
const WI_COLS = [
  { label: "Work Item", width: 3400 },
  { label: "Type", width: 900 },
  { label: "State", width: 900 },
  { label: "Assigned To", width: 1500 },
  { label: "Start", width: 900 },
  { label: "Target", width: 900 },
  { label: "Changed", width: 1100 },
  { label: "Effort", width: 700 },
];

function workItemTable(items){
  if (!items.length){
    return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "No work items in this category for the selected filters.", italics: true, size: 18, color: TEXT_SEC })] });
  }
  const sorted = items.slice().sort((a, b) => (a.id > b.id ? 1 : -1));
  const headerRow = new TableRow({ cantSplit: true, tableHeader: true, children: WI_COLS.map(c => headerCell(c.label, c.width)) });
  const rows = sorted.map((it, idx) => {
    const fill = idx % 2 === 1 ? "F7F8FA" : undefined;
    const assignee = it.assignee ? it.assignee.replace(" - ONESECURE", "") : "-";
    const changed = it.changed ? (it.changed + (it.changedTime ? " " + it.changedTime : "")) : "-";
    return new TableRow({
      cantSplit: true,
      children: [
        dataCell("#" + it.id + ": " + it.title, WI_COLS[0].width, { fill, bold: true }),
        dataCell(it.type || "-", WI_COLS[1].width, { fill }),
        dataCell(it.state || "-", WI_COLS[2].width, { fill }),
        dataCell(assignee, WI_COLS[3].width, { fill }),
        dataCell(it.start || "-", WI_COLS[4].width, { fill }),
        dataCell(it.target || "-", WI_COLS[5].width, { fill }),
        dataCell(changed, WI_COLS[6].width, { fill }),
        dataCell(typeof it.effort === "number" ? String(it.effort) : "-", WI_COLS[7].width, { fill, align: AlignmentType.CENTER }),
      ],
    });
  });
  return new Table({ width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA }, columnWidths: WI_COLS.map(c => c.width), rows: [headerRow, ...rows] });
}

// ---------------------------------------------------------------
// Main
// ---------------------------------------------------------------
const BOOLEAN_FLAGS = new Set(["full-only", "summary-only"]);

function parseArgs(argv){
  const [dataPath, outPath, ...rest] = argv;
  const opts = {};
  for (let i = 0; i < rest.length; i++){
    if (rest[i].startsWith("--")){
      const key = rest[i].slice(2);
      if (BOOLEAN_FLAGS.has(key)){
        opts[key] = true;
      } else {
        opts[key] = rest[i + 1];
        i++;
      }
    }
  }
  return { dataPath, outPath, opts };
}

// Builds the document body. Sprint Overview and Sprint Summary are always included in both variants;
// `includeFull` only toggles Sprint Details (the raw Overdue/Active/Completed work-item tables) —
// the "full" report has it, the "summary" report omits it so CEO/CTO/PM/Business Lead audiences get
// the headline numbers (Sprint Overview + Executive Summary + Sprint Health + charts + Activity by
// Developer + Trend Across Sprints) without the raw work-item tables. Confirmation has been removed
// entirely (from both variants).
function buildChildren(ctx, includeFull){
  const {
    data, items, stats, activity, sprintPeriod, todayIso, preparedBy, submissionDate, sprintName,
    keyHighlight, timelineIssue, chartsDir, trendSeries, previousTrendEntry, sprintStatus,
  } = ctx;

  // --- Sprint Health KPIs (deployed rate / bug ratio), RAG-colored ---
  // Note: Overdue Items is intentionally excluded from Sprint Summary's stats/health scoring (per
  // request) — it remains available in the full report's Sprint Details -> Overdue Tasks section.
  // "Deployed to Production" (Closed only) replaced the old blended "Completion Rate" (Closed +
  // Resolved) — Closed means shipped to production; Resolved means done in dev and sitting on
  // staging, still awaiting deploy approval, which is a materially different state (confirmed with
  // the business owner).
  // "Carried Over to Next Sprint" groups everything not yet Closed — a Resolved ticket may still get
  // deployed mid-sprint, so it's carried over too, just further along than an Active/New one. It has
  // 3 sub-parts, shown in its caption: Pending Deployment (Resolved), In Progress (Active/In Review),
  // Not Started (New). Pending Deployment no longer gets its own top-level card — it lives here.
  const deployedRate = stats.total > 0 ? Math.round((stats.deployedCount / stats.total) * 100) : 0;
  const carriedOverCount = stats.pendingDeploymentCount + stats.open;
  const carriedOverRate = stats.total > 0 ? Math.round((carriedOverCount / stats.total) * 100) : 0;
  // Each Carried Over sub-part also gets its own % of the sprint total (they sum to carriedOverRate).
  const pendingDeploymentPct = stats.total > 0 ? Math.round((stats.pendingDeploymentCount / stats.total) * 100) : 0;
  const inProgressPct = stats.total > 0 ? Math.round((stats.inProgressCount / stats.total) * 100) : 0;
  const notStartedPct = stats.total > 0 ? Math.round((stats.notStartedCount / stats.total) * 100) : 0;
  const ragDeployed = ragCompletionRate(deployedRate);
  const ragBug = ragBugRatio(stats.bugRatio);
  const overallRag = worstRag(ragDeployed, ragBug);

  const deltaDeployed = previousTrendEntry ? deltaText(deployedRate, previousTrendEntry.deployedRate, { unit: "%", higherIsBetter: true }) : null;
  const deltaBugRatio = previousTrendEntry ? deltaText(stats.bugRatio, previousTrendEntry.bugRatio, { unit: "%", higherIsBetter: false }) : null;

  const children = [];

  // --- Title + meta (always shown, in both variants) ---
  children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun({ text: includeFull ? "Sprint Report" : "Sprint Summary", bold: true, color: NAVY, size: 48 })] }));
  children.push(new Paragraph({
    spacing: { before: 100, after: 200 },
    children: [
      new TextRun({ text: "Reporting Period: ", bold: true, size: 20 }),
      new TextRun({ text: (sprintPeriod ? sprintPeriod.start : "—") + " – " + (sprintPeriod ? sprintPeriod.end : "—"), size: 20 }),
      new TextRun({ text: "     Prepared by: ", bold: true, size: 20 }),
      new TextRun({ text: preparedBy, size: 20 }),
      new TextRun({ text: "     Submission Date: ", bold: true, size: 20 }),
      new TextRun({ text: submissionDate, size: 20 }),
    ],
  }));

  // --- Sprint Overview (always shown, in both variants) ---
  children.push(sectionTitle("Sprint Overview"));
  const ovW = [CONTENT_WIDTH_DXA / 3, CONTENT_WIDTH_DXA / 3, CONTENT_WIDTH_DXA / 3];
  children.push(new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: ovW,
    rows: [
      new TableRow({ children: [headerCell("Sprint", ovW[0]), headerCell("Sprint Period", ovW[1]), headerCell("Status", ovW[2])] }),
      new TableRow({ children: [dataCell(sprintName, ovW[0], { bold: true }), dataCell((sprintPeriod ? sprintPeriod.start : "—") + " – " + (sprintPeriod ? sprintPeriod.end : "—"), ovW[1]), dataCell(sprintStatus, ovW[2])] }),
      new TableRow({ children: [labelCell("Key Highlight:", ovW[0]), dataCell(keyHighlight, ovW[1] + ovW[2], { colSpan: 2 })] }),
      new TableRow({ children: [labelCell("Timeline Issue:", ovW[0]), dataCell(timelineIssue, ovW[1] + ovW[2], { colSpan: 2 })] }),
    ],
  }));

  // --- Sprint Summary (always shown, in both variants) ---
  // Combined into one block: the health banner + rate KPIs (Deployed / Carried Over / Bug Ratio)
  // come first as the headline, followed by a compact by-work-type table breaking those same Closed
  // vs Resolved counts down by Features / Technical Debts / Bugs — one shared legend line covers
  // both instead of repeating the Closed/Resolved explanation in two separate sections.
  children.push(sectionTitle("Sprint Summary"));
  children.push(subsectionTitle("1. Executive Summary"));
  children.push(new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH_DXA],
    rows: [new TableRow({ children: [healthBanner(overallRag, CONTENT_WIDTH_DXA)] })],
  }));
  children.push(new Paragraph({ spacing: { before: 80 }, children: [] }));
  const kpiThirdW = Math.floor(CONTENT_WIDTH_DXA / 3);
  const kpiLastW = CONTENT_WIDTH_DXA - kpiThirdW * 2;
  children.push(new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [kpiThirdW, kpiThirdW, kpiLastW],
    rows: [new TableRow({ children: [
      kpiCard("Deployed to Production", deployedRate + "%", ragDeployed, deltaDeployed, kpiThirdW, stats.deployedCount + " ticket(s) are Closed, deployed to production"),
      statCard("Carried Over to Next Sprint", carriedOverRate + "%", [
        { runs: [{ text: carriedOverCount + " ticket(s) not yet Closed", bold: true, color: NAVY }] },
        { runs: [
          { text: "Pending Deployment (Resolved): " },
          { text: stats.pendingDeploymentCount + " tickets - " + pendingDeploymentPct + "%", bold: true },
        ] },
        { runs: [
          { text: "In Progress (Active / In Review): " },
          { text: stats.inProgressCount + " tickets - " + inProgressPct + "%", bold: true },
        ] },
        { runs: [
          { text: "Not Started (New): " },
          { text: stats.notStartedCount + " tickets - " + notStartedPct + "%", bold: true },
        ] },
      ], kpiThirdW),
      kpiCard("Bug Ratio", stats.bugRatio + "%", ragBug, deltaBugRatio, kpiLastW),
    ] })],
  }));
  children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [] }));

  // --- Completed work, by type (Closed vs Resolved) — compact table, not 6 separate cards ---
  const ctW0 = Math.floor(CONTENT_WIDTH_DXA * 0.34);
  const ctRest = CONTENT_WIDTH_DXA - ctW0;
  const ctW1 = Math.floor(ctRest / 3);
  const ctW = [ctW0, ctW1, ctW1, ctRest - ctW1 * 2];
  const completedRows = [
    { label: "Features Completed", closed: stats.featuresCompletedClosed, resolved: stats.featuresCompletedResolved },
    { label: "Technical Debts Completed", closed: stats.techDebtCompletedClosed, resolved: stats.techDebtCompletedResolved },
    { label: "Bugs Fixed", closed: stats.bugsFixedClosed, resolved: stats.bugsFixedResolved },
  ];
  const ctTotalClosed = completedRows.reduce((s, r) => s + r.closed, 0);
  const ctTotalResolved = completedRows.reduce((s, r) => s + r.resolved, 0);
  children.push(new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: ctW,
    rows: [
      new TableRow({ tableHeader: true, children: [
        headerCell("Completed Work (by Type)", ctW[0]),
        headerCell("Closed", ctW[1], AlignmentType.CENTER),
        headerCell("Resolved", ctW[2], AlignmentType.CENTER),
        headerCell("Total", ctW[3], AlignmentType.CENTER),
      ] }),
      ...completedRows.map((r, idx) => new TableRow({ children: [
        dataCell(r.label, ctW[0], { fill: idx % 2 === 1 ? "F7F8FA" : undefined, bold: true }),
        dataCell(String(r.closed), ctW[1], { fill: idx % 2 === 1 ? "F7F8FA" : undefined, align: AlignmentType.CENTER }),
        dataCell(String(r.resolved), ctW[2], { fill: idx % 2 === 1 ? "F7F8FA" : undefined, align: AlignmentType.CENTER }),
        dataCell(String(r.closed + r.resolved), ctW[3], { fill: idx % 2 === 1 ? "F7F8FA" : undefined, align: AlignmentType.CENTER }),
      ] })),
      new TableRow({ children: [
        dataCell("Total", ctW[0], { fill: ICE, bold: true }),
        dataCell(String(ctTotalClosed), ctW[1], { fill: ICE, bold: true, align: AlignmentType.CENTER }),
        dataCell(String(ctTotalResolved), ctW[2], { fill: ICE, bold: true, align: AlignmentType.CENTER }),
        dataCell(String(ctTotalClosed + ctTotalResolved), ctW[3], { fill: ICE, bold: true, align: AlignmentType.CENTER }),
      ] }),
    ],
  }));
  children.push(bodyText("Closed = deployed to production. Resolved = completed on staging, awaiting deploy approval.", { italic: true, color: TEXT_SEC }));
  children.push(new Paragraph({ spacing: { before: 160, after: 60 }, children: [] }));

  // --- Visual breakdown: two donuts side by side (pre-rendered PNGs; skipped if not supplied) ---
  if (chartsDir){
    const byTypeImg = imageParagraph(path.join(chartsDir, "byType_donut.png"), 3.5);
    const openClosedImg = imageParagraph(path.join(chartsDir, "open_closed_donut.png"), 3.0);
    if (byTypeImg || openClosedImg){
      const halfW = Math.floor(CONTENT_WIDTH_DXA / 2);
      children.push(new Table({
        width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
        columnWidths: [halfW, CONTENT_WIDTH_DXA - halfW],
        rows: [new TableRow({ children: [
          new TableCell({ width: { size: halfW, type: WidthType.DXA }, borders: { top: {style:BorderStyle.NONE}, bottom: {style:BorderStyle.NONE}, left: {style:BorderStyle.NONE}, right: {style:BorderStyle.NONE} }, children: [byTypeImg || new Paragraph({ text: "" })] }),
          new TableCell({ width: { size: CONTENT_WIDTH_DXA - halfW, type: WidthType.DXA }, borders: { top: {style:BorderStyle.NONE}, bottom: {style:BorderStyle.NONE}, left: {style:BorderStyle.NONE}, right: {style:BorderStyle.NONE} }, children: [openClosedImg || new Paragraph({ text: "" })] }),
        ] })],
      }));
    }
  }

  children.push(subsectionTitle("2. Activity by Developer"));
  if (chartsDir){
    const workloadImg = imageParagraph(path.join(chartsDir, "workload_bar.png"), 7.5);
    if (workloadImg) children.push(workloadImg);
  }
  if (!activity.length){
    children.push(bodyText("No activity for the selected assignees in this period.", { italic: true, color: TEXT_SEC }));
  } else {
    const devW = [2600, 1200, 1200, 1200, 1200];
    const devTotalW = devW.reduce((a, b) => a + b, 0);
    const scale = CONTENT_WIDTH_DXA / devTotalW;
    const w = devW.map(x => Math.floor(x * scale));
    const rows = [new TableRow({ cantSplit: true, tableHeader: true, children: [headerCell("Developer", w[0]), headerCell("Active", w[1], AlignmentType.CENTER), headerCell("Resolved", w[2], AlignmentType.CENTER), headerCell("Closed", w[3], AlignmentType.CENTER), headerCell("Effort", w[4], AlignmentType.CENTER)] })];
    let totActive = 0, totResolved = 0, totClosed = 0, totEffort = 0;
    activity.forEach((d, idx) => {
      totActive += d.active; totResolved += d.resolved; totClosed += d.closed; totEffort += d.effort;
      const fill = idx % 2 === 1 ? "F7F8FA" : undefined;
      rows.push(new TableRow({ children: [
        dataCell(d.name, w[0], { fill, bold: true }),
        dataCell(String(d.active), w[1], { fill, align: AlignmentType.CENTER }),
        dataCell(String(d.resolved), w[2], { fill, align: AlignmentType.CENTER }),
        dataCell(String(d.closed), w[3], { fill, align: AlignmentType.CENTER }),
        dataCell(String(d.effort), w[4], { fill, align: AlignmentType.CENTER }),
      ] }));
    });
    rows.push(new TableRow({ children: [
      dataCell("Totals", w[0], { bold: true, fill: ICE }),
      dataCell(String(totActive), w[1], { bold: true, fill: ICE, align: AlignmentType.CENTER }),
      dataCell(String(totResolved), w[2], { bold: true, fill: ICE, align: AlignmentType.CENTER }),
      dataCell(String(totClosed), w[3], { bold: true, fill: ICE, align: AlignmentType.CENTER }),
      dataCell(String(totEffort), w[4], { bold: true, fill: ICE, align: AlignmentType.CENTER }),
    ] }));
    children.push(new Table({ width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA }, columnWidths: w, rows }));
  }

  // --- Trend Across Sprints ---
  children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
  children.push(subsectionTitle("3. Trend Across Sprints"));
  if (!trendSeries || trendSeries.length < 2){
    children.push(bodyText("Not enough sprint history yet to chart a trend (need at least 2 sprints with ingested data).", { italic: true, color: TEXT_SEC }));
  } else {
    children.push(bodyText("Velocity, deployed rate, and bug ratio across all " + trendSeries.length + " sprints on record.", { italic: true, color: TEXT_SEC }));
    const trendImg = chartsDir ? imageParagraph(path.join(chartsDir, "trend_chart.png"), 7.5) : null;
    if (trendImg){
      children.push(trendImg);
    } else {
      // Fallback when no chart image was generated: a compact numeric table, still useful on its own.
      const tW = [1800, 1400, 1400, 1400, 1200];
      const tTotalW = tW.reduce((a, b) => a + b, 0);
      const tScale = CONTENT_WIDTH_DXA / tTotalW;
      const w2 = tW.map(x => Math.floor(x * tScale));
      const rows = [new TableRow({ cantSplit: true, tableHeader: true, children: [
        headerCell("Sprint", w2[0]), headerCell("Completed", w2[1], AlignmentType.CENTER), headerCell("Deployed %", w2[2], AlignmentType.CENTER),
        headerCell("Bug Ratio %", w2[3], AlignmentType.CENTER), headerCell("Total", w2[4], AlignmentType.CENTER),
      ] })];
      trendSeries.forEach((s, idx) => {
        const fill = s.folder === sprintName ? ICE : (idx % 2 === 1 ? "F7F8FA" : undefined);
        rows.push(new TableRow({ children: [
          dataCell(s.folder, w2[0], { fill, bold: s.folder === sprintName }),
          dataCell(String(s.totalCompleted), w2[1], { fill, align: AlignmentType.CENTER }),
          dataCell(s.deployedRate + "%", w2[2], { fill, align: AlignmentType.CENTER }),
          dataCell(s.bugRatio + "%", w2[3], { fill, align: AlignmentType.CENTER }),
          dataCell(String(s.total), w2[4], { fill, align: AlignmentType.CENTER }),
        ] }));
      });
      children.push(new Table({ width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA }, columnWidths: w2, rows }));
    }

    // --- Trend Insights: short auto-generated narrative, not just the raw chart/table ---
    const insights = computeTrendInsights(trendSeries);
    if (insights){
      children.push(new Paragraph({ spacing: { before: 140, after: 40 }, children: [new TextRun({ text: "Trend Insights", bold: true, size: 19, color: NAVY })] }));
      children.push(bodyText(insights.depSentence, { color: "333333" }));
      children.push(bodyText(insights.bugSentence, { color: "333333" }));
    }
  }

  if (includeFull){
    // --- Sprint Details ---
    children.push(sectionTitle("Sprint Details"));

    const overdue = stats.overdue || [];
    children.push(subsectionTitle("1. Overdue Tasks (" + overdue.length + ")"));
    children.push(bodyText("Open items whose Target date is before " + (data.generatedDateRange ? data.generatedDateRange.max : todayIso) + " (the most recent ingested daily report).", { italic: true, color: TEXT_SEC }));
    children.push(workItemTable(overdue));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    const activeItems = items.filter(i => OPEN_STATES.includes(i.state));
    children.push(subsectionTitle("2. Active Work Items (" + activeItems.length + ")"));
    children.push(workItemTable(activeItems));
    children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));

    const completedItems = items.filter(i => CLOSED_STATES.includes(i.state));
    children.push(subsectionTitle("3. Completed Work Items (" + completedItems.length + ")"));
    children.push(workItemTable(completedItems));
  }

  return children;
}

async function buildAndWrite(ctx, outPath, includeFull){
  const children = buildChildren(ctx, includeFull);
  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: PAGE_WIDTH_DXA, height: 15840 }, margin: { top: MARGIN_DXA, bottom: MARGIN_DXA, left: MARGIN_DXA, right: MARGIN_DXA } } },
      children,
    }],
  });
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buffer);
  console.log("Wrote " + outPath + " (" + buffer.length + " bytes)");
}

// Derives "<name>_Summary.<ext>" from "<name>.<ext>" for the default summary-report path.
function deriveSummaryPath(outPath){
  const ext = path.extname(outPath);
  const base = outPath.slice(0, outPath.length - ext.length);
  return base + "_Summary" + ext;
}

async function main(){
  const { dataPath, outPath, opts } = parseArgs(process.argv.slice(2));
  if (!dataPath || !outPath){
    console.error("Usage: node build_sprint_report_docx.js <sprint_data.json> <output.docx> [--status Completed|Delayed|Partial] [--prepared-by NAME] [--submission-date YYYY-MM-DD] [--sprint-name NAME] [--key-highlight TEXT] [--timeline-issue TEXT] [--developers \"Name1,Name2\"] [--trend <_sprint_trend.json>] [--charts-dir <dir>] [--summary-out <file.docx>] [--full-only] [--summary-only]");
    process.exit(1);
  }
  const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
  const items = data.items || [];
  const stats = data.stats;
  // --developers restricts the "Activity by Developer" table/chart to a chosen subset (comma-
  // separated names, matched against activity[].name) — a manual input, since which developers are
  // relevant to show is a judgment call, not something inferable from the data. Omit the flag (or
  // pass an empty value) to include everyone with activity in the sprint, which is also the fallback
  // if every requested name fails to match (keeps the report from silently going blank on a typo).
  // Sprint-wide stats (completion rate, bug ratio, features/bugs completed, etc.) are NOT filtered by
  // this — they always reflect the whole sprint regardless of which developers are shown here.
  let activity = data.activity || [];
  if (opts.developers){
    const wanted = String(opts.developers).split(",").map(s => s.trim()).filter(Boolean);
    if (wanted.length){
      const filtered = activity.filter(d => wanted.some(w => w.toLowerCase() === d.name.toLowerCase()));
      if (filtered.length) activity = filtered;
      else console.error("Warning: none of --developers \"" + opts.developers + "\" matched a known developer — showing all developers instead.");
    }
  }
  const sprintPeriod = data.sprintPeriod;
  const todayIso = new Date().toISOString().slice(0, 10);

  const preparedBy = opts["prepared-by"] || "Gerald Nguyen";
  const submissionDate = opts["submission-date"] || todayIso;
  const sprintName = opts["sprint-name"] || path.basename(path.dirname(path.resolve(dataPath)));
  const keyHighlight = opts["key-highlight"] || "None";
  const timelineIssue = opts["timeline-issue"] || "None";
  const chartsDir = opts["charts-dir"] || null;

  // Status is a manual input (there's no reliable way to infer "Delayed" vs "Partial" from the
  // ingested data alone) — restricted to this fixed set, case-insensitively, defaulting to
  // "Completed" when omitted or unrecognized.
  const STATUS_VALUES = ["Completed", "Delayed", "Partial"];
  const rawStatus = opts["status"];
  const matchedStatus = rawStatus ? STATUS_VALUES.find(v => v.toLowerCase() === String(rawStatus).toLowerCase()) : null;
  if (rawStatus && !matchedStatus){
    console.error("Warning: --status \"" + rawStatus + "\" is not one of Completed/Delayed/Partial — defaulting to Completed.");
  }
  const sprintStatus = matchedStatus || "Completed";

  // Trend data (optional): `trendSeries` (capped, e.g. last 6 sprints) drives the "Trend Across
  // Sprints" chart/table so it stays readable as history grows. The "vs previous sprint" delta on
  // the KPI row instead looks up `fullSeries` (every sprint on record, uncapped) so that lookup
  // stays correct even once history exceeds the display cap. Missing/absent trend file just means
  // these bits are skipped — the rest of the report renders exactly the same either way.
  let trendSeries = null, previousTrendEntry = null;
  if (opts.trend && fs.existsSync(opts.trend)){
    const trendData = JSON.parse(fs.readFileSync(opts.trend, "utf8"));
    trendSeries = trendData.series || [];
    const fullSeries = trendData.fullSeries || trendSeries;
    const idx = fullSeries.findIndex(s => s.folder === sprintName);
    if (idx > 0) previousTrendEntry = fullSeries[idx - 1];
  }

  const ctx = {
    data, items, stats, activity, sprintPeriod, todayIso, preparedBy, submissionDate, sprintName,
    keyHighlight, timelineIssue, chartsDir, trendSeries, previousTrendEntry, sprintStatus,
  };

  // By default, every run produces BOTH the full report and a Sprint-Summary-only version.
  // --full-only / --summary-only restrict to just one of the two.
  const summaryOutPath = opts["summary-out"] || deriveSummaryPath(outPath);
  const doFull = !opts["summary-only"];
  const doSummary = !opts["full-only"];

  if (doFull) await buildAndWrite(ctx, outPath, true);
  if (doSummary) await buildAndWrite(ctx, summaryOutPath, false);
}

main().catch(err => { console.error("ERROR:", err); process.exit(1); });
