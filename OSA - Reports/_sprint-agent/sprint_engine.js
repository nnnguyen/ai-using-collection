#!/usr/bin/env node
/**
 * sprint_engine.js — headless (no browser) port of Sprint_Report_Generator.html's parsing +
 * aggregation logic, so an Agent can run it directly via a shell command against whatever daily
 * report files sit inside one sprint's folder — no HTML form involved.
 *
 * All parsing/merge/stat logic below is copied verbatim from Sprint_Report_Generator.html
 * (parseWorkItemsFromText, parseWorkItemsFromSheetRows, parseWorkItemsFromHtml, mergeAllItems,
 * computeStats, computeDeveloperActivity) — only the file-reading layer changed, from
 * browser File/DOMParser/pdfjsLib-via-CDN/mammoth.js-via-CDN/SheetJS-via-CDN to Node equivalents
 * (fs, jsdom's DOMParser, pdfjs-dist, the `mammoth` npm package, the `xlsx` npm package).
 *
 * Setup (once per machine/session — see package.json in this folder):
 *   npm install
 *
 * Sprint folders are named `Sprint-<StartYYYYMMDD>_<EndYYYYMMDD>` (e.g. "Sprint-20260824_20260906").
 * When a folder ingested matches this pattern, the start/end dates are read straight from the
 * folder name (authoritative) instead of being inferred from file content.
 *
 * Daily report files are named `Webyith-Daily-Report_<YYYY_MM_DD>`. Because sprints run
 * back-to-back (next sprint's start = this sprint's end + 1), the first daily file of a sprint is
 * dated (sprint start + 1) and the last is dated (sprint end - 2) — the sprint's own last 2
 * calendar days have no daily file. `ingest` uses this to report which expected dates are missing.
 *
 * Usage:
 *   node sprint_engine.js ingest <folder> [--out <file.json>]
 *                                                 Parse every report file in <folder>, merge
 *                                                 across days, write aggregated stats + items to
 *                                                 <folder>/_sprint_data.json (or --out path).
 *   node sprint_engine.js list-sprints <baseFolder>
 *                                                 List subfolders of <baseFolder> matching the
 *                                                 Sprint-YYYYMMDD_YYYYMMDD naming convention, with
 *                                                 parsed start/end dates and which one (if any)
 *                                                 contains today's date.
 *   node sprint_engine.js parse-one <file>       Parse a single file, print its raw items as JSON.
 *                                                  (debugging / per-file inspection)
 */

"use strict";
const fs = require("fs");
const path = require("path");
const { JSDOM } = require("jsdom");
const mammoth = require("mammoth");
const XLSX = require("xlsx");
// pdfjs-dist's legacy Node build tries to polyfill DOMMatrix/Path2D on load and prints a noisy
// warning (straight to stdout, not stderr) when it can't find the optional `canvas` package —
// harmless for text-only extraction (no page rendering happens here), but it corrupts our JSON
// output if left alone. Pre-defining no-op stubs skips the polyfill attempt entirely.
if (typeof global.DOMMatrix === "undefined") global.DOMMatrix = function DOMMatrixStub(){};
if (typeof global.Path2D === "undefined") global.Path2D = function Path2DStub(){};
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.js");

const { window: jsdomWindow } = new JSDOM("");
const DOMParser = jsdomWindow.DOMParser;

const OPEN_STATES = ["New", "Active", "In Review"];
const CLOSED_STATES = ["Resolved", "Closed"];

/* ============================================================
   TEXT PARSER — for PDF/DOCX exports flattened to plain text
   (verbatim port of parseWorkItemsFromText)
   ============================================================ */
function parseWorkItemsFromText(rawText){
  let text = rawText || "";
  text = text.replace(/(\d{4})-\s+(\d{2}-\d{2})/g, "$1-$2");
  text = text.replace(/[ \t\r\n]+/g, " ");

  const genMatch = text.match(/Generated\s+(\d{4}-\d{2}-\d{2})/i);
  const generatedDate = genMatch ? genMatch[1] : null;

  const TYPE_STATE_RE = /\b(Bug|Epic|Feature|Task|User\s+Story|Issue)\s+(New|Active|In\s+Review|Resolved|Closed)\b/;
  const BREADCRUMB_RE = /Webyith\s*(\\[^·]*)?·/;
  const ASSIGNEE_RE = /([A-Z][A-Za-z.'\- ]*?)\s*-\s*ONESECURE/;

  const anchorRe = /#(\d{3,6}):/g;
  const anchors = [];
  let m;
  while ((m = anchorRe.exec(text)) !== null) anchors.push({ id: m[1], afterIdx: m.index + m[0].length });

  const items = [];
  for (let i = 0; i < anchors.length; i++){
    const start = anchors[i].afterIdx;
    const end = i + 1 < anchors.length ? anchors[i+1].afterIdx - (String(anchors[i+1].id).length + 2) : text.length;
    const win = text.slice(start, end);

    const typeStateMatch = TYPE_STATE_RE.exec(win);
    if (!typeStateMatch) continue;

    let preType = win.slice(0, typeStateMatch.index);
    let ticketIdCol = null, parentCol = null;
    const ticketParentMatch = preType.match(/(\d{3,6})\s+(-|#\d{3,6})\s*$/);
    if (ticketParentMatch){
      ticketIdCol = ticketParentMatch[1];
      parentCol = ticketParentMatch[2] === "-" ? null : ticketParentMatch[2].replace("#", "");
      preType = preType.slice(0, ticketParentMatch.index);
    }

    const breadcrumbMatch = BREADCRUMB_RE.exec(preType);
    const titleEnd = breadcrumbMatch ? breadcrumbMatch.index : preType.length;
    let title = preType.slice(0, titleEnd).trim().replace(/\s+/g, " ");
    if (title.length > 220) title = title.slice(0,220) + "…";
    if (!title) continue;

    const type = typeStateMatch[1].replace(/\s+/g," ");
    const stateVal = typeStateMatch[2].replace(/\s+/g," ");

    const afterTS = win.slice(typeStateMatch.index + typeStateMatch[0].length);
    const assigneeMatch = ASSIGNEE_RE.exec(afterTS);
    if (!assigneeMatch) continue;
    const assignee = (assigneeMatch[1].trim() + " - ONESECURE").replace(/\s+/g," ");

    let tail = afterTS.slice(assigneeMatch.index + assigneeMatch[0].length);
    const linkIdx = tail.search(/\[link\]/i);
    if (linkIdx !== -1) tail = tail.slice(0, linkIdx);
    tail = tail.replace(/cr:\s*\d{4}-\d{2}-\d{2}(\s+\d{2}:\d{2})?/gi, " ");
    tail = tail.replace(/\(?\s*\d*\s*d?\s*overdue\s*\)?/gi, " ");

    const dateMatches = [...tail.matchAll(/(\d{4}-\d{2}-\d{2})(?:\s+(\d{2}:\d{2}))?/g)];
    let startDate=null, targetDate=null, changedDate=null, changedTime=null;
    if (dateMatches.length >= 3){
      startDate = dateMatches[0][1]; targetDate = dateMatches[1][1];
      changedDate = dateMatches[2][1]; changedTime = dateMatches[2][2] || null;
    } else if (dateMatches.length === 2){
      startDate = dateMatches[0][1]; changedDate = dateMatches[1][1]; changedTime = dateMatches[1][2] || null;
    } else if (dateMatches.length === 1){
      startDate = dateMatches[0][1]; changedTime = dateMatches[0][2] || null;
    }
    let effort = null;
    if (dateMatches.length){
      const last = dateMatches[dateMatches.length-1];
      const afterLast = tail.slice(last.index + last[0].length);
      const em = afterLast.match(/(\d{1,3})/);
      if (em) effort = parseInt(em[1],10);
    }

    let parent = parentCol;
    if (!parent){
      const parentMatch = win.match(/Parent\s*:?\s*#?(\d{3,6})/i);
      if (parentMatch) parent = parentMatch[1];
    }

    items.push({
      id: ticketIdCol || anchors[i].id, title, type, state: stateVal, assignee,
      start: startDate, target: targetDate, changed: changedDate, changedTime,
      effort, parent, generatedDate,
      completeness: [startDate,targetDate,changedDate,effort].filter(v=>v!==null&&v!==undefined).length,
      changedSortKey: changedDate ? changedDate + "T" + (changedTime || "00:00") : "0000-00-00T00:00"
    });
  }
  return { generatedDate, items };
}

/* ============================================================
   SHEET-ROWS PARSER — for spreadsheet exports with a real header row
   (verbatim port of parseWorkItemsFromSheetRows)
   ============================================================ */
function parseWorkItemsFromSheetRows(rows){
  const items = [];
  let generatedDate = null;
  for (const row of rows){
    const joined = row.map(c => (c==null?"":String(c))).join(" ");
    const g = joined.match(/Generated\s+(\d{4}-\d{2}-\d{2})/i);
    if (g && !generatedDate) generatedDate = g[1];
  }
  const headerIdx = rows.findIndex(r => {
    const low = r.map(c => (c==null?"":String(c)).toLowerCase());
    return low.some(c=>c.includes("work item")||c.includes("title")) && low.some(c=>c.includes("type")) && low.some(c=>c.includes("state"));
  });
  if (headerIdx === -1) return { generatedDate, items: null };
  const header = rows[headerIdx].map(c => (c==null?"":String(c)).toLowerCase().trim());
  const colIdx = (names) => header.findIndex(h => names.some(n => h.includes(n)));
  const iWork = colIdx(["work item","title"]);
  const iType = colIdx(["type"]);
  const iState = colIdx(["state","status"]);
  const iAssignee = colIdx(["assigned"]);
  const iStart = colIdx(["start"]);
  const iTarget = colIdx(["target"]);
  const iChanged = colIdx(["changed"]);
  const iEffort = colIdx(["effort"]);
  const iParent = colIdx(["parent"]);
  for (let r = headerIdx+1; r < rows.length; r++){
    const row = rows[r];
    if (!row || !row.length) continue;
    const workRaw = iWork>=0 ? String(row[iWork]||"") : "";
    if (!workRaw.trim()) continue;
    const idm = workRaw.match(/#?(\d{3,6})/);
    const title = workRaw.replace(/^#?\d{3,6}:?\s*/, "").trim() || workRaw.trim();
    const effortRaw = iEffort>=0 ? row[iEffort] : null;
    const parentRaw = iParent>=0 ? String(row[iParent]||"").trim() : "";
    const parentM = parentRaw.match(/(\d{3,6})/);
    items.push({
      id: idm ? idm[1] : "r" + r,
      title, type: iType>=0 ? String(row[iType]||"").trim() : "",
      state: iState>=0 ? String(row[iState]||"").trim() : "",
      assignee: iAssignee>=0 ? String(row[iAssignee]||"").trim() : "",
      start: iStart>=0 ? String(row[iStart]||"").trim() || null : null,
      target: iTarget>=0 ? String(row[iTarget]||"").trim() || null : null,
      changed: iChanged>=0 ? String(row[iChanged]||"").trim() || null : null,
      changedTime: null,
      effort: effortRaw && !isNaN(parseInt(effortRaw,10)) ? parseInt(effortRaw,10) : null,
      parent: parentM ? parentM[1] : null,
      generatedDate,
      completeness: 4,
      changedSortKey: (iChanged>=0 ? String(row[iChanged]||"") : "") || "0000-00-00T00:00"
    });
  }
  return { generatedDate, items };
}

/* ============================================================
   HTML PARSER — for raw Webyith "DevOps Report" .html exports
   (verbatim port of parseWorkItemsFromHtml)
   ============================================================ */
function parseWorkItemsFromHtml(rawHtml){
  const parser = new DOMParser();
  const doc = parser.parseFromString(rawHtml || "", "text/html");

  const bodyText = doc.body ? doc.body.textContent : "";
  const genMatch = bodyText.match(/Generated\s+(\d{4}-\d{2}-\d{2})/);
  const generatedDate = genMatch ? genMatch[1] : null;

  const items = [];
  const tables = Array.from(doc.querySelectorAll("table"));
  for (const table of tables){
    const rows = Array.from(table.querySelectorAll("tr"));
    if (!rows.length) continue;
    const ths = Array.from(rows[0].querySelectorAll("th"));
    if (!ths.length || ths[0].textContent.trim() !== "Work Item") continue;

    const headers = ths.map(th => th.textContent.trim());
    const ncols = headers.length;
    const iType = headers.indexOf("Type");
    const iState = headers.indexOf("State");
    const iAssignee = headers.indexOf("Assigned To");
    const iStart = headers.indexOf("Start");
    const iTarget = headers.indexOf("Target");
    const iChanged = headers.indexOf("Changed");
    const iEffort = headers.indexOf("Effort");

    const levelStack = {};
    for (let i = 1; i < rows.length; i++){
      const tds = Array.from(rows[i].querySelectorAll("td"));
      if (tds.length !== ncols) continue;

      const firstTd = tds[0];
      const rawText = firstTd.textContent || "";
      const idMatch = rawText.match(/#(\d{3,6})/);
      if (!idMatch) continue;
      const id = idMatch[1];

      const titleSpan = firstTd.querySelector("span");
      let title = (titleSpan ? titleSpan.textContent : rawText).trim();
      title = title.replace(/^[^\w#]*#\d{3,6}:\s*/, "").replace(/\s+/g, " ");
      if (!title) continue;

      const style = firstTd.getAttribute("style") || "";
      const pm = style.match(/padding-left:\s*(\d+)px/);
      const level = pm ? Math.round((parseInt(pm[1], 10) - 10) / 22) : 0;
      const parent = level > 0 ? (levelStack[level - 1] || null) : null;
      levelStack[level] = id;
      Object.keys(levelStack).forEach(k => { if (parseInt(k, 10) > level) delete levelStack[k]; });

      const cellText = (idx) => idx >= 0 ? (tds[idx].textContent || "").trim() : "";
      const grabDate = (idx) => {
        const raw = cellText(idx);
        const dm = raw.match(/(\d{4}-\d{2}-\d{2})/);
        const tm = raw.match(/(\d{2}:\d{2})/);
        return { date: dm ? dm[1] : null, time: tm ? tm[1] : null };
      };
      const startD = grabDate(iStart);
      const targetD = grabDate(iTarget);
      const changedD = grabDate(iChanged);
      let effort = null;
      const effortRaw = cellText(iEffort);
      const em = effortRaw.match(/(\d+(?:\.\d+)?)/);
      if (em) effort = parseFloat(em[1]);

      items.push({
        id, title,
        type: cellText(iType), state: cellText(iState), assignee: cellText(iAssignee),
        start: startD.date, target: targetD.date, changed: changedD.date, changedTime: changedD.time,
        effort, parent, generatedDate,
        completeness: [startD.date, targetD.date, changedD.date, effort].filter(v => v !== null && v !== undefined).length,
        changedSortKey: changedD.date ? changedD.date + "T" + (changedD.time || "00:00") : "0000-00-00T00:00"
      });
    }
  }
  return { generatedDate, items };
}

/* ============================================================
   FILE READERS (Node equivalents of the browser extract* helpers)
   ============================================================ */
async function extractPdfText(filePath){
  const buf = fs.readFileSync(filePath);
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
  let full = "";
  for (let p = 1; p <= doc.numPages; p++){
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    full += " " + content.items.map(it => it.str).join(" ");
  }
  return full;
}
async function extractDocxText(filePath){
  const result = await mammoth.extractRawText({ path: filePath });
  return result.value || "";
}
function extractXlsxRows(filePath){
  const wb = XLSX.readFile(filePath);
  let rows = [];
  wb.SheetNames.forEach(name => {
    const sheet = wb.Sheets[name];
    const asRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    rows = rows.concat(asRows);
  });
  return rows;
}

async function parseFile(filePath){
  const ext = path.extname(filePath).toLowerCase().replace(".", "");
  const base = path.basename(filePath);
  const rec = { file: base, ext, generatedDate: null, items: [], status: "pending", warning: null };
  try {
    if (ext === "pdf"){
      const text = await extractPdfText(filePath);
      const { generatedDate, items } = parseWorkItemsFromText(text);
      rec.generatedDate = generatedDate; rec.items = items;
      rec.status = items.length ? "ok" : "warn";
      rec.warning = items.length ? null
        : text.trim().length < 20
          ? "No selectable text found in this PDF — it may be a scanned image without a text layer, or password-protected."
          : "Text was extracted (" + text.trim().length + " chars), but no #ID: work item rows matched the expected report layout.";
    } else if (ext === "docx" || ext === "doc"){
      const text = await extractDocxText(filePath);
      const { generatedDate, items } = parseWorkItemsFromText(text);
      rec.generatedDate = generatedDate; rec.items = items;
      rec.status = items.length ? "ok" : "warn";
      rec.warning = items.length ? null
        : text.trim().length < 20
          ? "No text content found in this document."
          : "Text was extracted (" + text.trim().length + " chars), but no #ID: work item rows matched the expected report layout.";
    } else if (ext === "html" || ext === "htm"){
      const rawHtml = fs.readFileSync(filePath, "utf8");
      const { generatedDate, items } = parseWorkItemsFromHtml(rawHtml);
      rec.generatedDate = generatedDate; rec.items = items;
      rec.status = items.length ? "ok" : "warn";
      rec.warning = items.length ? null : "No 'Work Item' tables found in this HTML file — it may not be a Webyith daily report export.";
    } else if (ext === "xlsx" || ext === "xls"){
      const rows = extractXlsxRows(filePath);
      let res = parseWorkItemsFromSheetRows(rows);
      if (res.items === null){
        const flat = rows.map(r => r.join(" ")).join(" ");
        res = parseWorkItemsFromText(flat);
      }
      rec.generatedDate = res.generatedDate; rec.items = res.items || [];
      rec.status = rec.items.length ? "ok" : "warn";
      rec.warning = rec.items.length ? null : "No recognizable work item table found in this spreadsheet.";
    } else {
      rec.status = "skip";
    }
  } catch (err){
    rec.status = "warn";
    rec.warning = "Could not read this file (" + (err && err.message ? err.message : "unknown error") + ").";
  }
  // Fall back to the date encoded in the filename (Webyith-Daily-Report_YYYY_MM_DD) whenever the
  // file's own content didn't yield a "Generated <date>" line — keeps date-range logic working
  // even for a file whose content parsing only partially succeeded.
  if (!rec.generatedDate) rec.generatedDate = dateFromFilename(base);
  return rec;
}

/* ============================================================
   AGGREGATION (verbatim port of mergeAllItems / computeStats / computeDeveloperActivity)
   ============================================================ */
function mergeAllItems(files){
  const byId = new Map();
  files.forEach(f => {
    (f.items || []).forEach(it => {
      const existing = byId.get(it.id);
      if (!existing) { byId.set(it.id, it); return; }
      if (it.changedSortKey > existing.changedSortKey) byId.set(it.id, it);
      else if (it.changedSortKey === existing.changedSortKey && it.completeness >= existing.completeness) byId.set(it.id, it);
    });
  });
  return [...byId.values()];
}

function computeStats(items, refDate){
  const total = items.length;
  const open = items.filter(i => OPEN_STATES.includes(i.state)).length;
  const closed = items.filter(i => CLOSED_STATES.includes(i.state)).length;
  // Closed and Resolved mean different things operationally (confirmed with the business owner):
  // "Closed" = deployed to production (truly done); "Resolved" = passed dev, sitting on staging,
  // still awaiting deploy approval. Split out separately so the report can distinguish "actually
  // shipped" from "done but not yet shipped" instead of blending them into one Completion Rate.
  const deployedCount = items.filter(i => i.state === "Closed").length;
  const pendingDeploymentCount = items.filter(i => i.state === "Resolved").length;
  // Breakdown of the still-open bucket (New/Active/In Review) into "not started" vs "in progress" —
  // used together with pendingDeploymentCount to show the 3 sub-parts of "Carried Over to Next
  // Sprint" (a ticket that's Resolved may still get deployed mid-sprint, so it's carried-over too,
  // just further along than one that's merely Active or brand new).
  const notStartedCount = items.filter(i => i.state === "New").length;
  const inProgressCount = items.filter(i => i.state === "Active" || i.state === "In Review").length;
  const openBugs = items.filter(i => OPEN_STATES.includes(i.state) && i.type === "Bug").length;
  const totalEffort = items.reduce((s,i)=> s + (typeof i.effort === "number" ? i.effort : 0), 0);
  const overdue = items.filter(i => !CLOSED_STATES.includes(i.state) && i.target && refDate && i.target < refDate);
  const byType = {};
  items.forEach(i => { const t = i.type || "Unknown"; byType[t] = (byType[t]||0)+1; });

  const featuresCompleted = items.filter(i => CLOSED_STATES.includes(i.state) && (i.type === "Feature" || i.type === "User Story")).length;
  const techDebtCompleted = items.filter(i => CLOSED_STATES.includes(i.state) && i.type === "Issue").length;
  const bugsFixed = items.filter(i => CLOSED_STATES.includes(i.state) && i.type === "Bug").length;
  // Split each "Completed" count by its actual status (Closed = deployed to production, Resolved =
  // done on staging, awaiting deploy approval) so the report can show both instead of one blended
  // number — mirrors the Deployed/Pending Deployment split above, applied to these 3 metrics too.
  const featuresCompletedClosed = items.filter(i => i.state === "Closed" && (i.type === "Feature" || i.type === "User Story")).length;
  const featuresCompletedResolved = items.filter(i => i.state === "Resolved" && (i.type === "Feature" || i.type === "User Story")).length;
  const techDebtCompletedClosed = items.filter(i => i.state === "Closed" && i.type === "Issue").length;
  const techDebtCompletedResolved = items.filter(i => i.state === "Resolved" && i.type === "Issue").length;
  const bugsFixedClosed = items.filter(i => i.state === "Closed" && i.type === "Bug").length;
  const bugsFixedResolved = items.filter(i => i.state === "Resolved" && i.type === "Bug").length;

  const hasParentData = items.some(i => i.parent);
  let bugRatioDenominator, bugRatioMode;
  let totalBugs = null, totalFeatures = null, totalTasks = null, totalUserStories = null, totalIssues = null;
  if (hasParentData){
    // Denominator counting rules (confirmed with the business owner):
    //   - Task can never be counted on its own (it's just a breakdown of its parent's work,
    //     regardless of whether it happens to have no parent recorded) -> always contributes 0.
    //   - User Story and Issue are always counted as their own unit, whether or not they're a
    //     child of a Feature (unlike Task, they represent independently trackable work).
    //   - Feature is counted UNLESS it has a direct (level-1) child that is a User Story or Issue —
    //     in that case the child(ren) already represent that work, so counting the Feature too would
    //     double-count it. A Feature whose only children are Tasks (or none) still counts itself.
    const childTypesByParent = {};
    items.forEach(i => {
      if (i.parent){
        if (!childTypesByParent[i.parent]) childTypesByParent[i.parent] = new Set();
        childTypesByParent[i.parent].add(i.type);
      }
    });
    const hasUserStoryOrIssueChild = id => { const s = childTypesByParent[id]; return !!s && (s.has("User Story") || s.has("Issue")); };

    totalBugs = items.filter(i => i.type === "Bug").length; // all bugs, any level — a bug is real defect work regardless of hierarchy
    totalFeatures = items.filter(i => i.type === "Feature" && !hasUserStoryOrIssueChild(i.id)).length;
    totalTasks = 0; // Task never counted independently
    totalUserStories = items.filter(i => i.type === "User Story").length; // counted regardless of parent
    totalIssues = items.filter(i => i.type === "Issue").length; // counted regardless of parent
    bugRatioDenominator = totalFeatures + totalTasks + totalUserStories + totalIssues;
    bugRatioMode = "structural";
  } else {
    totalBugs = openBugs;
    bugRatioDenominator = open;
    bugRatioMode = "simple";
  }
  const bugRatioNumerator = bugRatioMode === "structural" ? totalBugs : openBugs;
  const bugRatio = bugRatioDenominator > 0 ? Math.round((bugRatioNumerator/bugRatioDenominator)*100) : 0;

  return {
    total, open, closed, deployedCount, pendingDeploymentCount, notStartedCount, inProgressCount, bugs: openBugs, bugRatio, bugRatioMode, bugRatioDenominator,
    totalBugs, totalFeatures, totalTasks, totalUserStories, totalIssues,
    totalEffort, overdue, byType,
    featuresCompleted, techDebtCompleted, bugsFixed,
    featuresCompletedClosed, featuresCompletedResolved,
    techDebtCompletedClosed, techDebtCompletedResolved,
    bugsFixedClosed, bugsFixedResolved
  };
}

function computeDeveloperActivity(items, assignees){
  return assignees.map(name => {
    const mine = items.filter(i => i.assignee === name);
    return {
      name: name.replace(" - ONESECURE",""),
      active: mine.filter(i => OPEN_STATES.includes(i.state)).length,
      resolved: mine.filter(i => i.state === "Resolved").length,
      closed: mine.filter(i => i.state === "Closed").length,
      effort: mine.reduce((s,i)=> s + (typeof i.effort === "number" ? i.effort : 0), 0)
    };
  }).filter(d => d.active + d.resolved + d.closed > 0);
}

/* ============================================================
   SPRINT FOLDER NAMING CONVENTION: Sprint-<StartYYYYMMDD>_<EndYYYYMMDD>
   ============================================================ */
const SPRINT_FOLDER_RE = /^Sprint-(\d{8})_(\d{8})$/;

function ymdToIso(ymd){
  return ymd.slice(0,4) + "-" + ymd.slice(4,6) + "-" + ymd.slice(6,8);
}

// Returns { start, end } (ISO yyyy-mm-dd) if `folderName` matches the convention, else null.
function parseSprintFolderName(folderName){
  const m = SPRINT_FOLDER_RE.exec(folderName);
  if (!m) return null;
  return { start: ymdToIso(m[1]), end: ymdToIso(m[2]) };
}

function addDaysIso(iso, n){
  const d = new Date(iso + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

/* ============================================================
   DAILY REPORT FILENAME CONVENTION: Webyith-Daily-Report_<YYYY_MM_DD>
   The first daily file of a sprint is dated (sprint start + 1); the last is dated
   (next sprint's start - 3), which — since sprints run back-to-back (next start = this end + 1) —
   works out to (this sprint's end - 2). The last 2 calendar days of a sprint have no daily file.
   ============================================================ */
const DAILY_FILENAME_RE = /Webyith-Daily-Report_(\d{4})_(\d{2})_(\d{2})/;

function dateFromFilename(filename){
  const m = DAILY_FILENAME_RE.exec(filename);
  if (!m) return null;
  return m[1] + "-" + m[2] + "-" + m[3];
}

// Inclusive list of ISO dates a fully-up-to-date sprint folder should have one daily file for.
function expectedDailyDates(sprintStart, sprintEnd){
  const first = addDaysIso(sprintStart, 1);
  const last = addDaysIso(sprintEnd, -2);
  const dates = [];
  for (let d = first; d <= last; d = addDaysIso(d, 1)) dates.push(d);
  return dates;
}

function listSprintFolders(baseFolder){
  const todayIso = new Date().toISOString().slice(0, 10);
  const entries = fs.readdirSync(baseFolder, { withFileTypes: true })
    .filter(e => e.isDirectory())
    .map(e => e.name)
    .map(name => {
      const period = parseSprintFolderName(name);
      if (!period) return null;
      return {
        folder: name,
        start: period.start,
        end: period.end,
        isCurrent: todayIso >= period.start && todayIso <= period.end
      };
    })
    .filter(Boolean)
    .sort((a, b) => (a.start < b.start ? 1 : -1)); // newest first
  return { today: todayIso, sprints: entries };
}

/* ============================================================
   CLI
   ============================================================ */
const REPORT_EXTS = new Set(["pdf","docx","doc","html","htm","xlsx","xls"]);

// Does the actual parse+merge+stats work and returns the result object. Does NOT write any file —
// callers (the CLI `ingest` command, and `computeTrendSeries` below) decide what to do with it.
async function ingestFolderData(folder){
  // The sprint's own folder name (Sprint-<start>_<end>) is authoritative for the reporting PERIOD
  // when present — it doesn't depend on which daily files happen to be in the folder yet. Computed
  // up front so we can flag missing daily files even when the folder is empty or partly filled.
  const sprintPeriod = parseSprintFolderName(path.basename(path.resolve(folder)));
  const expectedDates = sprintPeriod ? expectedDailyDates(sprintPeriod.start, sprintPeriod.end) : null;

  const entries = fs.readdirSync(folder, { withFileTypes: true })
    .filter(e => e.isFile())
    .map(e => e.name)
    .filter(name => REPORT_EXTS.has(path.extname(name).toLowerCase().replace(".", "")))
    .sort();

  if (!entries.length){
    return {
      error: "No report files (.pdf/.docx/.html/.xlsx) found in " + folder,
      sprintPeriod: sprintPeriod ? { start: sprintPeriod.start, end: sprintPeriod.end, source: "folder-name" } : null,
      missingDailyReports: expectedDates || null
    };
  }

  const files = [];
  for (const name of entries){
    const rec = await parseFile(path.join(folder, name));
    files.push(rec);
  }

  const allItems = mergeAllItems(files);
  const generatedDates = files.map(f => f.generatedDate).filter(Boolean).sort();
  // Deliberately NOT used for the overdue "as of" reference date below: a sprint in progress has an
  // end date in the future, and comparing against it would wrongly mark still-on-track items as
  // overdue. For that, "as of" is always the latest daily report actually ingested (falling back to
  // today if the folder has no dated files yet).
  const todayIso = new Date().toISOString().slice(0, 10);
  const refDate = generatedDates.length ? generatedDates[generatedDates.length - 1] : todayIso;
  const stats = computeStats(allItems, refDate);
  const assignees = [...new Set(allItems.map(i => i.assignee).filter(Boolean))];
  const activity = computeDeveloperActivity(allItems, assignees);

  // Cross-check which expected daily report dates (per the Sprint-start_end + filename-date
  // convention) have no corresponding file in this folder yet. This is PURELY INFORMATIONAL — it
  // never blocks stats/report generation. Real sprints routinely miss the theoretical first day,
  // last day, and/or days in between (public holidays, no export that day, etc.); the Agent should
  // just compute everything from whichever daily files actually exist and, if useful, mention this
  // list as a "coverage" footnote rather than treating it as an error.
  const presentDates = new Set(generatedDates);
  const missingDailyReports = expectedDates ? expectedDates.filter(d => !presentDates.has(d)) : null;

  return {
    folder,
    sprintPeriod: sprintPeriod ? { start: sprintPeriod.start, end: sprintPeriod.end, source: "folder-name" } : null,
    filesProcessed: files.map(f => ({ file: f.file, status: f.status, warning: f.warning, generatedDate: f.generatedDate, itemCount: f.items.length })),
    generatedDateRange: generatedDates.length ? { min: generatedDates[0], max: generatedDates[generatedDates.length-1] } : null,
    missingDailyReports,
    mergedItemCount: allItems.length,
    stats,
    activity,
    items: allItems
  };
}

async function ingestFolder(folder, outPath){
  const result = await ingestFolderData(folder);
  fs.writeFileSync(outPath, JSON.stringify(result, null, 2));
  if (result.error){
    process.stderr.write("No report files found. Wrote error to " + outPath + "\n");
    return;
  }
  // Short human-readable summary on stdout — the full data lives in outPath for the Agent to Read.
  const missingDailyReports = result.missingDailyReports;
  const gapNote = missingDailyReports && missingDailyReports.length
    ? " (FYI, no daily file for: " + missingDailyReports.join(", ") + " — could be holidays/weekends; stats above are computed from the files present.)"
    : "";
  process.stdout.write(
    "Ingested " + result.filesProcessed.length + " file(s) from " + folder + " -> " + result.mergedItemCount + " unique work items." + gapNote + "\n" +
    "Wrote full result to: " + outPath + "\n"
  );
}

/* ============================================================
   TREND ACROSS SPRINTS — reuses each sprint's own _sprint_data.json when present (ingesting fresh
   if not), and reduces each to the handful of headline numbers used for sprint-over-sprint and
   multi-sprint trend reporting. Always computes the FULL chronological history (cheap: it just
   re-reads each sprint's already-ingested _sprint_data.json) — capping to "last N sprints" for
   display happens separately in the CLI output below, so the "vs previous sprint" delta can always
   find the true previous sprint even once history grows past the display window.
   ============================================================ */
async function computeTrendSeries(baseFolder){
  const { sprints } = listSprintFolders(baseFolder); // newest-first
  const chronological = sprints.slice().sort((a, b) => (a.start < b.start ? -1 : 1));

  const series = [];
  for (const s of chronological){
    const folderPath = path.join(baseFolder, s.folder);
    const dataPath = path.join(folderPath, "_sprint_data.json");
    let data;
    if (fs.existsSync(dataPath)){
      data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
    } else {
      data = await ingestFolderData(folderPath);
    }
    if (data.error) continue; // empty sprint folder — nothing to plot yet
    const stats = data.stats;
    const totalCompleted = stats.featuresCompleted + stats.techDebtCompleted + stats.bugsFixed;
    series.push({
      folder: s.folder,
      start: s.start,
      end: s.end,
      isCurrent: s.isCurrent,
      total: stats.total,
      closed: stats.closed,
      // Deployed Rate = Closed only / total (truly shipped to production) — replaces the old
      // "completionRate" which blended Closed+Resolved together. Pending Deployment / Carried Over
      // are tracked alongside it so multi-sprint trend data has them available if needed later.
      deployedRate: stats.total > 0 ? Math.round((stats.deployedCount / stats.total) * 100) : 0,
      pendingDeploymentRate: stats.total > 0 ? Math.round((stats.pendingDeploymentCount / stats.total) * 100) : 0,
      // Carried Over to Next Sprint = everything not yet Closed (Resolved + still-open) — a Resolved
      // ticket may still get deployed mid-sprint, so it counts as carried over too, just further
      // along than one that's merely Active/In Review/New. Replaces the old open-only definition.
      carriedOverRate: stats.total > 0 ? Math.round(((stats.pendingDeploymentCount + stats.open) / stats.total) * 100) : 0,
      bugRatio: stats.bugRatio,
      overdueCount: stats.overdue.length,
      totalCompleted,
      featuresCompleted: stats.featuresCompleted,
      techDebtCompleted: stats.techDebtCompleted,
      bugsFixed: stats.bugsFixed
    });
  }
  return { today: new Date().toISOString().slice(0, 10), series };
}

async function main(){
  const [cmd, arg, ...rest] = process.argv.slice(2);
  if (cmd === "ingest" && arg){
    const outFlagIdx = rest.indexOf("--out");
    const outPath = outFlagIdx !== -1 && rest[outFlagIdx + 1] ? rest[outFlagIdx + 1] : path.join(arg, "_sprint_data.json");
    await ingestFolder(arg, outPath);
  } else if (cmd === "list-sprints" && arg){
    const result = listSprintFolders(arg);
    process.stdout.write(JSON.stringify(result, null, 2) + "\n");
  } else if (cmd === "trend" && arg){
    const outFlagIdx = rest.indexOf("--out");
    const outPath = outFlagIdx !== -1 && rest[outFlagIdx + 1] ? rest[outFlagIdx + 1] : path.join(arg, "_sprint_trend.json");
    const limitFlagIdx = rest.indexOf("--limit");
    // Default: last 6 sprints for the trend chart/table (keeps it readable as history grows).
    // `series` below is this capped, display-ready list; `fullSeries` keeps every sprint on record
    // so "vs previous sprint" delta lookups stay correct even once history exceeds the display cap.
    const limit = limitFlagIdx !== -1 && rest[limitFlagIdx + 1] ? parseInt(rest[limitFlagIdx + 1], 10) : 6;
    const full = await computeTrendSeries(arg);
    const capped = limit > 0 && full.series.length > limit ? full.series.slice(full.series.length - limit) : full.series;
    const result = { today: full.today, limit, series: capped, fullSeries: full.series };
    fs.writeFileSync(outPath, JSON.stringify(result, null, 2));
    process.stdout.write("Computed trend across " + full.series.length + " sprint(s) on record, showing last " + capped.length + ". Wrote " + outPath + "\n");
  } else if (cmd === "parse-one" && arg){
    const rec = await parseFile(arg);
    process.stdout.write(JSON.stringify(rec, null, 2) + "\n");
  } else {
    console.error("Usage:\n  node sprint_engine.js ingest <folder> [--out <file.json>]\n  node sprint_engine.js list-sprints <baseFolder>\n  node sprint_engine.js trend <baseFolder> [--out <file.json>] [--limit N]\n  node sprint_engine.js parse-one <file>");
    process.exit(1);
  }
}

main().catch(err => { console.error("ERROR:", err); process.exit(1); });
