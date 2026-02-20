/**
 * Helpers.gs
 * - No CONFIG declaration here.
 * - Single shared header cache in this file ONLY.
 */

var __HEADERS_CACHE = {}; // global project scope cache (declare once)

//==============================================================
// LOGGING (Execution log JSON)
//==============================================================
function createLogCtx_(runId, scope, meta) {
  return { runId: runId || "", scope: scope || "", meta: meta || {}, t0: Date.now() };
}

function childCtx_(ctx, childScope) {
  const p = (ctx && ctx.scope) ? ctx.scope : "";
  return {
    runId: (ctx && ctx.runId) ? ctx.runId : "",
    scope: p ? `${p} > ${childScope}` : childScope,
    meta: (ctx && ctx.meta) ? ctx.meta : {},
    t0: Date.now()
  };
}

function _log_(level, ctx, msg, data) {
  const o = {
    ts: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss.SSS"),
    level: String(level || "INFO"),
    runId: (ctx && ctx.runId) ? ctx.runId : "",
    scope: (ctx && ctx.scope) ? ctx.scope : "",
    msg: msg || "",
    meta: (ctx && ctx.meta) ? ctx.meta : {},
    data: data || {}
  };
  const line = JSON.stringify(o);
  Logger.log(line);
  try { console.log(line); } catch (_) {}
}

function logInfo_(ctx, msg, data) { _log_("INFO", ctx, msg, data); }
function logDebug_(ctx, msg, data) { _log_("DEBUG", ctx, msg, data); }
function logWarn_(ctx, msg, data) { _log_("WARN", ctx, msg, data); }
function logError_(ctx, msg, data) { _log_("ERROR", ctx, msg, data); }

//==============================================================
// BASIC UTILITIES
//==============================================================
function getSheetOrAlert_(ss, sheetName) {
  const s = ss.getSheetByName(sheetName);
  if (!s) SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
  return s;
}

function normalizeString_(s) {
  return String(s === null || s === undefined ? "" : s).trim();
}

function normalizeJobKey_(s) {
  return normalizeString_(s).replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeNotes_(s) {
  return normalizeString_(s)
    .replace(/\u00A0/g, " ")
    .replace(/\s*;\s*/g, "; ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 864e5));
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function dateToStr_(v) {
  const d = parseDate_(v);
  if (!d) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "M/dd/yyyy");
}

function colAny_(h, names) {
  for (let i = 0; i < names.length; i++) if (h[names[i]] !== undefined) return h[names[i]];
  return undefined;
}

function normalizeJobKeyForCompare_(key) {
  const k = normalizeJobKey_(key);
  return k.endsWith(" 0000") ? k.replace(" 0000", "") : k;
}

function parseNotesParts_(noteStr) {
  const note = normalizeString_(noteStr);

  const endMatch = note.match(/End Date=([^;]+)/i);
  const endDate = endMatch ? normalizeString_(endMatch[1]) : "";

  // UPDATED: Capture ALL P-entries (concatenated)
  const pMatches = note.match(/P-[0-9]{1,2}\/[0-9]{1,2}[^;]*/gi) || [];
  const pFull = pMatches.map(s => s.trim()).join("; ");

  const cspMatch = note.match(/CSP[^;]*/i);
  const csp = cspMatch ? normalizeString_(cspMatch[0]) : "";

  const custom = note
    .replace(/End Date=[^;]+;?/gi, "")
    .replace(/P-[^;]+;?/gi, "") // Strips all P- entries
    .replace(/CSP[^;]*;?/gi, "")
    .split(";")
    .map(s => s.trim())
    .filter(Boolean)
    .join("; ");

  return { endDate, pFull, csp, custom };
}

//==============================================================
// HEADERS (cache first occurrence)
//==============================================================
function getHeaders_(sheet) {
  const key = sheet.getSheetId() + ":" + sheet.getLastColumn();
  if (__HEADERS_CACHE[key]) return __HEADERS_CACHE[key];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};

  headers.forEach((cell, i) => {
    const clean = normalizeString_(cell);
    if (!clean) return;
    if (map[clean] === undefined) map[clean] = i;
    // Convenience alias
    if (clean.toLowerCase() === "item" && map["Item No."] === undefined) map["Item No."] = i;
  });

  __HEADERS_CACHE[key] = map;
  return map;
}

function getSuffixKey_(h) {
  if (h["Suffix"] !== undefined) return "Suffix";
  if (h["Job Suffix"] !== undefined) return "Job Suffix";
  return null;
}

function getCompositeJobKey_(j, s, sub) {
  if (!j) return "";
  const job = normalizeJobKey_(j);
  const suf = parseInt(s, 10) || 0;
  return suf !== 0 ? `${job} ${String(suf).padStart(4, "0")}` : (sub ? `${job} 0000` : job);
}

//==============================================================
// EXTERNAL LOG (Update Change Log)
//==============================================================
function ensureUpdateLogHeader_(sheet) {
  const headers = [
    "Timestamp", "Report Sheet", "Project Coordinator", "Job Order",
    "Customer/Action", "PO", "Qty", "Item No.", "Change Details",
    "Severity", "Run ID"
  ];
  const current = sheet.getRange(1, 1, 1, Math.max(headers.length, sheet.getLastColumn())).getValues()[0];
  const needs = headers.some((h, i) => normalizeString_(current[i]) !== h);

  if (sheet.getLastRow() === 0 || needs) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
}

function writeToExternalLog_(logs, targetSheetName, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "writeToExternalLog_", {}), "writeToExternalLog_");
  try {
    if (!logs || !logs.length) {
      logDebug_(ctx, "No logs to write");
      return;
    }

    const ss = SpreadsheetApp.openById(CONFIG.LOG_SHEET_ID);
    const tabName = targetSheetName || "Update Change Log";
    const sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);

    ensureUpdateLogHeader_(sheet);

    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    const values = logs.map(l => [
      ts,
      l.reportSheet || "",
      l.projectCoordinator || "",
      l.jobOrder || "",
      l.customer || "",
      l.po || "",
      l.qty || "",
      l.itemNo || "",
      (l.changes || []).join("\n"),
      l.severity || "INFO",
      l.runId || ""
    ]);

    sheet.insertRowsAfter(1, values.length);
    const r = sheet.getRange(2, 1, values.length, values[0].length);
    
    // Protect key columns from auto-parsing as dates before setting values
    sheet.getRange(2, 4, values.length, 1).setNumberFormat("@"); // Job Order
    sheet.getRange(2, 6, values.length, 1).setNumberFormat("@"); // PO
    sheet.getRange(2, 8, values.length, 1).setNumberFormat("@"); // Item No.
    
    r.setValues(values);
    r.setWrap(true);
    r.setVerticalAlignment("top");

    logInfo_(ctx, "External log write OK", { tabName, rowsWritten: values.length });
  } catch (e) {
    logError_(ctx, "External log write FAILED", { message: e.message || String(e), stack: e.stack || "" });
  }
}

function logToolAction_(action, details, sheetName, targetSheetName, severity, runId, parentCtx) {
  const userKey = Session.getTemporaryActiveUserKey() || "Anonymous";
  const entry = {
    reportSheet: sheetName,
    projectCoordinator: resolveUserName_(userKey),
    jobOrder: "TOOL ACTION",
    customer: action,
    po: "N/A",
    qty: "N/A",
    itemNo: "N/A",
    severity: severity || "INFO",
    runId: runId || "",
    changes: [details]
  };
  writeToExternalLog_([entry], targetSheetName || "Update Change Log", parentCtx);
}

//==============================================================
// WRITE-ONLY-CHANGED (contiguous grouping)
//==============================================================
function applyColumnUpdates_(sheet, col1Based, rowToValueMap) {
  const rows = Array.from(rowToValueMap.keys()).sort((a, b) => a - b);
  if (rows.length === 0) return 0;

  let start = rows[0];
  let prev = rows[0];

  const flush = (s, e) => {
    const num = e - s + 1;
    const values = [];
    for (let r = s; r <= e; r++) values.push([rowToValueMap.get(r)]);
    sheet.getRange(s, col1Based, num, 1).setValues(values);
  };

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r === prev + 1) { prev = r; continue; }
    flush(start, prev);
    start = prev = r;
  }
  flush(start, prev);

  return rows.length;
}

//==============================================================
// TRACKING SHEET UPDATE (due date + notes + PC fill)
//==============================================================
function processSingleReportSheet_(sheet, sourceData, shortageData, cspData, runId, summary, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(runId, `processSingleReportSheet(${sheet.getName()})`, {}), `processSingleReportSheet(${sheet.getName()})`);

  const logs = [];
  const h = getHeaders_(sheet);

  const missing = [];
  if (h["Job Order"] === undefined) missing.push("Job Order");
  if (h["MTL Due Date"] === undefined) missing.push("MTL Due Date");
  if (h["End Date Notes"] === undefined) missing.push("End Date Notes");
  if (missing.length) {
    logWarn_(ctx, "Missing columns; no updates", { missing });
    return { logs };
  }

  const lastRow = getActualLastRow_(sheet, h["Job Order"] + 1);
  if (lastRow <= 1) return { logs };

  const raw = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const disp = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();

  const dueCol = h["MTL Due Date"] + 1;
  const noteCol = h["End Date Notes"] + 1;

  const hasPcCol = (h["Project Coordinator"] !== undefined);
  const pcCol = hasPcCol ? (h["Project Coordinator"] + 1) : null;

  const dueUpdates = new Map();
  const noteUpdates = new Map();
  const pcUpdates = new Map();

  let matchCount = 0;
  let unmatchedCount = 0;

  const sheetName = sheet.getName();
  const sampleLimit = CONFIG.UNMATCHED_SAMPLE_LIMIT_PER_SHEET || 10;

  for (let i = 0; i < disp.length; i++) {
    const rowNum = i + 2;
    const jobKey = normalizeJobKey_(disp[i][h["Job Order"]]);

    let dataKey = sourceData.jobsInSource.has(jobKey) ? jobKey : null;
    if (!dataKey && jobKey.includes(" 0000")) {
      const base = jobKey.replace(" 0000", "");
      if (sourceData.jobsInSource.has(base)) dataKey = base;
    }

    if (!dataKey) {
      unmatchedCount++;
      if (summary && summary.unmatchedSamples && (summary.unmatchedSamples[sheetName] || []).length < sampleLimit) {
        summary.unmatchedSamples[sheetName].push(jobKey || `(row ${rowNum})`);
      }
      continue;
    }
    matchCount++;

    // Current values
    const oldDue = parseDate_(raw[i][h["MTL Due Date"]]);
    const oldDueStr = dateToStr_(oldDue);

    const oldNote = normalizeString_(disp[i][h["End Date Notes"]]);
    const oldNoteNorm = normalizeNotes_(oldNote);

    const oldPc = hasPcCol ? normalizeString_(disp[i][h["Project Coordinator"]]) : "";

    // Source values
    const srcDue = parseDate_(sourceData.dateMap.get(dataKey));
    const srcDueStr = dateToStr_(srcDue);

    const srcEnd = parseDate_(sourceData.endMap.get(dataKey));
    const srcEndStr = dateToStr_(srcEnd);

    const srcItem = normalizeString_(sourceData.itemMap.get(dataKey) || "");
    const srcCust = normalizeString_(sourceData.customerMap.get(dataKey) || "");
    const srcStatus = normalizeString_(sourceData.statusMap.get(dataKey) || "");
    const srcAssignedTo = normalizeString_(sourceData.assignedToMap.get(dataKey) || "");

    // Rebuild notes (preserve manual)
    const endStr = srcEnd ? `End Date=${Utilities.formatDate(srcEnd, Session.getScriptTimeZone(), "M/dd/yy")}` : "";
    
    // Handle array of shortages and join
    const shortageList = shortageData.get(dataKey) || [];
    // Sort array by Date ascending so P-dates are in order
    shortageList.sort((a, b) => {
        const dA = parseDate_(a.date) || 0;
        const dB = parseDate_(b.date) || 0;
        return dA - dB;
    });

    const pStr = shortageList.map(s => 
      `P-${Utilities.formatDate(s.date, Session.getScriptTimeZone(), "M/dd")} (${normalizeString_(s.item)})`
    ).join("; ");

    const cspStr = cspData.get(dataKey) || "";

    const cleanCustom = oldNote
      .replace(/End Date=[^;]+/g, "")
      .replace(/P-[^;]*/g, "")
      .replace(/CSP[^;]*/g, "")
      .split(";")
      .map(s => s.trim())
      .filter(Boolean);

    const newNote = [endStr, pStr, cspStr, ...cleanCustom].filter(Boolean).join("; ");
    const newNoteNorm = normalizeNotes_(newNote);

    let dueChanged = false;
    let notesChanged = false;
    let pcChanged = false;

    if (srcDue) {
      const oldTime = oldDue ? oldDue.getTime() : null;
      if (!oldTime || srcDue.getTime() !== oldTime) {
        dueUpdates.set(rowNum, srcDue);
        dueChanged = true;
      }
    }

    if (oldNoteNorm !== newNoteNorm) {
      noteUpdates.set(rowNum, newNote);
      notesChanged = true;
    }

    // Fill Project Coordinator only if blank
    if (hasPcCol && oldPc === "" && srcAssignedTo !== "") {
      pcUpdates.set(rowNum, srcAssignedTo);
      pcChanged = true;
    }

    // Log changes
    if (dueChanged || notesChanged || pcChanged) {
      const headlinePieces = [];
      if (srcCust) headlinePieces.push(srcCust);
      if (srcItem) headlinePieces.push(`Item ${srcItem}`);
      if (srcStatus) headlinePieces.push(`Status ${srcStatus}`);

      const lines = [];
      lines.push(headlinePieces.length
        ? `Run ${runId}: ${jobKey} (${headlinePieces.join(" • ")})`
        : `Run ${runId}: ${jobKey}`);

      if (dueChanged) lines.push(`Due date: ${oldDueStr || "blank"} → ${srcDueStr || "blank"}`);
      if (pcChanged) lines.push(`Project Coordinator: blank → ${srcAssignedTo}`);

      if (notesChanged) {
        const oldParts = parseNotesParts_(oldNote);
        const newParts = parseNotesParts_(newNote);

        const deltas = [];
        if (oldParts.endDate !== newParts.endDate) deltas.push(`End Date ${oldParts.endDate || "blank"} → ${newParts.endDate || "blank"}`);

        const oldP = oldParts.pFull || "none";
        const newP = newParts.pFull || "none";
        if (oldP !== newP) deltas.push(`Shortage ${oldP} → ${newP}`);

        if (oldParts.csp !== newParts.csp) deltas.push(`CSP ${oldParts.csp || "blank"} → ${newParts.csp || "blank"}`);
        if (oldParts.custom !== newParts.custom) deltas.push(`Manual notes changed`);

        lines.push(deltas.length ? `Notes updated: ${deltas.join(" | ")}` : "Notes updated");
      }

      lines.push(`SyteLine: Due ${srcDueStr || "-"}, End ${srcEndStr || "-"}`);
      lines.push(`Row ${rowNum} • Match ${dataKey}`);

      logs.push({
        reportSheet: sheetName,
        jobOrder: jobKey,
        projectCoordinator: (pcChanged ? srcAssignedTo : (disp[i][h["Project Coordinator"]] || "")),
        customer: disp[i][h["Customer"]] || srcCust || "",
        po: disp[i][h["PO"]] || "",
        qty: disp[i][h["Qty"]] || "",
        itemNo: disp[i][h["Item No."]] || srcItem || "",
        severity: "INFO",
        runId,
        changes: lines
      });
    }
  }

  // Apply write-only-changed updates
  const dueWritten = applyColumnUpdates_(sheet, dueCol, dueUpdates);
  const noteWritten = applyColumnUpdates_(sheet, noteCol, noteUpdates);
  const pcWritten = (hasPcCol ? applyColumnUpdates_(sheet, pcCol, pcUpdates) : 0);

  if (summary) {
    summary.dueDateChanges += dueWritten;
    summary.noteChanges += noteWritten;
    summary.pcFilledChanges += pcWritten;
    summary.unmatchedBySheet[sheetName] = unmatchedCount;
  }

  if (summary && summary.unmatchedSamples && (summary.unmatchedSamples[sheetName] || []).length) {
    logWarn_(ctx, "UNMATCHED SAMPLE", {
      sheet: sheetName,
      unmatchedCount,
      samples: summary.unmatchedSamples[sheetName].slice(0, sampleLimit)
    });
  }

  logInfo_(ctx, "FINISH", {
    matches: `${matchCount}/${disp.length}`,
    dueWritten,
    noteWritten,
    pcWritten,
    unmatched: unmatchedCount,
    logs: logs.length
  });

  return { logs };
}

//==============================================================
// SOURCE DATA LOADERS
//==============================================================
function loadSourceJobData_(sheet, splitSet, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "loadSourceJobData_", {}), "loadSourceJobData_");

  const h = getHeaders_(sheet);
  const data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();

  const sufKey = getSuffixKey_(h) || "Job Suffix";
  const itemCol = colAny_(h, ["Item", "Item No.", "Item No", "Item Number"]);
  const custCol = colAny_(h, ["Customer", "Customer Name", "Cust Num", "Cust", "CustomerNum"]);
  const statusCol = colAny_(h, ["Status", "Job Status", "Stat"]);
  const custPoCol = colAny_(h, ["Cust PO", "Customer PO", "Customer PO#", "CustomerPO", "CustPO"]);
  const assignedToCol = colAny_(h, ["Assigned To", "AssignedTo", "Assigned"]);

  const map = {
    jobsInSource: new Set(),
    dateMap: new Map(),
    endMap: new Map(),
    itemMap: new Map(),
    customerMap: new Map(),
    statusMap: new Map(),
    custPoMap: new Map(),
    assignedToMap: new Map()
  };

  data.forEach(r => {
    const job = r[h["Job"]];
    const suf = r[h[sufKey]];
    const key = normalizeJobKey_(getCompositeJobKey_(job, suf, splitSet.has(normalizeJobKey_(job))));
    if (!key) return;

    map.jobsInSource.add(key);
    map.dateMap.set(key, r[h["Due Date"]]);
    map.endMap.set(key, r[h["End Date"]]);
    map.itemMap.set(key, (itemCol !== undefined) ? r[itemCol] : "");
    map.customerMap.set(key, (custCol !== undefined) ? r[custCol] : "");
    map.statusMap.set(key, (statusCol !== undefined) ? r[statusCol] : "");
    map.custPoMap.set(key, (custPoCol !== undefined) ? r[custPoCol] : "");
    map.assignedToMap.set(key, (assignedToCol !== undefined) ? r[assignedToCol] : "");
  });

  logInfo_(ctx, "DONE", {
    rows: data.length,
    jobs: map.jobsInSource.size,
    sufKey,
    hasAssignedTo: assignedToCol !== undefined,
    hasCustPO: custPoCol !== undefined
  });

  return map;
}

function scanForSubassemblies_(sheet, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "scanForSubassemblies_", {}), "scanForSubassemblies_");
  const h = getHeaders_(sheet);
  const set = new Set();

  const sufKey = getSuffixKey_(h);
  if (!sufKey || h["Job"] === undefined) return set;

  const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (parseInt(r[h[sufKey]], 10) > 0) set.add(normalizeJobKey_(r[h["Job"]]));
  }

  logInfo_(ctx, "DONE", { subJobs: set.size, sufKey });
  return set;
}

function loadProducedItemsSet_(sheet, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "loadProducedItemsSet_", {}), "loadProducedItemsSet_");
  const h = getHeaders_(sheet);
  const set = new Set();

  const itemCol = colAny_(h, ["Item", "Item No.", "Item No", "Item Number"]);
  if (itemCol === undefined) return set;

  const vals = sheet.getRange(2, itemCol + 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues();
  vals.forEach(r => { if (r[0]) set.add(normalizeString_(r[0])); });

  logInfo_(ctx, "DONE", { size: set.size });
  return set;
}

function loadProductClassMap_(sheet, splitSet, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "loadProductClassMap_", {}), "loadProductClassMap_");
  const h = getHeaders_(sheet);
  const map = new Map();

  const pcCol = colAny_(h, ["Product Class", "Product Code"]);
  const sufKey = getSuffixKey_(h);
  if (pcCol === undefined || !sufKey || h["Job"] === undefined) return map;

  const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
  values.forEach(r => {
    const job = r[h["Job"]];
    const suf = r[h[sufKey]];
    const key = normalizeJobKey_(getCompositeJobKey_(job, suf, splitSet.has(normalizeJobKey_(job))));
    if (key) map.set(key, r[pcCol]);
  });

  logInfo_(ctx, "DONE", { size: map.size });
  return map;
}

function loadCustomerPOMap_(sheet, splitSet, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "loadCustomerPOMap_", {}), "loadCustomerPOMap_");
  const h = getHeaders_(sheet);
  const map = new Map();

  const custPoCol = colAny_(h, ["Cust PO", "Customer PO", "Customer PO#", "CustomerPO", "CustPO"]);
  const sufKey = getSuffixKey_(h);
  if (custPoCol === undefined || !sufKey || h["Job"] === undefined) return map;

  const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
  values.forEach(r => {
    const job = r[h["Job"]];
    const suf = r[h[sufKey]];
    const key = normalizeJobKey_(getCompositeJobKey_(job, suf, splitSet.has(normalizeJobKey_(job))));
    if (key) map.set(key, r[custPoCol]);
  });

  logInfo_(ctx, "DONE", { size: map.size });
  return map;
}

//==============================================================
// SHORTAGE LIST PIPELINE
//==============================================================
function loadJobMaterialDemands_(sheet, splitSet, pClassMap, producedSet, custPoMap, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "loadJobMaterialDemands_", {}), "loadJobMaterialDemands_");
  const h = getHeaders_(sheet);
  const demands = [];

  const sufKey = getSuffixKey_(h);
  if (!sufKey || h["Job"] === undefined) return demands;

  const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
  values.forEach(row => {
    const item = normalizeString_(row[h["Item"]]);
    if (!item || producedSet.has(item)) return;

    const job = normalizeJobKey_(row[h["Job"]]);
    const key = normalizeJobKey_(getCompositeJobKey_(job, row[h[sufKey]], splitSet.has(job)));
    if (!key) return;

    demands.push({
      item,
      description: row[h["Material Description"]],
      jobOrder: key,
      productClass: pClassMap.get(key) || "",
      custPo: (custPoMap && custPoMap.get) ? (custPoMap.get(key) || "") : "",
      qtyShort: parseFloat(row[h["Qty Short"]] || 0),
      um: row[h["U/M"]],
      assignedTo: row[h["Assigned To"]],
      jobEndDate: parseDate_(row[h["End Date"]]) || ""
    });
  });

  logInfo_(ctx, "DONE", { demands: demands.length });
  return demands;
}

function loadPoSupplies_(sheet, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "loadPoSupplies_", {}), "loadPoSupplies_");
  const h = getHeaders_(sheet);
  const map = new Map();

  const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
  values.forEach(r => {
    const item = normalizeString_(r[h["Item"]]);
    if (!item) return;

    if (!map.has(item)) map.set(item, []);
    const qty = parseFloat(r[h["Ordered"]] || 0) - parseFloat(r[h["Received"]] || 0);

    if (qty > 0) {
      map.get(item).push({
        po: r[h["PO"]],
        dueDate: parseDate_(r[h["Due Date"]]) || "",
        qtyOrdered: qty
      });
    }
  });

  logInfo_(ctx, "DONE", { items: map.size });
  return map;
}

function loadShortageData_(sheet, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "loadShortageData_", {}), "loadShortageData_");
  const h = getHeaders_(sheet);
  const map = new Map();

  const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
  values.forEach(r => {
    const key = normalizeJobKey_(r[h["Job Order"]]);
    const date = parseDate_(r[h["PO Due Date"]]);
    if (key && date) {
      if (!map.has(key)) map.set(key, []);
      map.get(key).push({ date, item: r[h["Item"]] });
    }
  });

  logInfo_(ctx, "DONE", { size: map.size });
  return map;
}

function loadCustomerPartData_(sheet, splitSet, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "loadCustomerPartData_", {}), "loadCustomerPartData_");
  const h = getHeaders_(sheet);
  const map = new Map();

  const sufKey = getSuffixKey_(h);
  if (!sufKey || h["Job"] === undefined) return map;

  const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
  values.forEach(r => {
    const job = normalizeJobKey_(r[h["Job"]]);
    const suf = r[h[sufKey]];
    const key = normalizeJobKey_(getCompositeJobKey_(job, suf, splitSet.has(job)));
    if (!key) return;

    const pct = parseFloat(r[h["Percent Complete"]] || 0);
    map.set(key, pct === 100 ? "CSP received" : (pct > 0 ? "CSP partially received" : "CSP not received"));
  });

  logInfo_(ctx, "DONE", { size: map.size });
  return map;
}

function allocateMaterials_(demandsList, suppliesMap, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "allocateMaterials_", {}), "allocateMaterials_");

  const demandsByItem = new Map();
  demandsList.forEach(d => {
    if (!demandsByItem.has(d.item)) demandsByItem.set(d.item, []);
    demandsByItem.get(d.item).push(d);
  });

  const results = [];
  for (const [item, demands] of demandsByItem.entries()) {
    const suppliesRaw = suppliesMap.get(item) || [];
    const supplies = suppliesRaw.map(s => ({
      po: s.po,
      dueDate: (s.dueDate instanceof Date) ? new Date(s.dueDate.getTime()) : (parseDate_(s.dueDate) || ""),
      qtyOrdered: Number(s.qtyOrdered || 0)
    }));

    demands.sort((a, b) => (parseDate_(a.jobEndDate) || 0) - (parseDate_(b.jobEndDate) || 0));
    supplies.sort((a, b) => (parseDate_(a.dueDate) || 0) - (parseDate_(b.dueDate) || 0));

    let sIdx = 0;
    for (const d of demands) {
      let needed = Number(d.qtyShort || 0);
      let firstPo = null;

      while (needed > 0 && sIdx < supplies.length) {
        if (!firstPo) firstPo = supplies[sIdx];

        const take = Math.min(needed, supplies[sIdx].qtyOrdered);
        needed -= take;
        supplies[sIdx].qtyOrdered -= take;

        if (supplies[sIdx].qtyOrdered <= 0.001) sIdx++;
      }

      results.push({
        ...d,
        status: needed <= 0.001 ? "ALLOCATED" : "BUY MORE",
        po: firstPo ? firstPo.po : "-",
        poDueDate: firstPo ? (firstPo.dueDate instanceof Date ? firstPo.dueDate : (parseDate_(firstPo.dueDate) || "")) : "",
        poQtyRemaining: firstPo ? Math.max(0, firstPo.qtyOrdered) : "-",
        qtyToBuy: needed > 0 ? needed : 0
      });
    }
  }

  logInfo_(ctx, "DONE", { demands: demandsList.length, items: demandsByItem.size, results: results.length });
  return results;
}

function writeShortageList_(ss, results, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "writeShortageList_", {}), "writeShortageList_");

  let sheet = ss.getSheetByName("Shortage List") || ss.insertSheet("Shortage List");

  const headers = [
    "Assigned To", "Job Order", "Product Class", "Cust PO", "Job End Date",
    "Item", "Material Description", "U/M", "Qty Short",
    "PO", "PO Due Date", "PO Qty Remaining", "Status", "Qty To Buy"
  ];

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), headers.length)).clearContent();
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");

  if (results.length > 0) {
    // CRITICAL: Format ID columns as Text BEFORE writing values to prevent 
    // Google Sheets from auto-parsing strings like "1825-01" as dates
    sheet.getRange(2, 2, results.length, 1).setNumberFormat("@"); // Job Order
    sheet.getRange(2, 4, results.length, 1).setNumberFormat("@"); // Cust PO
    sheet.getRange(2, 6, results.length, 1).setNumberFormat("@"); // Item
    sheet.getRange(2, 10, results.length, 1).setNumberFormat("@"); // PO

    const out = results.map(r => [
      r.assignedTo,
      r.jobOrder,
      r.productClass,
      r.custPo || "",
      (r.jobEndDate instanceof Date ? r.jobEndDate : (parseDate_(r.jobEndDate) || "")),
      r.item,
      r.description,
      r.um,
      r.qtyShort,
      r.po,
      (r.poDueDate instanceof Date ? r.poDueDate : (parseDate_(r.poDueDate) || "")),
      (r.poQtyRemaining === "-" ? "" : r.poQtyRemaining),
      r.status,
      r.qtyToBuy
    ]);

    sheet.getRange(2, 1, out.length, headers.length).setValues(out);

    sheet.getRange(2, 5, out.length, 1).setNumberFormat("M/dd/yyyy");  // Job End Date
    sheet.getRange(2, 11, out.length, 1).setNumberFormat("M/dd/yyyy"); // PO Due Date

    sheet.getRange(2, 1, out.length, headers.length).sort({ column: 11, ascending: true });
  }

  logInfo_(ctx, "DONE", { rows: results.length });
}

//==============================================================
// IMPORT PARSER
//==============================================================
function normalizeImportTargetName_(fileName) {
  return normalizeString_(fileName).replace(/\.(csv|txt|tsv)$/i, "");
}

function parseAndWriteCsvToSheet_(sheet, content, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "parseAndWriteCsvToSheet_", {}), "parseAndWriteCsvToSheet_");

  let text = normalizeString_(content);
  if (!text) return;

  text = text.replace(/^\uFEFF/, "").replace(/\u0000/g, "");
  const firstLine = text.split(/\r?\n/, 1)[0] || "";
  const delimiter = firstLine.includes("\t") ? "\t" : ",";

  let data = Utilities.parseCsv(text, delimiter);
  if (!data || data.length === 0) return;

  const trimmed = data.map(r => {
    let end = r.length;
    while (end > 0 && normalizeString_(r[end - 1]) === "") end--;
    return r.slice(0, end);
  });
  const maxLen = Math.max(...trimmed.map(r => r.length), 0);
  data = trimmed.map(r => r.concat(Array(Math.max(0, maxLen - r.length)).fill("")));

  sheet.clearContents();

  // CRITICAL: Format ID columns as Text BEFORE writing values
  if (data.length > 0) {
    const textColumns = ["item", "item no.", "item no", "item number", "job", "job order", "po", "cust po"];
    const headerRow = data[0].map(c => normalizeString_(c).toLowerCase());
    
    headerRow.forEach((colName, idx) => {
      if (textColumns.includes(colName)) {
        sheet.getRange(2, idx + 1, Math.max(1, data.length - 1), 1).setNumberFormat("@");
      }
    });
  }

  sheet.getRange(1, 1, data.length, maxLen).setValues(data);

  logInfo_(ctx, "DONE", { sheet: sheet.getName(), rows: data.length, cols: maxLen, delimiter });
}

const REQUIRED_IMPORT_HEADERS = {
  "ToExcel_CustomerPart": ["Job", "Suffix", "Item", "Percent Complete", "Status"],
  "ToExcel_JobMaterialsListing": ["Job", "Suffix", "Item", "Material Description", "U/M", "Qty Short", "Assigned To", "End Date", "Status"],
  "ToExcel_JobOrders": ["Job", "Job Suffix", "Due Date", "End Date", "Item", "Customer", "Status", "Assigned To", "Cust PO"],
  "ToExcel_PurchaseOrderListing": ["PO", "Item", "Ordered", "Received", "Due Date", "Status"]
};

function validateImportedSheet_(sheetName, sheet, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "validateImportedSheet_", {}), `validateImportedSheet_(${sheetName})`);

  const required = REQUIRED_IMPORT_HEADERS[sheetName];
  if (!required) return [];

  const h = getHeaders_(sheet);

  const out = [];
  required.forEach(name => {
    if (name === "Suffix") {
      if (h["Suffix"] === undefined && h["Job Suffix"] === undefined) out.push("Missing expected column 'Suffix' (or 'Job Suffix')");
      return;
    }
    if (h[name] === undefined) out.push(`Missing expected column '${name}'`);
  });

  if (out.length) logWarn_(ctx, "VALIDATION WARN", { warnings: out });
  else logInfo_(ctx, "VALIDATION OK");

  return out;
}

//==============================================================
// SHEET UTILITIES
//==============================================================
function getActualLastRow_(sheet, col) {
  const last = sheet.getLastRow();
  const data = sheet.getRange(1, col, last, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "" && data[i][0] !== null && data[i][0] !== undefined) return i + 1;
  }
  return 0;
}

function processMoveOperation_(src, tar, rows, desc) {
  if (!rows || !rows.length) return;

  // Group contiguous rows into blocks to minimize API copy operations.
  const sorted = rows.slice().sort((a, b) => a - b);
  let nextRow = tar.getLastRow() + 1;
  const lastCol = src.getLastColumn();

  const ranges = [];
  let start = sorted[0], prev = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    const r = sorted[i];
    if (r === prev + 1) { 
      prev = r; 
      continue; 
    }
    ranges.push([start, prev]);
    start = prev = r;
  }
  ranges.push([start, prev]);

  // Execute copy blocks
  ranges.forEach(([s, e]) => {
    const numRows = e - s + 1;
    const sourceRange = src.getRange(s, 1, numRows, lastCol);
    const targetRange = tar.getRange(nextRow, 1);
    sourceRange.copyTo(targetRange, { contentsOnly: false });
    nextRow += numRows;
  });

  batchDeleteRows_(src, sorted);
}

function batchDeleteRows_(sheet, rows) {
  if (!rows.length) return;
  rows = rows.slice().sort((a, b) => a - b);

  const ranges = [];
  let start = rows[0], prev = rows[0];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r === prev + 1) { prev = r; continue; }
    ranges.push([start, prev]);
    start = prev = r;
  }
  ranges.push([start, prev]);

  for (let i = ranges.length - 1; i >= 0; i--) {
    const [s, e] = ranges[i];
    sheet.deleteRows(s, e - s + 1);
  }
}

//==============================================================
// USER RESOLUTION (System_Users)
//==============================================================
function resolveUserName_(key) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = SpreadsheetApp.openById(CONFIG.LOG_SHEET_ID).getSheetByName(CONFIG.USERS_SHEET_NAME);
    if (!sheet) return key;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(key)) return data[i][1];
    }

    sheet.appendRow([key, "Unknown (Rename Me)"]);
    return "Unknown (Rename Me)";
  } catch (e) {
    return key;
  } finally {
    lock.releaseLock();
  }
}

//==============================================================
// AUDIT: SyteLine jobs not tracked (✅ PO column filled with Cust PO)
//==============================================================
function auditSyteLineJobsNotTracked_(ss, sourceData, reportSheetNames, runId, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(runId, "auditSyteLineJobsNotTracked_", { spreadsheet: ss.getName() }), "auditSyteLineJobsNotTracked_");

  const tracked = new Set();

  reportSheetNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const h = getHeaders_(sh);
    if (h["Job Order"] === undefined) return;

    const lastRow = getActualLastRow_(sh, h["Job Order"] + 1);
    if (lastRow <= 1) return;

    const vals = sh.getRange(2, h["Job Order"] + 1, lastRow - 1, 1).getDisplayValues();
    vals.forEach(r => {
      const k = normalizeJobKey_(r[0]);
      if (!k) return;
      tracked.add(normalizeJobKeyForCompare_(k));
    });
  });

  const all = Array.from(sourceData.jobsInSource.values());
  const missing = [];
  for (let i = 0; i < all.length; i++) {
    const key = normalizeJobKey_(all[i]);
    const norm = normalizeJobKeyForCompare_(key);
    if (!tracked.has(norm)) missing.push(key);
  }

  const total = missing.length;
  const cap = Number(CONFIG.AUDIT_SOURCE_ONLY_MAX_ENTRIES || 500);
  const logged = Math.min(total, cap);

  const entries = [];

  entries.push({
    reportSheet: "SyteLine Audit",
    projectCoordinator: "",
    jobOrder: "TOOL ACTION",
    customer: "SYTELINE NOT TRACKED",
    po: "",
    qty: "",
    itemNo: "",
    severity: (total > 0 ? "WARN" : "INFO"),
    runId,
    changes: [
      `Run ${runId}: Found ${total} SyteLine job(s) not listed in tracking sheets (${reportSheetNames.join(", ")}).`,
      `Logged ${logged} item(s) this run (cap=${cap}).`
    ]
  });

  for (let i = 0; i < logged; i++) {
    const key = missing[i];

    const due = dateToStr_(sourceData.dateMap.get(key));
    const end = dateToStr_(sourceData.endMap.get(key));
    const item = normalizeString_(sourceData.itemMap.get(key) || "");
    const cust = normalizeString_(sourceData.customerMap.get(key) || "");
    const status = normalizeString_(sourceData.statusMap.get(key) || "");
    const custPo = normalizeString_(sourceData.custPoMap.get(key) || "");

    entries.push({
      reportSheet: "SyteLine Only",
      projectCoordinator: "",
      jobOrder: key,
      customer: "SYTELINE NOT TRACKED",
      po: custPo || "",
      qty: "",
      itemNo: item,
      severity: "WARN",
      runId,
      changes: [
        `Run ${runId}: Job exists in SyteLine export but not in tracking sheets.`,
        `Customer: ${cust || "-"}`,
        `Status: ${status || "-"}`,
        `Cust PO: ${custPo || "-"}`,
        `Due date: ${due || "-"}`,
        `End date: ${end || "-"}`
      ]
    });
  }

  logInfo_(ctx, "AUDIT RESULT", { tracked: tracked.size, sourceJobs: all.length, missing: total, cap, logged });
  return { total, logged, entries };
}
