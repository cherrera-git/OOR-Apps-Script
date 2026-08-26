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

function parseNumber_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const clean = String(v).replace(/,/g, "");
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
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
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "string" && v.trim() === "") return null;
  
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

  const pMatches = note.match(/P-(?:\d{1,2}\/\d{1,2}|TBD)[^;]*/gi) || [];
  const pFull = pMatches.map(s => s.trim()).join("; ");

  const cspMatch = note.match(/CSP[^;]*/i);
  const csp = cspMatch ? normalizeString_(cspMatch[0]) : "";

  const custom = note
    .replace(/End Date=[^;]+;?/gi, "")
    .replace(/P-(?:\d{1,2}\/\d{1,2}|TBD)[^;]*;?/gi, "") 
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
// EXTERNAL LOG (Bypassed per Refactoring)
//==============================================================
function writeToExternalLog_(logs, targetSheetName, parentCtx) {
  return;
}

function logToolAction_(action, details, sheetName, targetSheetName, severity, runId, parentCtx) {
  return;
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
  const h = getHeaders_(sheet);
  const logs = [];

  const missing = [];
  if (h["Job Order"] === undefined) missing.push("Job Order");
  if (h["MTL Due Date"] === undefined) missing.push("MTL Due Date");
  if (h["End Date Notes"] === undefined) missing.push("End Date Notes");
  if (missing.length) return { logs };

  const lastRow = getActualLastRow_(sheet, h["Job Order"] + 1);
  if (lastRow <= 1) return { logs };

  const raw = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const disp = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();

  const dueCol = h["MTL Due Date"] + 1;
  const noteCol = h["End Date Notes"] + 1;
  const hasPcCol = (h["Project Coordinator"] !== undefined);
  const pcCol = hasPcCol ? (h["Project Coordinator"] + 1) : null;
  const pcNotesCol = (h["PC Notes"] !== undefined) ? (h["PC Notes"] + 1) : 21;

  const dueUpdates = new Map();
  const noteUpdates = new Map();
  const pcUpdates = new Map();
  const pcNotesUpdates = new Map();

  let unmatchedCount = 0;
  const sheetName = sheet.getName();

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
      if (summary && summary.unmatchedSamples) {
        if (sheetName === CONFIG.STOCK_SHEET_NAME) {
          if (!summary.unmatchedStockItems) summary.unmatchedStockItems = [];
          summary.unmatchedStockItems.push(jobKey || `(row ${rowNum})`);
        }
        if ((summary.unmatchedSamples[sheetName] || []).length < 10) {
          summary.unmatchedSamples[sheetName].push(jobKey || `(row ${rowNum})`);
        }
      }
      continue;
    }

    const oldDue = parseDate_(raw[i][h["MTL Due Date"]]);
    const oldNote = normalizeString_(disp[i][h["End Date Notes"]]);
    const oldNoteNorm = normalizeNotes_(oldNote);
    const oldPc = hasPcCol ? normalizeString_(disp[i][h["Project Coordinator"]]) : "";

    const srcDue = parseDate_(sourceData.dateMap.get(dataKey));
    const srcEnd = parseDate_(sourceData.endMap.get(dataKey));
    const srcAssignedTo = normalizeString_(sourceData.assignedToMap.get(dataKey) || "");

    const endStr = srcEnd ? `End Date=${Utilities.formatDate(srcEnd, Session.getScriptTimeZone(), "M/dd/yy")}` : "";
    const shortageList = shortageData.get(dataKey) || [];
    shortageList.sort((a, b) => (parseDate_(a.date) ? parseDate_(a.date).getTime() : Infinity) - (parseDate_(b.date) ? parseDate_(b.date).getTime() : Infinity));

    const pStr = shortageList.map(s => {
      const d = parseDate_(s.date);
      const dStr = d ? Utilities.formatDate(d, Session.getScriptTimeZone(), "M/dd") : "TBD";
      return `P-${dStr} (${normalizeString_(s.item)})`;
    }).join("; ");

    const cspStr = cspData.get(dataKey) || "";
    const cleanCustom = oldNote.replace(/End Date=[^;]+/gi, "").replace(/P-(?:\d{1,2}\/\d{1,2}|TBD)[^;]*/gi, "").replace(/CSP[^;]*/gi, "").split(";").map(s => s.trim()).filter(Boolean);
    const newNote = [endStr, pStr, cspStr, ...cleanCustom].filter(Boolean).join("; ");
    
    let dueChanged = false;
    let notesChanged = false;
    let pcChanged = false;

    if (srcDue && (!oldDue || srcDue.getTime() !== oldDue.getTime())) {
      dueUpdates.set(rowNum, srcDue);
      dueChanged = true;
    }

    if (oldNoteNorm !== normalizeNotes_(newNote)) {
      noteUpdates.set(rowNum, newNote);
      notesChanged = true;
    }

    if (hasPcCol && oldPc === "" && srcAssignedTo !== "") {
      pcUpdates.set(rowNum, srcAssignedTo);
      pcChanged = true;
    }

    if (dueChanged || notesChanged || pcChanged) {
      const rowChangelogs = [];
      if (dueChanged) {
        rowChangelogs.push(`* Due date: ${oldDue ? Utilities.formatDate(oldDue, Session.getScriptTimeZone(), "M/dd") : "Blank"} → ${srcDue ? Utilities.formatDate(srcDue, Session.getScriptTimeZone(), "M/dd") : "Cleared"}`);
      }
      if (pcChanged) {
        rowChangelogs.push(`* PC: ${oldPc || "Blank"} → ${srcAssignedTo || "Cleared"}`);
      }
      if (notesChanged) {
        const oldParts = parseNotesParts_(oldNote);
        const newParts = parseNotesParts_(newNote);
        
        if (oldParts.endDate !== newParts.endDate) {
          rowChangelogs.push(`* End Date: ${oldParts.endDate ? oldParts.endDate.replace(/\/\d{2,4}$/, '') : "Blank"} → ${newParts.endDate ? newParts.endDate.replace(/\/\d{2,4}$/, '') : "Cleared"}`);
        }
        
        const oldPArr = oldParts.pFull ? oldParts.pFull.split(';').map(s => s.trim()).filter(Boolean) : [];
        const newPArr = newParts.pFull ? newParts.pFull.split(';').map(s => s.trim()).filter(Boolean) : [];
        const added = newPArr.filter(x => !oldPArr.includes(x));
        const removed = oldPArr.filter(x => !newPArr.includes(x));
        
        if (added.length || removed.length) {
          const getBase = (s) => (s.match(/\((.*)\)/) ? s.match(/\((.*)\)/)[1].trim() : s);
          const getD = (s) => (s.match(/(P-(?:\d{1,2}\/\d{1,2}|TBD))/i) ? s.match(/(P-(?:\d{1,2}\/\d{1,2}|TBD))/i)[1].trim() : "");
          
          const remMap = new Map(); removed.forEach(r => remMap.set(getBase(r), r));
          const addMap = new Map(); added.forEach(a => addMap.set(getBase(a), a));
          
          removed.forEach(r => { if (!addMap.has(getBase(r))) rowChangelogs.push(`* Arrived: ${getBase(r)}`); });
          added.forEach(a => {
            if (remMap.has(getBase(a))) {
              const o = getD(remMap.get(getBase(a))).replace('P-', '');
              const n = getD(a).replace('P-', '');
              if (o) rowChangelogs.push(`* Shifted: ${getBase(a)} (P-${o}→P-${n})`);
              else rowChangelogs.push(`* New Short: ${a}`);
            } else {
              rowChangelogs.push(`* New Short: ${a}`);
            }
          });
        }
        
        // Log changes to CSP Status cleanly
        if (oldParts.csp !== newParts.csp) {
          const n = (newParts.csp || "").toLowerCase();
          if (n === "") {
            rowChangelogs.push(`* CSP Cleared`);
          } else {
            rowChangelogs.push(`* ${n.includes("not received") ? "Waiting on CSP" : (n.includes("partially") ? "Partial CSP Arrived" : "CSP Arrived")}`);
          }
        }
      }

      if (rowChangelogs.length > 0 && typeof buildCleanPCNotes_ === 'function') {
        const currentPCNotes = normalizeString_(disp[i][pcNotesCol - 1]);
        pcNotesUpdates.set(rowNum, buildCleanPCNotes_(currentPCNotes, rowChangelogs));
      }
      logs.push({ jobOrder: jobKey });
    }
  }

  const dueWritten = applyColumnUpdates_(sheet, dueCol, dueUpdates);
  const noteWritten = applyColumnUpdates_(sheet, noteCol, noteUpdates);
  const pcWritten = (hasPcCol ? applyColumnUpdates_(sheet, pcCol, pcUpdates) : 0);
  applyColumnUpdates_(sheet, pcNotesCol, pcNotesUpdates);

  if (summary) {
    summary.dueDateChanges += dueWritten;
    summary.noteChanges += noteWritten;
    summary.pcFilledChanges += pcWritten;
    summary.unmatchedBySheet[sheetName] = unmatchedCount;
  }

  return { logs };
}

//==============================================================
// SOURCE DATA LOADERS
//==============================================================
function loadSourceJobData_(sheet, splitSet, parentCtx) {
  const h = getHeaders_(sheet);
  const data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();

  const sufKey = getSuffixKey_(h) || "Job Suffix";
  const itemCol = colAny_(h, ["Item", "Item No.", "Item No", "Item Number"]);
  const custCol = colAny_(h, ["Customer", "Customer Name", "Cust Num", "Cust", "CustomerNum"]);
  const statusCol = colAny_(h, ["Status", "Job Status", "Stat"]);
  const custPoCol = colAny_(h, ["Cust PO", "Customer PO", "Customer PO#", "CustomerPO", "CustPO"]);
  const assignedToCol = colAny_(h, ["Assigned To", "AssignedTo", "Assigned"]);

  const map = { jobsInSource: new Set(), dateMap: new Map(), endMap: new Map(), itemMap: new Map(), customerMap: new Map(), statusMap: new Map(), custPoMap: new Map(), assignedToMap: new Map() };

  data.forEach(r => {
    const key = normalizeJobKey_(getCompositeJobKey_(r[h["Job"]], r[h[sufKey]], splitSet.has(normalizeJobKey_(r[h["Job"]]))));
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
  return map;
}

function scanForSubassemblies_(sheet, parentCtx) {
  const h = getHeaders_(sheet);
  const set = new Set();
  const sufKey = getSuffixKey_(h);
  if (!sufKey || h["Job"] === undefined) return set;
  sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues().forEach(r => {
    if (parseInt(r[h[sufKey]], 10) > 0) set.add(normalizeJobKey_(r[h["Job"]]));
  });
  return set;
}

function loadProducedItemsSet_(sheet, parentCtx) {
  const h = getHeaders_(sheet);
  const set = new Set();
  const itemCol = colAny_(h, ["Item", "Item No.", "Item No", "Item Number"]);
  if (itemCol === undefined) return set;
  sheet.getRange(2, itemCol + 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues().forEach(r => { if (r[0]) set.add(normalizeString_(r[0])); });
  return set;
}

function loadProductClassMap_(sheet, splitSet, parentCtx) {
  const h = getHeaders_(sheet);
  const map = new Map();
  const pcCol = colAny_(h, ["Product Class", "Product Code"]);
  const sufKey = getSuffixKey_(h);
  if (pcCol === undefined || !sufKey || h["Job"] === undefined) return map;
  sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues().forEach(r => {
    const key = normalizeJobKey_(getCompositeJobKey_(r[h["Job"]], r[h[sufKey]], splitSet.has(normalizeJobKey_(r[h["Job"]]))));
    if (key) map.set(key, r[pcCol]);
  });
  return map;
}

function loadCustomerPOMap_(sheet, splitSet, parentCtx) {
  const h = getHeaders_(sheet);
  const map = new Map();
  const custPoCol = colAny_(h, ["Cust PO", "Customer PO", "Customer PO#", "CustomerPO", "CustPO"]);
  const sufKey = getSuffixKey_(h);
  if (custPoCol === undefined || !sufKey || h["Job"] === undefined) return map;
  sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues().forEach(r => {
    const key = normalizeJobKey_(getCompositeJobKey_(r[h["Job"]], r[h[sufKey]], splitSet.has(normalizeJobKey_(r[h["Job"]]))));
    if (key) map.set(key, r[custPoCol]);
  });
  return map;
}

//==============================================================
// SHORTAGE LIST PIPELINE
//==============================================================
function loadJobMaterialDemands_(sheet, splitSet, pClassMap, producedSet, custPoMap, parentCtx) {
  const h = getHeaders_(sheet);
  const demands = [];

  const sufKey = getSuffixKey_(h);
  if (!sufKey || h["Job"] === undefined) return demands;

  const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
  values.forEach(row => {
    // FIX: Ignore lines with no actual material shortage
    const qtyShort = parseNumber_(row[h["Qty Short"]] || 0);
    if (qtyShort <= 0) return;
    
    // FIX: Ignore completed or closed material lines
    const status = (h["Status"] !== undefined) ? normalizeString_(row[h["Status"]]).toLowerCase() : "";
    if (status === "complete" || status === "closed") return;

    const item = normalizeString_(row[h["Item"]]);
    if (!item || producedSet.has(item)) return;

    // FIX: Strictly exclude Customer Supplied Parts (CSPs) from the Purchase Order Shortage List.
    const desc = normalizeString_(row[h["Material Description"]]);
    const isCSP = item.toUpperCase().includes("CSP") || 
                  desc.toUpperCase().includes("CSP") || 
                  item.toUpperCase().includes("CUSTOMER PART") || 
                  desc.toUpperCase().includes("CUSTOMER PART");
    
    if (isCSP) return; 

    const job = normalizeJobKey_(row[h["Job"]]);
    const key = normalizeJobKey_(getCompositeJobKey_(job, row[h[sufKey]], splitSet.has(job)));
    if (!key) return;

    demands.push({
      item,
      description: desc,
      jobOrder: key,
      productClass: pClassMap.get(key) || "",
      custPo: (custPoMap && custPoMap.get) ? (custPoMap.get(key) || "") : "",
      qtyShort: qtyShort,
      um: row[h["U/M"]],
      assignedTo: row[h["Assigned To"]],
      jobEndDate: parseDate_(row[h["Due Date"]] !== undefined ? row[h["Due Date"]] : row[h["End Date"]]) || ""
    });
  });

  return demands;
}

function loadPoSupplies_(sheet, parentCtx) {
  const h = getHeaders_(sheet);
  const map = new Map();
  sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues().forEach(r => {
    const item = normalizeString_(r[h["Item"]]);
    if (!item) return;
    if (!map.has(item)) map.set(item, []);
    const qty = parseNumber_(r[h["Ordered"]] || 0) - parseNumber_(r[h["Received"]] || 0);
    if (qty > 0) map.get(item).push({ po: r[h["PO"]], dueDate: parseDate_(r[h["Due Date"]]) || "", qtyOrdered: qty });
  });
  return map;
}

function loadShortageData_(sheet, parentCtx) {
  const h = getHeaders_(sheet);
  const map = new Map();
  sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues().forEach(r => {
    const key = normalizeJobKey_(r[h["Job Order"]]);
    const date = parseDate_(r[h["PO Due Date"]]);
    const item = normalizeString_(r[h["Item"]]); 
    if (key && item) {
      if (!map.has(key)) map.set(key, []);
      map.get(key).push({ date, item });
    }
  });
  return map;
}

function loadCustomerPartData_(sheet, splitSet, parentCtx) {
  const h = getHeaders_(sheet);
  const map = new Map();
  const sufKey = getSuffixKey_(h);
  
  // Safety check: Make sure Percent Complete exists in the merged sheet
  if (!sufKey || h["Job"] === undefined || h["Percent Complete"] === undefined) return map;
  
  sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues().forEach(r => {
    // FIX: SyteLine populates Percent Complete for standard materials too.
    // We must strictly filter for rows where the Item or Description indicates it is a CSP.
    const item = normalizeString_(r[h["Item"]]);
    const desc = normalizeString_(r[h["Material Description"]]);
    const isCSP = item.toUpperCase().includes("CSP") || 
                  desc.toUpperCase().includes("CSP") || 
                  item.toUpperCase().includes("CUSTOMER PART") || 
                  desc.toUpperCase().includes("CUSTOMER PART");
    
    if (!isCSP) return;

    const pctVal = r[h["Percent Complete"]];
    if (pctVal === "" || pctVal === undefined || pctVal === null) return;
    
    const key = normalizeJobKey_(getCompositeJobKey_(normalizeJobKey_(r[h["Job"]]), r[h[sufKey]], splitSet.has(normalizeJobKey_(r[h["Job"]]))));
    if (!key) return;
    
    const pct = parseNumber_(pctVal || 0);
    const newStatus = pct === 100 ? "CSP received" : (pct > 0 ? "CSP partially received" : "CSP not received");
    
    // Safely handle multiple CSP items on a single job. 
    // Prioritize the "worst-case" status so the job doesn't falsely show as cleared.
    const currentStatus = map.get(key);
    if (!currentStatus || currentStatus === "CSP received") {
       map.set(key, newStatus); 
    } else if (currentStatus === "CSP partially received" && newStatus === "CSP not received") {
       map.set(key, newStatus);
    }
  });
  return map;
}

function allocateMaterials_(demandsList, suppliesMap, parentCtx) {
  const demandsByItem = new Map();
  demandsList.forEach(d => {
    if (!demandsByItem.has(d.item)) demandsByItem.set(d.item, []);
    demandsByItem.get(d.item).push(d);
  });
  const results = [];
  
  const getTime = (d) => { const date = parseDate_(d); return date ? date.getTime() : Infinity; };
  
  for (const [item, demands] of demandsByItem.entries()) {
    const suppliesRaw = suppliesMap.get(item) || [];
    const supplies = suppliesRaw.map(s => ({ po: s.po, dueDate: (s.dueDate instanceof Date) ? new Date(s.dueDate.getTime()) : (parseDate_(s.dueDate) || ""), qtyOrdered: parseNumber_(s.qtyOrdered || 0) }));
    
    demands.sort((a, b) => getTime(a.jobEndDate) - getTime(b.jobEndDate));
    supplies.sort((a, b) => getTime(a.dueDate) - getTime(b.dueDate));
    
    let sIdx = 0;
    for (const d of demands) {
      let needed = parseNumber_(d.qtyShort || 0);
      const usedPos = [];
      
      while (needed > 0 && sIdx < supplies.length) {
        const currentPo = supplies[sIdx];
        const take = Math.min(needed, currentPo.qtyOrdered);
        
        if (take > 0) {
          usedPos.push({ po: currentPo.po, dueDate: currentPo.dueDate, qtyRemaining: currentPo.qtyOrdered - take });
        }
        
        needed -= take;
        currentPo.qtyOrdered -= take;
        if (currentPo.qtyOrdered <= 0.001) sIdx++;
      }
      
      let poStr = "-", poDueDate = "", poQtyRem = "-";
      
      if (usedPos.length > 0) {
        poStr = usedPos.map(u => u.po).join(", ");
        poDueDate = usedPos[0].dueDate; 
        poQtyRem = usedPos[usedPos.length - 1].qtyRemaining; 
      }

      results.push({
        ...d, 
        status: needed <= 0.001 ? "ALLOCATED" : "BUY MORE",
        po: poStr, 
        poDueDate: poDueDate instanceof Date ? poDueDate : (parseDate_(poDueDate) || ""),
        poQtyRemaining: poQtyRem, 
        qtyToBuy: needed > 0 ? needed : 0
      });
    }
  }
  return results;
}

function writeShortageList_(ss, results, parentCtx) {
  let sheet = ss.getSheetByName("Shortage List") || ss.insertSheet("Shortage List");

  const headers = [
    "Assigned To", "Job Order", "Product Class", "Cust PO", "Job Due Date",
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

    sheet.getRange(2, 5, out.length, 1).setNumberFormat("M/dd/yyyy"); // Job Due Date
    sheet.getRange(2, 11, out.length, 1).setNumberFormat("M/dd/yyyy"); // PO Due Date

    // Sort by Column 5 (Job Due Date) from oldest to newest
    sheet.getRange(2, 1, out.length, headers.length).sort({ column: 5, ascending: true });
  }
}

//==============================================================
// IMPORT PARSER
//==============================================================
function normalizeImportTargetName_(fileName) { return normalizeString_(fileName).replace(/\.(csv|txt|tsv)$/i, ""); }

function parseAndWriteCsvToSheet_(sheet, content, parentCtx) {
  let text = normalizeString_(content).replace(/^\uFEFF/, "").replace(/\u0000/g, "");
  if (!text) return;
  const delimiter = (text.split(/\r?\n/, 1)[0] || "").includes("\t") ? "\t" : ",";
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
  if (data.length > 0) {
    const textColumns = ["item", "item no.", "item no", "item number", "job", "job order", "po", "cust po"];
    data[0].map(c => normalizeString_(c).toLowerCase()).forEach((colName, idx) => {
      if (textColumns.includes(colName)) sheet.getRange(2, idx + 1, Math.max(1, data.length - 1), 1).setNumberFormat("@");
    });
  }
  sheet.getRange(1, 1, data.length, maxLen).setValues(data);
}

const REQUIRED_IMPORT_HEADERS = {
  "ToExcel_JobMaterialsListing": ["Job", "Suffix", "Item", "Material Description", "U/M", "Qty Short", "Assigned To", "End Date", "Status", "Percent Complete"],
  "ToExcel_JobOrders": ["Job", "Job Suffix", "Due Date", "End Date", "Item", "Customer", "Status", "Assigned To", "Cust PO"],
  "ToExcel_PurchaseOrderListing": ["PO", "Item", "Ordered", "Received", "Due Date", "Status"]
};

function validateImportedSheet_(sheetName, sheet, parentCtx) {
  const required = REQUIRED_IMPORT_HEADERS[sheetName];
  if (!required) return [];
  const h = getHeaders_(sheet), out = [];
  required.forEach(name => {
    if (name === "Suffix") { if (h["Suffix"] === undefined && h["Job Suffix"] === undefined) out.push("Missing expected column 'Suffix'"); return; }
    if (h[name] === undefined) out.push(`Missing expected column '${name}'`);
  });
  return out;
}

//==============================================================
// SHEET UTILITIES
//==============================================================
function getActualLastRow_(sheet, col) {
  const data = sheet.getRange(1, col, sheet.getLastRow(), 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) if (data[i][0] !== "" && data[i][0] !== null && data[i][0] !== undefined) return i + 1;
  return 0;
}

function processMoveOperation_(src, tar, rows, desc) {
  if (!rows || !rows.length) return;
  const maxSrcRows = src.getMaxRows();
  const validRows = [...new Set(rows.map(Number).filter(r => !isNaN(r) && r > 0 && r <= maxSrcRows))].sort((a, b) => a - b);
  if (!validRows.length) return;
  let nextRow = tar.getLastRow() + 1;
  const lastCol = src.getLastColumn();
  if (lastCol === 0) return;
  const requiredTarRows = nextRow + validRows.length - 1;
  if (requiredTarRows > tar.getMaxRows()) tar.insertRowsAfter(tar.getMaxRows(), requiredTarRows - tar.getMaxRows());
  const ranges = [];
  let start = validRows[0], prev = validRows[0];
  for (let i = 1; i < validRows.length; i++) {
    if (validRows[i] === prev + 1) { prev = validRows[i]; continue; }
    ranges.push([start, prev]); start = prev = validRows[i];
  }
  ranges.push([start, prev]);
  ranges.forEach(([s, e]) => {
    const numRows = e - s + 1;
    src.getRange(s, 1, numRows, lastCol).copyTo(tar.getRange(nextRow, 1), { contentsOnly: false });
    nextRow += numRows;
  });
  batchDeleteRows_(src, validRows);
}

function batchDeleteRows_(sheet, rows) {
  if (!rows || !rows.length) return;
  const maxRows = sheet.getMaxRows();
  const sorted = [...new Set(rows.map(Number).filter(r => !isNaN(r) && r > 0 && r <= maxRows))].sort((a, b) => a - b);
  if (!sorted.length) return;
  const ranges = [];
  let start = sorted[0], prev = sorted[0];
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === prev + 1) { prev = sorted[i]; continue; }
    ranges.push([start, prev]); start = prev = sorted[i];
  }
  ranges.push([start, prev]);
  for (let i = ranges.length - 1; i >= 0; i--) {
    const [s, e] = ranges[i];
    const numRows = e - s + 1, currentMax = sheet.getMaxRows();
    if (s <= currentMax) {
      const safeNumRows = Math.min(numRows, currentMax - s + 1);
      if (safeNumRows > 0) sheet.deleteRows(s, safeNumRows);
    }
  }
}

//==============================================================
// AUDIT: SyteLine jobs not tracked (Unique Job Order Fuzzy Match)
//==============================================================
function auditSyteLineJobsNotTracked_(ss, sourceData, reportSheetNames, runId, parentCtx) {
  const ctx = childCtx_(parentCtx || createLogCtx_(runId, "auditSyteLineJobsNotTracked_", { spreadsheet: ss.getName() }), "auditSyteLineJobsNotTracked_");

  const tracked = new Set();
  const makeFuzzy = (k) => String(k).toUpperCase().replace(/[\s\-]/g, "");

  reportSheetNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    
    // STRICT CHECK: Look exclusively for unique Job Order or Job columns (ignoring Sales Order to prevent duplicate grouping false positives)
    const h = getHeaders_(sh);
    const jobColIdx = colAny_(h, ["Job Order", "Job"]);
    if (jobColIdx === undefined) return;
    
    const lastRow = getActualLastRow_(sh, jobColIdx + 1);
    if (lastRow <= 1) return;
    
    sh.getRange(2, jobColIdx + 1, lastRow - 1, 1).getDisplayValues().forEach(r => {
      const k = normalizeJobKey_(r[0]);
      if (k) {
        const norm = normalizeJobKeyForCompare_(k).toUpperCase();
        tracked.add(norm); 
        tracked.add(makeFuzzy(norm)); 
        
        const base = norm.split(/[\s\-]/)[0];
        if (base) tracked.add(base); 
      }
    });
  });

  const all = Array.from(sourceData.jobsInSource.values());
  const missing = [];
  
  for (let i = 0; i < all.length; i++) {
    const key = normalizeJobKey_(all[i]);
    const norm = normalizeJobKeyForCompare_(key).toUpperCase();
    const fuzzy = makeFuzzy(norm);
    const base = norm.split(/[\s\-]/)[0];
    
    if (!tracked.has(norm) && !tracked.has(fuzzy) && !tracked.has(base)) {
      missing.push(key);
    }
  }

  const missingArray = missing.map(key => {
    const cust = normalizeString_(sourceData.customerMap.get(key) || "");
    const status = normalizeString_(sourceData.statusMap.get(key) || "");
    const custPo = normalizeString_(sourceData.custPoMap.get(key) || "");
    const item = normalizeString_(sourceData.itemMap.get(key) || "");
    return `${key} | ${cust || '-'} | ${status || '-'} | PO: ${custPo || '-'} | Item: ${item || '-'}`;
  });

  return { total: missing.length, logged: 0, entries: [], missingArray };
}

function escapeRegex_(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
