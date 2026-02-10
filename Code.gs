/**
 * Code.gs
 * - CONFIG defined here.
 * - Main logic for Cleanup and Update workflows.
 */

const CONFIG = {
  // External centralized log spreadsheet
  LOG_SHEET_ID: "1EKxcC6E02op2Q4v1kSusUByNDWHuvUVrw94ICPL2eWc",
  USERS_SHEET_NAME: "System_Users",

  // Dialog settings
  DIALOG_WIDTH: 720,
  DIALOG_HEIGHT: 610,

  // Settings
  WORK_START_HOUR: 6,
  WORK_END_HOUR: 18,
  SHARED_ACCOUNT_EMAIL: "mtntechstaff@gmail.com",
  AUDIT_SOURCE_ONLY_ENABLED: true,
  AUDIT_SOURCE_ONLY_MAX_ENTRIES: 500,
  UNMATCHED_SAMPLE_LIMIT_PER_SHEET: 10,

  // Safety
  SAFETY_ENABLED: true,
  SAFETY_PHRASE_CLEANUP: "CLEANUP",
  SAFETY_PHRASE_UPDATE: "UPDATE",
  SAFETY_IMPORT_PHRASE: "IMPORT",
  SAFETY_REQUIRE_IMPORT_META: true,

  // Sheets
  STOCK_SHEET_NAME: "STOCK ITEMS",
};

//==============================================================
// MENU
//==============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Update Tools")
    .addItem("Step 1: Cleanup (Move & Archive)", "showCleanupDialog")
    .addSeparator()
    .addItem("Step 2: Import & Update", "showImportDialog")
    .addToUi();
}

function generateRunId_() {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const rand = Math.random().toString(36).slice(2, 6).toUpperCase();
  return `${ts}-${rand}`;
}

//==============================================================
// HTML DIALOG DISPATCHERS
//==============================================================
function showCleanupDialog() {
  const t = HtmlService.createTemplateFromFile('confirmation');
  t.title = "Confirm Cleanup";
  t.message = "1. Move New Orders to OOR<br>2. Archive Closed/Invalid from OOR<br><br>This will modify your tracking sheets.";
  t.phrase = CONFIG.SAFETY_PHRASE_CLEANUP;
  t.action = "CLEANUP";
  
  const html = t.evaluate().setWidth(420).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, "Step 1: Cleanup");
}

function showUpdateConfirmation() {
  const t = HtmlService.createTemplateFromFile('confirmation');
  t.title = "Ready for Update?";
  t.message = "Files imported successfully.<br><br>Proceed with Full Report Update (Shortage List + Tracking Sheets)?";
  t.phrase = CONFIG.SAFETY_PHRASE_UPDATE;
  t.action = "UPDATE";
  
  const html = t.evaluate().setWidth(420).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, "Step 2: Confirm Update");
}

function processSafetyGate(action) {
  if (action === "CLEANUP") {
    return runCleanupSequence_();
  } else if (action === "UPDATE") {
    return runFullUpdateSequence_();
  }
}

//==============================================================
// MERGED STEP 1: CLEANUP LOGIC
//==============================================================
function runCleanupSequence_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const runId = generateRunId_();
  const ctx = createLogCtx_(runId, "runCleanupSequence", { spreadsheet: ss.getName() });

  logInfo_(ctx, "START");

  uiToastPhase_(ss, runId, "Moving New Orders...");
  const moveCount = executeMoveNewOrdersToOOR_(ss, ctx);

  uiToastPhase_(ss, runId, "Archiving Closed Jobs...");
  const archiveCount = executeMoveOORToArchive_(ss, ctx);

  uiToastDone_(ss, "Cleanup Complete");
  
  logToolAction_("Step 1: Cleanup", `Moved: ${moveCount}, Archived: ${archiveCount}`, ss.getName(), "Update Change Log", "INFO", runId, ctx);
  logInfo_(ctx, "FINISH", { moveCount, archiveCount });

  return `<b>Cleanup Complete</b> (Run ${runId})<br><br>` +
         `• New Orders Moved: <b>${moveCount}</b><br>` +
         `• Rows Archived: <b>${archiveCount}</b>`;
}

function executeMoveNewOrdersToOOR_(ss, parentCtx) {
  const ctx = childCtx_(parentCtx, "executeMoveNewOrdersToOOR_");
  const source = getSheetOrAlert_(ss, "New Orders");
  const target = getSheetOrAlert_(ss, "OOR");
  if (!source || !target) return 0;

  const h = getHeaders_(source);
  
  // Calculate MAX rows based on whichever is higher: Job Order or Sales Order
  let lastRow = 0;
  if (h["Job Order"] !== undefined) lastRow = Math.max(lastRow, getActualLastRow_(source, h["Job Order"] + 1));
  if (h["Sales Order"] !== undefined) lastRow = Math.max(lastRow, getActualLastRow_(source, h["Sales Order"] + 1));
  
  // Fallback if both are empty/missing
  if (lastRow <= 1) return 0;

  const vals = source.getRange(2, 1, lastRow - 1, source.getLastColumn()).getValues();
  const moveIdx = [];

  for (let i = 0; i < vals.length; i++) {
    const line = normalizeString_(vals[i][h["Line Status"]]).toLowerCase();
    const parts = normalizeString_(vals[i][h["Parts Status"]]).toLowerCase();
    if (line === "closed" || line === "invalid" || (parts !== "" && parts !== "not in wip")) {
      moveIdx.push(i + 2);
    }
  }

  if (moveIdx.length > 0) processMoveOperation_(source, target, moveIdx, "New Orders to OOR");
  return moveIdx.length;
}

function executeMoveOORToArchive_(ss, parentCtx) {
  const ctx = childCtx_(parentCtx, "executeMoveOORToArchive_");
  const archive = getSheetOrAlert_(ss, "Archive(temp)");
  if (!archive) return 0;

  let total = 0;
  ["OOR", CONFIG.STOCK_SHEET_NAME].forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;

    const h = getHeaders_(s);
    
    // Calculate MAX rows based on whichever is higher: Job Order or Sales Order
    let lastRow = 0;
    if (h["Job Order"] !== undefined) lastRow = Math.max(lastRow, getActualLastRow_(s, h["Job Order"] + 1));
    if (h["Sales Order"] !== undefined) lastRow = Math.max(lastRow, getActualLastRow_(s, h["Sales Order"] + 1));
    
    // Fallback to "Line Status" if primary columns are missing or empty
    if (lastRow <= 1 && h["Line Status"] !== undefined) {
       lastRow = getActualLastRow_(s, h["Line Status"] + 1);
    }
    
    if (lastRow <= 1) return;

    const data = s.getRange(2, 1, lastRow - 1, s.getLastColumn()).getValues();
    const rows = [];

    for (let i = 0; i < data.length; i++) {
      const st = normalizeString_(data[i][h["Line Status"]]).toLowerCase();
      if (st === "closed" || st === "invalid") rows.push(i + 2);
    }

    if (rows.length > 0) {
      processMoveOperation_(s, archive, rows, "Archive");
      total += rows.length;
    }
  });

  return total;
}

//==============================================================
// MERGED STEP 2: IMPORT & UPDATE LOGIC
//==============================================================
function showImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile("dialog")
    .setWidth(CONFIG.DIALOG_WIDTH)
    .setHeight(CONFIG.DIALOG_HEIGHT);
  SpreadsheetApp.getUi().showModalDialog(html, "Step 2: Import & Update");
}

function runFullUpdateSequence_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const runId = generateRunId_();
  const ctx = createLogCtx_(runId, "runFullUpdateSequence_", { spreadsheet: ss.getName() });

  logInfo_(ctx, "START");
  const lock = uiTryLockOrAlert_(runId);
  if (!lock) return "Error: Update already running. Try again in a moment.";

  try {
    const preflightOk = uiRequireSheets_(ss, ["ToExcel_JobOrders", "ToExcel_JobMaterialsListing", "ToExcel_PurchaseOrderListing", "ToExcel_CustomerPart"], "Missing Sheets", runId);
    if (!preflightOk) return "Error: Missing required export sheets. Please import them first.";

    uiToastPhase_(ss, runId, "Building Shortage List...");
    const shortageCount = createShortageList(childCtx_(ctx, "createShortageList"));

    uiToastPhase_(ss, runId, "Updating Tracking Sheets...");
    const results = updateOORSheetData(runId, childCtx_(ctx, "updateOORSheetData"));

    uiToastDone_(ss, "Update complete!");
    logToolAction_("Step 2: Import & Update", `Run ${runId} Complete`, ss.getName(), "Update Change Log", "INFO", runId, ctx);
    logInfo_(ctx, "FINISH", { shortageCount, results });

    return `<b>Update Complete</b> (Run ${runId})<br><br>` +
           `Shortage List: <b>${shortageCount}</b> records<br>` +
           `Changes Logged: <b>${results.changeLogCount}</b><br>` +
           `<span style="color:#5f6368;font-size:12px">Due: ${results.dueDateChanges}, Notes: ${results.noteChanges}, PC: ${results.pcFilledChanges}</span><br><br>` +
           `Unmatched: OOR (${results.unmatchedBySheet.OOR}), Stock (${results.unmatchedBySheet[CONFIG.STOCK_SHEET_NAME]})<br>` +
           `SyteLine Not Tracked: ${results.sourceOnlyTotal}`;

  } catch (e) {
    logError_(ctx, "FAILED", { message: e.message });
    return `Error: ${e.message}`;
  } finally {
    lock.releaseLock();
  }
}

function importCsvFiles(fileData, meta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const runId = generateRunId_();
  const ctx = createLogCtx_(runId, "importCsvFiles", { spreadsheet: ss.getName() });

  logInfo_(ctx, "START");

  if (CONFIG.SAFETY_ENABLED && CONFIG.SAFETY_REQUIRE_IMPORT_META) {
    const m = meta || {};
    const okAck = !!m.ack;
    const okPhrase = String(m.phrase || "").trim().toUpperCase() === String(CONFIG.SAFETY_IMPORT_PHRASE).toUpperCase();
    if (!okAck || !okPhrase) {
      logWarn_(ctx, "Import blocked by safety gate", { okAck, okPhrase });
      throw new Error("Import cancelled: safety confirmation required (checkbox + phrase).");
    }
  }

  const updatedSheets = [];
  const ignoredFiles = [];
  const warnings = [];

  for (const f in (fileData || {})) {
    const targetName = normalizeImportTargetName_(f);
    const target = ss.getSheetByName(targetName);

    logDebug_(ctx, "Map file", { file: f, targetName, exists: !!target });

    if (!target) {
      ignoredFiles.push(f);
      continue;
    }

    parseAndWriteCsvToSheet_(target, fileData[f], ctx);
    updatedSheets.push(target.getName());

    const v = validateImportedSheet_(target.getName(), target, ctx);
    if (v && v.length) warnings.push(...v.map(w => `${target.getName()}: ${w}`));
  }

  logToolAction_(
    "Import Data",
    `Run ${runId}:\n` +
      `Updated: ${updatedSheets.length ? updatedSheets.join(", ") : "(none)"}\n` +
      `Ignored: ${ignoredFiles.length ? ignoredFiles.join(", ") : "(none)"}\n` +
      `Warnings:\n${warnings.length ? warnings.join("\n") : "(none)"}`,
    ss.getName(),
    "Update Change Log",
    warnings.length ? "WARN" : "INFO",
    runId,
    ctx
  );

  logInfo_(ctx, "FINISH", { updatedSheets, ignoredFiles, warningsCount: warnings.length });
  return { updatedSheets, ignoredFiles, warnings };
}

//==============================================================
// STEP 4A - CREATE SHORTAGE LIST
//==============================================================
function createShortageList(parentCtx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = childCtx_(parentCtx || createLogCtx_(generateRunId_(), "createShortageList", { spreadsheet: ss.getName() }), "createShortageList");
  logInfo_(ctx, "START");

  const jobMat = ss.getSheetByName("ToExcel_JobMaterialsListing");
  const poListing = ss.getSheetByName("ToExcel_PurchaseOrderListing");
  const jobOrders = ss.getSheetByName("ToExcel_JobOrders");
  if (!jobMat || !poListing || !jobOrders) {
    logWarn_(ctx, "Missing required sheets", { jobMat: !!jobMat, poListing: !!poListing, jobOrders: !!jobOrders });
    return 0;
  }

  const splitJobsSet = scanForSubassemblies_(jobMat, ctx);
  const pClassMap = loadProductClassMap_(jobOrders, splitJobsSet, ctx);
  const custPoMap = loadCustomerPOMap_(jobOrders, splitJobsSet, ctx);
  const producedSet = loadProducedItemsSet_(jobOrders, ctx);
  const demands = loadJobMaterialDemands_(jobMat, splitJobsSet, pClassMap, producedSet, custPoMap, ctx);
  const supplies = loadPoSupplies_(poListing, ctx);
  const results = allocateMaterials_(demands, supplies, ctx);

  writeShortageList_(ss, results, ctx);

  logInfo_(ctx, "FINISH", { results: results.length });
  return results.length;
}

//==============================================================
// STEP 4B - UPDATE TRACKING SHEETS
//==============================================================
function updateOORSheetData(runId, parentCtx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = childCtx_(parentCtx || createLogCtx_(runId, "updateOORSheetData", { spreadsheet: ss.getName() }), "updateOORSheetData");
  logInfo_(ctx, "START");

  const jobOrders = ss.getSheetByName("ToExcel_JobOrders");
  const shortage = ss.getSheetByName("Shortage List");
  const custPart = ss.getSheetByName("ToExcel_CustomerPart");
  if (!jobOrders || !shortage || !custPart) {
    logWarn_(ctx, "Missing required sheets", { jobOrders: !!jobOrders, shortage: !!shortage, custPart: !!custPart });
    return {
      changeLogCount: 0,
      dueDateChanges: 0,
      noteChanges: 0,
      pcFilledChanges: 0,
      unmatchedBySheet: { "OOR": 0, [CONFIG.STOCK_SHEET_NAME]: 0, "New Orders": 0 },
      sourceOnlyTotal: 0,
      sourceOnlyLogged: 0
    };
  }

  const splitJobsSet = scanForSubassemblies_(jobOrders, ctx);
  const sourceData = loadSourceJobData_(jobOrders, splitJobsSet, ctx);
  const shortageData = loadShortageData_(shortage, ctx);
  const cspData = loadCustomerPartData_(custPart, splitJobsSet, ctx);

  const sheetsToProcess = ["OOR", CONFIG.STOCK_SHEET_NAME, "New Orders"];

  const summary = {
    runId,
    dueDateChanges: 0,
    noteChanges: 0,
    pcFilledChanges: 0,
    unmatchedBySheet: { "OOR": 0, [CONFIG.STOCK_SHEET_NAME]: 0, "New Orders": 0 },
    unmatchedSamples: { "OOR": [], [CONFIG.STOCK_SHEET_NAME]: [], "New Orders": [] }
  };

  const combinedLog = [];

  sheetsToProcess.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const r = processSingleReportSheet_(sh, sourceData, shortageData, cspData, runId, summary, ctx);
    combinedLog.push(...r.logs);
  });

  combinedLog.push(buildUnmatchedSummaryLog_(runId, summary));
  combinedLog.push(buildRunSummaryLog_(runId, summary));

  const audit = (CONFIG.AUDIT_SOURCE_ONLY_ENABLED === false)
    ? { total: 0, logged: 0, entries: [] }
    : auditSyteLineJobsNotTracked_(ss, sourceData, sheetsToProcess, runId, ctx);

  if (audit.entries.length) combinedLog.push(...audit.entries);

  if (combinedLog.length) writeToExternalLog_(combinedLog, "Update Change Log", ctx);

  const changeLogCount = combinedLog.filter(e =>
    e.jobOrder !== "TOOL ACTION" &&
    e.customer !== "SYTELINE NOT TRACKED" &&
    e.customer !== "RUN SUMMARY" &&
    e.customer !== "UNMATCHED SUMMARY"
  ).length;

  const out = {
    changeLogCount,
    dueDateChanges: summary.dueDateChanges,
    noteChanges: summary.noteChanges,
    pcFilledChanges: summary.pcFilledChanges,
    unmatchedBySheet: summary.unmatchedBySheet,
    sourceOnlyTotal: audit.total,
    sourceOnlyLogged: audit.logged
  };

  logInfo_(ctx, "FINISH", out);
  return out;
}

//==============================================================
// SUMMARY LOG HELPERS
//==============================================================
function buildRunSummaryLog_(runId, summary) {
  return {
    reportSheet: "Run Summary",
    projectCoordinator: "",
    jobOrder: "TOOL ACTION",
    customer: "RUN SUMMARY",
    po: "",
    qty: "",
    itemNo: "",
    severity: "INFO",
    runId,
    changes: [
      `Run ${runId} finished updating the tracking sheets.`,
      `Changes applied (by type):`,
      `• Due date updates: ${summary.dueDateChanges}`,
      `• Notes refreshed: ${summary.noteChanges}`,
      `• Project Coordinator filled: ${summary.pcFilledChanges}`
    ]
  };
}

function buildUnmatchedSummaryLog_(runId, summary) {
  const stock = CONFIG.STOCK_SHEET_NAME;

  const lines = [
    `Run ${runId}: Some rows did not match any Job Order in ToExcel_JobOrders.`,
    `OOR: ${summary.unmatchedBySheet.OOR} unmatched.`,
    `${stock}: ${summary.unmatchedBySheet[stock]} unmatched.`,
    `New Orders: ${summary.unmatchedBySheet["New Orders"]} unmatched.`
  ];

  const addSamples = (sheetName) => {
    const sample = (summary.unmatchedSamples[sheetName] || []).slice(0, CONFIG.UNMATCHED_SAMPLE_LIMIT_PER_SHEET || 10);
    if (sample.length) lines.push(`${sheetName} examples: ${sample.join(", ")}`);
  };

  addSamples("OOR");
  addSamples(stock);
  addSamples("New Orders");

  return {
    reportSheet: "Unmatched Summary",
    projectCoordinator: "",
    jobOrder: "TOOL ACTION",
    customer: "UNMATCHED SUMMARY",
    po: "",
    qty: "",
    itemNo: "",
    severity: (summary.unmatchedBySheet["New Orders"] > 0 ? "WARN" : "INFO"),
    runId,
    changes: lines
  };
}
