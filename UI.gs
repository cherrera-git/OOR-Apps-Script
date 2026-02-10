/**
 * UI.gs
 * - UI helper functions only (no HTML content in this file).
 * - Do NOT declare CONFIG here.
 * - Do NOT declare __HEADERS_CACHE here.
 */

function uiAlert_(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function uiConfirm_(title, message) {
  const ui = SpreadsheetApp.getUi();
  return ui.alert(title, message, ui.ButtonSet.YES_NO) === ui.Button.YES;
}

function uiPrompt_(title, message, defaultText) {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  return String(resp.getResponseText() || defaultText || "");
}

/**
 * Safety gate pattern:
 * - optional 1st confirm
 * - optional 2nd confirm
 * - required phrase input (exact match)
 */
function uiSafetyGate_(opts) {
  const ui = SpreadsheetApp.getUi();

  if (opts && opts.title && opts.message) {
    const r1 = ui.alert(opts.title, opts.message, ui.ButtonSet.YES_NO);
    if (r1 !== ui.Button.YES) return false;
  }

  if (opts && opts.confirmTwice) {
    const r2 = ui.alert(opts.confirm2Title || "Confirm", opts.confirm2Message || "Proceed?", ui.ButtonSet.YES_NO);
    if (r2 !== ui.Button.YES) return false;
  }

  if (opts && opts.phrase) {
    const pr = ui.prompt(
      opts.phraseTitle || "Type-to-confirm",
      (opts.phraseMessage || "Type the phrase to continue.") + `\n\nRequired phrase: ${opts.phrase}`,
      ui.ButtonSet.OK_CANCEL
    );
    if (pr.getSelectedButton() !== ui.Button.OK) return false;

    const typed = String(pr.getResponseText() || "");
    return typed === String(opts.phrase);
  }

  return true;
}

/**
 * Toast helpers: consistent usage for long operations.
 */
function uiToastPhase_(ss, runId, message) {
  try {
    ss.toast(`Run ${runId}: ${message}`, "Update in Progress", -1);
  } catch (_) {}
}

function uiToastDone_(ss, message) {
  try {
    ss.toast(message || "Done", "Success", 5);
  } catch (_) {}
}

/**
 * Required sheets gate
 */
function uiRequireSheets_(ss, sheetNames, title, runId) {
  const missing = [];
  (sheetNames || []).forEach(n => { if (!ss.getSheetByName(n)) missing.push(n); });

  if (missing.length) {
    uiAlert_(
      title || "Missing sheets",
      `Run ID: ${runId || ""}\n\nMissing required sheet(s):\n• ${missing.join("\n• ")}`
    );
    return false;
  }
  return true;
}

/**
 * Prevent double-runs: acquire a lock with user feedback.
 */
function uiTryLockOrAlert_(runId) {
  const lock = LockService.getDocumentLock();
  const ok = lock.tryLock(30000);
  if (!ok) {
    uiAlert_("Update already running", `Run ID: ${runId}\n\nAnother update is currently running. Try again in a moment.`);
    return null;
  }
  return lock;
}
