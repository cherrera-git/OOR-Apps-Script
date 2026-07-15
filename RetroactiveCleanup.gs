/**
 * RetroactiveCleanup.gs
 * - One-time macro to instantly execute the Self-Healing deduplication 
 * across all existing PC Notes in the Open Order Report.
 * - Supports QMS Data Integrity by removing legacy cell bloat immediately.
 */

function forceCleanLegacyPCNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Target the active tracking sheet
  const sheet = ss.getSheetByName("Open Order Report") || ss.getActiveSheet(); 
  
  if (!sheet) {
    Logger.log("Target sheet not found.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 1. Locate the PC Notes column dynamically
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const notesColIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes("pc notes"));

  if (notesColIdx === -1) {
    Logger.log("PC Notes column not found.");
    return;
  }

  // 2. Fetch all current PC Notes
  const range = sheet.getRange(2, notesColIdx + 1, lastRow - 1, 1);
  const notesData = range.getValues();

  // 3. Process each cell through the Self-Healing engine
  const cleanedData = notesData.map(row => {
    const cellText = row[0];
    if (!cellText || typeof cellText !== 'string') return [cellText];

    const rawLines = cellText.split('\n');
    const autoLines = rawLines.filter(line => line.trim().startsWith('*'));

    // By passing an empty array [] as the third parameter, the engine 
    // forces the existing autoLines to self-heal and deduplicate against themselves.
    const healedText = buildCleanPCNotes_(rawLines, autoLines, []);
    
    return [healedText];
  });

  // 4. Write the cleaned, deduplicated notes back to the sheet in a single batch operation
  range.setValues(cleanedData);
  Logger.log("Retroactive cleanup complete.");
}
