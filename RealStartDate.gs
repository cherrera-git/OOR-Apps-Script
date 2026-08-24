/**
 * Automatically stamps Real Start Date when a job is 100% material cleared.
 * Works on both "OOR" and "New Orders" sheets.
 * File Name in Apps Script: RealStartDate.gs
 * Trigger: Simple or Installed onEdit trigger.
 */
function handleMaterialClearance(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Monitor both "OOR" and "New Orders" tabs
  if (sheetName !== "OOR" && sheetName !== "New Orders") return;

  const editedRow = e.range.getRow();
  if (editedRow < 2) return; // Skip header row

  // Column Indexes for "OOR" / "New Orders" tabs
  const COL_PARTS_STATUS = 12;    // Column L: Parts Status
  const COL_END_DATE_NOTES = 19;   // Column S: End Date Notes
  const COL_REAL_START_DATE = 29;  // Column AC: Real Start Date

  const partsStatus = sheet.getRange(editedRow, COL_PARTS_STATUS).getValue().toString().trim();
  const endDateNotes = sheet.getRange(editedRow, COL_END_DATE_NOTES).getValue().toString().trim();
  const realStartCell = sheet.getRange(editedRow, COL_REAL_START_DATE);

  const isPartsCleared = (partsStatus === "Picked" || partsStatus === "No Picking");
  const hasShortage = endDateNotes.includes("P-");

  if (isPartsCleared && !hasShortage) {
    // Only stamp if empty to preserve original clearance timestamp
    if (realStartCell.getValue() === "") {
      realStartCell.setValue(new Date());
      realStartCell.setNumberFormat("YYYY-MM-DD HH:mm:ss");
    }
  } else {
    // Clear timestamp if job loses cleared status or gets shortage flag
    realStartCell.setValue("");
  }
}
