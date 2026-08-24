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
  
  const editedCol = e.range.getColumn();

  // Column Indexes for "OOR" / "New Orders" tabs
  const COL_PARTS_STATUS = 12;    // Column L: Parts Status
  const COL_END_DATE_NOTES = 19;   // Column S: End Date Notes
  const COL_REAL_START_DATE = 29;  // Column AC: Real Start Date

  // Optimization: Only proceed if the edit happened in Parts Status or End Date Notes columns
  if (editedCol !== COL_PARTS_STATUS && editedCol !== COL_END_DATE_NOTES) return;

  // Optimization: Fetch all needed data in a single API call
  // We need data from column L (12) to S (19). That is 8 columns.
  const dataRange = sheet.getRange(editedRow, COL_PARTS_STATUS, 1, (COL_END_DATE_NOTES - COL_PARTS_STATUS) + 1);
  const rowData = dataRange.getValues()[0];

  const partsStatus = String(rowData[0]).trim(); // Index 0 is Column L
  const endDateNotes = String(rowData[7]).trim(); // Index 7 is Column S
  
  const realStartCell = sheet.getRange(editedRow, COL_REAL_START_DATE);
  const currentStartValue = realStartCell.getValue();

  const isPartsCleared = (partsStatus === "Picked" || partsStatus === "No Picking");
  const hasShortage = endDateNotes.includes("P-");

  if (isPartsCleared && !hasShortage) {
    // Only stamp if empty to preserve original clearance timestamp
    if (currentStartValue === "") {
      realStartCell.setValue(new Date());
      realStartCell.setNumberFormat("YYYY-MM-DD HH:mm:ss");
    }
  } else {
    // Clear timestamp if job loses cleared status or gets shortage flag
    // Optimization: Only clear if there's actually a value there
    if (currentStartValue !== "") {
      realStartCell.setValue("");
    }
  }
}
