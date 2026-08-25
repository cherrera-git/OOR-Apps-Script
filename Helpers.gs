// ... existing code ...
const REQUIRED_IMPORT_HEADERS = {
  "ToExcel_JobMaterialsListing": ["Job", "Suffix", "Item", "Material Description", "U/M", "Qty Short", "Assigned To", "End Date", "Status", "Percent Complete"],
  "ToExcel_JobOrders": ["Job", "Job Suffix", "Due Date", "End Date", "Item", "Customer", "Status", "Assigned To", "Cust PO"],
  "ToExcel_PurchaseOrderListing": ["PO", "Item", "Ordered", "Received", "Due Date", "Status"]
};

function validateImportedSheet_(sheetName, sheet, parentCtx) {
// ... existing code ...
function loadCustomerPartData_(sheet, splitSet, parentCtx) {
  const h = getHeaders_(sheet);
  const map = new Map();
  const sufKey = getSuffixKey_(h);
  
  // Safety check: Make sure Percent Complete exists in the merged sheet
  if (!sufKey || h["Job"] === undefined || h["Percent Complete"] === undefined) return map;
  
  sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues().forEach(r => {
    const pctVal = r[h["Percent Complete"]];
    
    // Skip regular material rows that don't have CSP percentage data
    if (pctVal === "" || pctVal === undefined) return;
    
    const key = normalizeJobKey_(getCompositeJobKey_(normalizeJobKey_(r[h["Job"]]), r[h[sufKey]], splitSet.has(normalizeJobKey_(r[h["Job"]]))));
    if (!key) return;
    
    const pct = parseNumber_(pctVal || 0);
    map.set(key, pct === 100 ? "CSP received" : (pct > 0 ? "CSP partially received" : "CSP not received"));
  });
  return map;
}

function allocateMaterials_(demandsList, suppliesMap, parentCtx) {
// ... existing code ...
