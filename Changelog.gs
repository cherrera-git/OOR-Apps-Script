/**
 * Appends a concise changelog to Column U ("PC Notes") whenever a row is edited.
 * If you already have an onEdit(e) function in Code.gs, just call appendChangelogToPCNotes(e) inside it.
 */
function onEdit(e) {
  appendChangelogToPCNotes(e);
}

function appendChangelogToPCNotes(e) {
  // Ensure the event object and range exist
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  
  // Optional: Restrict to a specific sheet name
  // if (sheet.getName() !== "Tracking Sheet") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const pcNotesColIndex = 21; // Column U is the 21st column

  // Skip header row, multi-cell edits (like copy-pasting ranges), or edits directly to the PC Notes column
  if (row === 1 || col === pcNotesColIndex || e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
    return;
  }

  // Get the header name of the edited column to provide context
  const header = sheet.getRange(1, col).getValue() || `Col ${col}`;
  const newValue = e.value !== undefined ? e.value : "Cleared";
  const oldValue = e.oldValue !== undefined ? e.oldValue : "Blank";

  // Create a concise log matching the requested sample format
  const newLog = `* ${header}: ${oldValue} → ${newValue}`;

  // Fetch current PC Notes
  const notesRange = sheet.getRange(row, pcNotesColIndex);
  const currentNotes = notesRange.getValue().toString();

  if (currentNotes) {
    // Split existing notes into lines
    let lines = currentNotes.split('\n');
    
    // Filter out any older changelogs (lines starting with "* ")
    lines = lines.filter(line => !line.startsWith('* '));
    
    // Append the new changelog to the end
    lines.push(newLog);
    
    // Update the cell
    notesRange.setValue(lines.join('\n'));
  } else {
    notesRange.setValue(newLog);
  }
}
