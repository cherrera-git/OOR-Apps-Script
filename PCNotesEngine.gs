/**
 * PCNotesEngine.gs
 * - Dedicated module for handling PC Notes changelog deduplication and sorting.
 * - Centralizes the Self-Healing logic to maintain QMS data integrity and prevent cell bloat.
 */

/**
 * Builds the final PC Notes cell content by merging manual notes with the
 * deduplicated and sorted automated changelog history.
 *
 * @param {string[]} rawLines - The original lines from the current PC Notes cell.
 * @param {string[]} autoLines - The existing automated history lines (starting with '*').
 * @param {string[]} cleanNewEntries - The newly generated Syteline update lines.
 * @returns {string} The final merged string to write back to the sheet.
 */
function buildCleanPCNotes_(rawLines, autoLines, cleanNewEntries) {
  
  // 0. Fail-safes: Ensure all inputs are iterable arrays to prevent runtime errors
  rawLines = Array.isArray(rawLines) ? rawLines : [];
  autoLines = Array.isArray(autoLines) ? autoLines : [];
  cleanNewEntries = Array.isArray(cleanNewEntries) ? cleanNewEntries : [];

  // 1. Combine old history and new entries
  const allLinesRaw = [...autoLines, ...cleanNewEntries];

  // 2. NORMALIZATION: Split multi-part Syteline lines into individual atomic lines.
  const expandedLines = [];
  allLinesRaw.forEach(line => {
    const prefixMatch = line.match(/^(\*\s*(?:New Short|Shifted|Arrived|Picked Short):\s*)(.*)/i);
    if (prefixMatch) {
      const prefix = prefixMatch[1];
      let content = prefixMatch[2];
      
      // Detach any hybrid CSP tags temporarily so they aren't duplicated during splitting
      let cspSuffix = "";
      const cspSplit = content.match(/(\s*\|\s*[^|]*CSP[^|]*)/i);
      if (cspSplit) {
        cspSuffix = cspSplit[1];
        content = content.replace(cspSplit[1], "");
      }

      // Split the list by comma and rebuild into individual atomic log lines
      const parts = content.split(",").map(s => s.trim()).filter(Boolean);
      parts.forEach((p, index) => {
        // Re-attach the CSP suffix only to the first item to prevent visual duplication
        expandedLines.push(prefix + p + (index === 0 ? cspSuffix : ""));
      });
    } else {
      expandedLines.push(line);
    }
  });

  // 3. Helper function for strict extraction on normalized single-item lines
  const extractPartNumber = (line) => {
    let extractedRaw = null;
    
    // Format A: "* New Short: P-TBD (1011-349-1205)" -> Extracts what is inside ( )
    if (line.match(/^\*\s*New Short:/i)) {
      const match = line.match(/\(([^)]+)\)/); 
      if (match) extractedRaw = match[1];
    } 
    // Format B: "* Shifted: 1011-349-1205 (P-TBD→P-6/26)" or "* Arrived: AW12S"
    else if (line.match(/^\*\s*(?:Shifted|Arrived|Picked Short):/i)) {
      let content = line.replace(/^\*\s*(?:Shifted|Arrived|Picked Short):\s*/i, "");
      const match = content.match(/^([^(]+)/); 
      if (match) extractedRaw = match[1];
    }
    
    return extractedRaw ? extractedRaw.trim() : null;
  };
  
  // 4. Self-Healing Retroactive Deduplication (Process from Newest to Oldest)
  const seenItems = new Set();
  const seenMetrics = new Set();
  const healedLines = [];

  for (let i = expandedLines.length - 1; i >= 0; i--) {
    let line = expandedLines[i];
    let keepLine = true;

    // A. Independent CSP Deduplication (Scans anywhere in the line)
    if (line.match(/\bCSP\b/i)) {
      if (seenMetrics.has('csp')) {
        // Newer CSP status already captured. 
        // If this is purely a standalone CSP line, drop it completely.
        if (line.match(/^\*\s*(CSP|Waiting on CSP|Partial CSP)/i)) {
          keepLine = false;
        } else {
          // It's a hybrid line (e.g., "* New Short: ... | Waiting on CSP")
          // Strip the outdated CSP phrase but keep the line for the part number history
          line = line.replace(/\s*\|\s*[^|]*CSP[^|]*/i, "");
        }
      } else {
        seenMetrics.add('csp');
      }
    }

    if (!keepLine) continue; // Skip further processing if the line was fully dropped

    // B. Check for other top-level schedule metrics
    const metricMatch = line.match(/^\*\s*(Due date|End Date|PC):/i);
    if (metricMatch) {
      const metricType = metricMatch[1].toLowerCase();
      if (seenMetrics.has(metricType)) {
        keepLine = false; // Newer update already captured, drop legacy metric
      } else {
        seenMetrics.add(metricType);
      }
    } 
    // C. Check for granular part numbers
    else {
      const item = extractPartNumber(line);
      if (item && item.length > 1) {
        if (seenItems.has(item)) {
          keepLine = false; // This specific part has a newer update, drop legacy line
        } else {
          seenItems.add(item); // Register the newest status for this part
        }
      }
    }

    // Prepend kept lines to maintain original chronology before sorting
    if (keepLine) {
      healedLines.unshift(line); 
    }
  }

  // 5. Automated Visual Hierarchy Sorting
  healedLines.sort((a, b) => {
    const getWeight = (str) => {
      const s = str.toLowerCase();
      if (s.includes("due date")) return 1;
      if (s.includes("end date")) return 2;
      if (s.includes("pc:")) return 3;
      if (s.includes("csp")) return 4;
      if (s.includes("new short")) return 5;
      if (s.includes("shifted")) return 6;
      if (s.includes("arrived")) return 7;
      return 8; 
    };
    return getWeight(a) - getWeight(b);
  });

  let combinedChangelog = healedLines;

  // 6. Final Assembly: Separate manual notes from automated history
  const manualNotes = rawLines.filter(line => !line.trim().startsWith('*'));

  // Clean up infinite spaces and prevent duplicate blank lines from stacking
  const cleanManual = manualNotes.join('\n').trim();
  const cleanHistory = combinedChangelog.join('\n').trim();

  let finalNote = "";
  
  if (cleanManual.length > 0 && cleanHistory.length > 0) {
    // 1 newline at top, manual notes, 1 blank line gap, then history
    finalNote = "\n" + cleanManual + "\n\n" + cleanHistory;
  } else if (cleanManual.length > 0) {
    // Only manual notes: 1 newline at top, no trailing gaps
    finalNote = "\n" + cleanManual;
  } else if (cleanHistory.length > 0) {
    // Only history: 1 newline at top
    finalNote = "\n" + cleanHistory;
  }

  return finalNote;
}
