/**
 * GOOGLE APPS SCRIPT: RF CQ Data Sync
 * Author: Google Apps Script Automation Specialist
 * * Description:
 * Matches "FUZE Project ID" from "Daily Data Dump" to "Site Detail".
 * Pulls "CQ RF Completed (F)" -> "RF CQ  forecast"
 * Pulls "CQ RF Completed (A)" -> "RF CQ completed"
 * Includes a function to set a daily automated trigger.
 */

const APP_CONFIG = {
  MASTER_SHEET: 'Site Detail',
  DUMP_SHEET: 'Daily Data Dump',
  
  // Unique Identifier used in both sheets to match rows
  ID_HEADER_VARIATIONS: ['FUZE Project ID', 'Fuze Project ID', 'FUZE', 'Site ID'], 
  
  // Mapping of Source (Dump) headers to Target (Main Tab) headers
  MAPPING: [
    { source: 'CQ RF Completed (F)', target: 'RF CQ  forecast' },
    { source: 'CQ RF Completed (A)', target: 'RF CQ completed' }
  ]
};

/**
 * Run this function ONCE to schedule the daily 8 AM refresh.
 */
function createDailyTrigger() {
  const functionName = 'dailyAutomationSync';
  
  // First, remove any existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new trigger for 8 AM daily
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
    
  console.log("Success: Daily 8 AM trigger has been scheduled.");
}

function dailyAutomationSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(APP_CONFIG.MASTER_SHEET);
  const dumpSheet = ss.getSheetByName(APP_CONFIG.DUMP_SHEET);
  
  if (!masterSheet || !dumpSheet) {
    throw new Error(`Sheet tabs not found. Ensure tabs are named '${APP_CONFIG.MASTER_SHEET}' and '${APP_CONFIG.DUMP_SHEET}'.`);
  }

  const masterData = masterSheet.getDataRange().getValues();
  const dumpDataArr = dumpSheet.getDataRange().getValues();

  /**
   * Helper to find which row contains the headers and return metadata
   */
  const locateHeaders = (dataRows, searchArray) => {
    const normalizedSearches = searchArray.map(s => s.toString().toLowerCase().replace(/\s+/g, ' ').trim());
    
    for (let r = 0; r < Math.min(dataRows.length, 15); r++) {
      const row = dataRows[r];
      const idIdx = row.findIndex(cell => {
        if (!cell) return false;
        const normalizedCell = cell.toString().toLowerCase().replace(/\s+/g, ' ').trim();
        return normalizedSearches.includes(normalizedCell);
      });
      
      if (idIdx !== -1) return { rowIdx: r, idColIdx: idIdx, headers: row };
    }
    return null;
  };

  // 1. Locate headers in both sheets
  const masterHeaderInfo = locateHeaders(masterData, APP_CONFIG.ID_HEADER_VARIATIONS);
  const dumpHeaderInfo = locateHeaders(dumpDataArr, APP_CONFIG.ID_HEADER_VARIATIONS);

  if (!masterHeaderInfo) throw new Error("Could not find ID column in Site Detail. Check for 'Fuze Project ID'.");
  if (!dumpHeaderInfo) throw new Error("Could not find ID column in Daily Data Dump. Check for 'FUZE Project ID'.");

  /**
   * Helper to find specific column indices in the identified header row
   */
  const getColIndex = (headerRow, name) => {
    const normalizedSearch = name.toString().toLowerCase().replace(/\s+/g, ' ').trim();
    return headerRow.findIndex(h => {
      if (!h) return false;
      return h.toString().toLowerCase().replace(/\s+/g, ' ').trim() === normalizedSearch;
    });
  };

  // 2. Identify mapping indices
  const activeMappings = APP_CONFIG.MAPPING.map(m => ({
    sourceIdx: getColIndex(dumpHeaderInfo.headers, m.source),
    targetIdx: getColIndex(masterHeaderInfo.headers, m.target),
    name: m.target
  })).filter(m => m.sourceIdx !== -1 && m.targetIdx !== -1);

  if (activeMappings.length === 0) {
    throw new Error("Target columns (RF CQ forecast/completed) not found in Site Detail. Please check headers.");
  }

  // 3. Map the Dump Data (Skip leading rows and header row)
  const dumpLookup = new Map();
  for (let i = dumpHeaderInfo.rowIdx + 1; i < dumpDataArr.length; i++) {
    const row = dumpDataArr[i];
    const fuzeId = row[dumpHeaderInfo.idColIdx] ? row[dumpHeaderInfo.idColIdx].toString().trim() : null;
    if (fuzeId) {
      dumpLookup.set(fuzeId, row);
    }
  }

  // 4. Update the Master Data array in memory (We will only write back the specific columns)
  // Create arrays for each target column to perform bulk updates
  const updateArrays = activeMappings.map(m => ({
    targetIdx: m.targetIdx,
    values: []
  }));

  let updateCount = 0;
  // Start from row after header
  for (let j = masterHeaderInfo.rowIdx + 1; j < masterData.length; j++) {
    const masterFuzeId = masterData[j][masterHeaderInfo.idColIdx] ? masterData[j][masterHeaderInfo.idColIdx].toString().trim() : null;
    const matchedDumpRow = dumpLookup.get(masterFuzeId);

    activeMappings.forEach((map, index) => {
      let finalValue = masterData[j][map.targetIdx]; // Default to current value
      
      if (matchedDumpRow) {
        const newValue = matchedDumpRow[map.sourceIdx];
        if (newValue !== "" && newValue !== null && newValue !== undefined) {
          finalValue = newValue;
        }
      }
      updateArrays[index].values.push([finalValue]);
    });
    
    if (matchedDumpRow) updateCount++;
  }

  // 5. Write back ONLY the specific mapped columns to avoid data validation errors in other columns
  activeMappings.forEach((map, index) => {
    const colValues = updateArrays[index].values;
    if (colValues.length > 0) {
      masterSheet.getRange(masterHeaderInfo.rowIdx + 2, map.targetIdx + 1, colValues.length, 1)
                 .setValues(colValues);
    }
  });
  
  console.log(`Sync complete. Found ${dumpLookup.size} source entries and updated ${updateCount} matches in Site Detail.`);
}