/**
 * AUTOMATION: Sync Tech/BB Config from Daily Dump to Site Detail
 * DESCRIPTION:
 * Matches FUZE IDs between 'Site Detail' and 'Daily Data Dump'.
 * If 'Tech' is blank OR contains "Not found" in Site Detail, it pulls 
 * 'BB Config' from the dump. 
 * * Simple Formatting:
 * If it sees "4G/5G" anywhere in the cell, it outputs "4G/5G".
 * If it sees "5G" anywhere in the cell, it outputs "5G".
 */

function updateTechFromDump() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const siteSheet = ss.getSheetByName("Site Detail");
  const dumpSheet = ss.getSheetByName("Daily Data Dump");
  
  if (!siteSheet || !dumpSheet) {
    SpreadsheetApp.getUi().alert("Error: Ensure tabs 'Site Detail' and 'Daily Data Dump' exist.");
    return;
  }
  
  // --- CONFIGURATION ---
  const SITE_FUZE_IDX = 0;   // Column A
  const SITE_TECH_IDX = 15;  // Column P
  const DUMP_FUZE_IDX = 0;   // Column A
  const DUMP_BB_IDX = 38;    // Column AM
  // ---------------------
  
  const siteData = siteSheet.getDataRange().getValues();
  const dumpData = dumpSheet.getDataRange().getValues();
  
  // 1. Build lookup map from Dump
  const dumpMap = new Map();
  for (let i = 2; i < dumpData.length; i++) {
    let fuzeId = dumpData[i][DUMP_FUZE_IDX];
    let bbConfig = dumpData[i][DUMP_BB_IDX];
    if (fuzeId) {
      dumpMap.set(fuzeId.toString().trim(), bbConfig ? bbConfig.toString().trim() : "");
    }
  }
  
  // 2. Prepare update array
  const techRange = siteSheet.getRange(1, SITE_TECH_IDX + 1, siteData.length, 1);
  const techValues = techRange.getValues();
  let changeCounter = 0;
  
  // 3. Process Site Detail
  for (let j = 1; j < siteData.length; j++) {
    let currentFuzeId = siteData[j][SITE_FUZE_IDX];
    let currentTechVal = techValues[j][0] ? techValues[j][0].toString().trim() : "";
    
    // Update if blank or "Not found"
    if (currentTechVal === "" || currentTechVal.toLowerCase() === "not found") {
      if (currentFuzeId) {
        let cleanId = currentFuzeId.toString().trim();
        
        if (dumpMap.has(cleanId)) {
          let rawVal = dumpMap.get(cleanId);
          let finalVal = "";

          // SIMPLEST LOGIC POSSIBLE
          if (rawVal.includes("4G/5G")) {
            finalVal = "4G/5G";
          } else if (rawVal.includes("5G")) {
            finalVal = "5G";
          } else if (rawVal !== "") {
            finalVal = rawVal; // Keep it as is if no specific match
          } else {
            finalVal = "Not found";
          }

          techValues[j][0] = finalVal;
          changeCounter++;
        }
      }
    }
  }
  
  // 4. Save
  if (changeCounter > 0) {
    techRange.setValues(techValues);
    SpreadsheetApp.getUi().alert(`Updated ${changeCounter} rows.`);
  } else {
    SpreadsheetApp.getUi().alert("No new data to update.");
  }
}