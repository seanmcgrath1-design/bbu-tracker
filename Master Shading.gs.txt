function masterSiteUpdateAndHighlight() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var siteDetailSheet = ss.getSheetByName("Site Detail");
  var dataDumpSheet = ss.getSheetByName("Daily Data Dump");
  
  if (!siteDetailSheet || !dataDumpSheet) {
    SpreadsheetApp.getUi().alert("Sheets not found. Please verify 'Site Detail' and 'Daily Data Dump'.");
    return;
  }

  // ==========================================
  // 1. GET DUMP SHEET COLUMNS (Headers on Row 2)
  // ==========================================
  var dumpLastRow = dataDumpSheet.getLastRow();
  var dumpLastCol = dataDumpSheet.getLastColumn();
  if (dumpLastRow < 3) return; 
  
  var dumpData = dataDumpSheet.getRange(1, 1, dumpLastRow, dumpLastCol).getValues();
  var dumpHeaders = dumpData[1]; // Row 2
  
  var dumpFuzeCol = -1, dumpActivationCol = -1, dumpIntegrationCol = -1, dumpBbReadyCol = -1;
  
  for (var h = 0; h < dumpHeaders.length; h++) {
    var cleanHeader = dumpHeaders[h].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (cleanHeader === "fuzeprojectid") dumpFuzeCol = h;
    if (cleanHeader === "inserviceactivationa") dumpActivationCol = h;
    if (cleanHeader === "bbintegrationcompleteda") dumpIntegrationCol = h;
    if (cleanHeader === "bbintegrationreadya") dumpBbReadyCol = h; // <-- UPDATED HERE
  }

  if (dumpFuzeCol === -1 || dumpActivationCol === -1 || dumpIntegrationCol === -1 || dumpBbReadyCol === -1) {
    SpreadsheetApp.getUi().alert("Error: Missing one of the required target columns in the Daily Data Dump.");
    return;
  }

  // ==========================================
  // 2. BUILD MASTER DICTIONARY
  // ==========================================
  var projectMap = {};
  
  // Start reading data on Row 3
  for (var r = 2; r < dumpData.length; r++) {
    var fuzeId = dumpData[r][dumpFuzeCol].toString().replace(/[^a-zA-Z0-9]/g, '');
    if (!fuzeId) continue;

    var actVal = dumpData[r][dumpActivationCol];
    var intVal = dumpData[r][dumpIntegrationCol];
    var bbReadyVal = dumpData[r][dumpBbReadyCol]; // <-- UPDATED HERE

    var mappedData = { activationDate: "", isIntegrated: false, isBbReady: false };

    // Check Priority 1: Activation
    if (actVal !== "" && actVal != null) {
      if (actVal instanceof Date) {
        mappedData.activationDate = Utilities.formatDate(actVal, Session.getScriptTimeZone(), "MM-dd-yyyy");
      } else {
        var parsedDate = new Date(actVal);
        if (!isNaN(parsedDate.getTime())) {
          mappedData.activationDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
        } else {
          mappedData.activationDate = actVal.toString().trim();
        }
      }
    }

    // Check Priority 2: Integration
    if (intVal instanceof Date && intVal.getFullYear() > 2000) {
      mappedData.isIntegrated = true;
    }

    // Check Priority 3: BB Integration Ready (formerly CQ Completed)
    if (bbReadyVal !== "" && bbReadyVal != null) {
      mappedData.isBbReady = true; // <-- UPDATED HERE
    }

    projectMap[fuzeId] = mappedData;
  }

  // ==========================================
  // 3. GET SITE DETAIL COLUMNS (Headers on Row 1)
  // ==========================================
  var siteData = siteDetailSheet.getDataRange().getValues();
  var siteHeaders = siteData[0]; 
  
  var siteFuzeCol = -1, nodeActiveCol = -1, siteNameCol = -1; 
  
  for (var i = 0; i < siteHeaders.length; i++) {
    var cleanSiteHeader = siteHeaders[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (cleanSiteHeader === "fuzeprojectid") siteFuzeCol = i;
    if (cleanSiteHeader === "nodeactivated") nodeActiveCol = i;
    if (cleanSiteHeader === "sitename") siteNameCol = i; 
  }
  
  if (siteFuzeCol === -1 || nodeActiveCol === -1 || siteNameCol === -1) {
    SpreadsheetApp.getUi().alert("Missing a required column in the Site Detail tab.");
    return;
  }

  // ==========================================
  // 4. APPLY PRIORITIES & PREPARE OUTPUT
  // ==========================================
  var dateOutputValues = [];
  
  // Grab existing backgrounds so we only update what needs updating
  var siteNameRange = siteDetailSheet.getRange(2, siteNameCol + 1, siteData.length - 1, 1);
  var backgrounds = siteNameRange.getBackgrounds();
  var changesMade = false;
  
  var actColor = "#36ab1f"; // Priority 1
  var intColor = "#93c47d"; // Priority 2
  var bbReadyColor  = "#f4cccc"; // Priority 3 (Light Red 3)

  // Loop through Site Detail (Row 2 onward)
  for (var j = 1; j < siteData.length; j++) {
    var siteFuzeId = siteData[j][siteFuzeCol].toString().replace(/[^a-zA-Z0-9]/g, '');
    var currentBg = backgrounds[j-1][0]; // j-1 because backgrounds array starts at 0 for row 2
    
    if (siteFuzeId !== "" && projectMap.hasOwnProperty(siteFuzeId)) {
      var projInfo = projectMap[siteFuzeId];
      
      // -- Determine Date Column Output --
      if (projInfo.activationDate) {
        dateOutputValues.push([projInfo.activationDate]);
      } else {
        dateOutputValues.push([""]); // Clear it if there is no activation date
      }

      // -- Determine Color Priority --
      var newBg = currentBg;
      
      if (projInfo.activationDate) {
        newBg = actColor;
      } else if (projInfo.isIntegrated) {
        newBg = intColor;
      } else if (projInfo.isBbReady) {
        newBg = bbReadyColor; // <-- UPDATED HERE
      }

      // Apply color change if needed
      if (newBg !== currentBg) {
        backgrounds[j-1][0] = newBg;
        changesMade = true;
      }
      
    } else {
      // No match found in dump
      dateOutputValues.push([""]); 
    }
  }

  // ==========================================
  // 5. WRITE DATA TO SHEET
  // ==========================================
  if (dateOutputValues.length > 0) {
    siteDetailSheet.getRange(2, nodeActiveCol + 1, dateOutputValues.length, 1).setValues(dateOutputValues);
  }
  
  if (changesMade) {
    siteNameRange.setBackgrounds(backgrounds);
  }
}