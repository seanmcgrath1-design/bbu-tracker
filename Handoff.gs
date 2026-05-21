// ==========================================
// AUTOMATED DAILY HANDOFF SCRIPT
// ==========================================
function generateHandoffDrafts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daily Data Dump");
  var siteDetailSheet = ss.getSheetByName("Site Detail");
  
  if (!sheet || !siteDetailSheet) {
    console.error("Missing tabs: Ensure 'Daily Data Dump' and 'Site Detail' exist.");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[1]; // Row 2
  var rows = data.slice(2);
  
  var sdData = siteDetailSheet.getDataRange().getValues();
  var sdHeaders = sdData[0];

  function getCol(h, name) {
    var searchName = String(name).trim().toLowerCase();
    for (var i = 0; i < h.length; i++) {
      if (String(h[i]).trim().toLowerCase() === searchName) return i;
    }
    return -1; 
  }
  
  // Daily Data Dump Column Mapping
  var colNodeName = getCol(headers, "Site Name");
  var colMDG = getCol(headers, "MDG Location ID");
  var colFuzeId = getCol(headers, "Fuze Project ID");
  var colHubSite = getCol(headers, "Proj Final Hub Site Name");
  var colProjType = getCol(headers, "Project Type");
  var col4GeNB = getCol(headers, "New 4G eNB"); 
  var col5GgNB = getCol(headers, "CBand gNB ID") !== -1 ? getCol(headers, "CBand gNB ID") : getCol(headers, "5G gNB ID"); 
  var colSector = getCol(headers, "Hub Scope of Work"); // Represents 6672_Sector
  var colBBU = getCol(headers, "BBU EO #");
  var colDWDM = getCol(headers, "DWDM EO");
  var colPO = getCol(headers, "Installation PO");
  var colConstStart = getCol(headers, "Construction Started (F)");
  var colSpliceTest = getCol(headers, "Splice and Test (A)"); 
  var colCQ = getCol(headers, "CQ IP Completed (A)"); 
  var colBBReady = getCol(headers, "BB Integration Ready (A)"); 
  var colGroundEO = getCol(headers, "Ground EO Released (A)");
  var colInservice = getCol(headers, "Inservice Activation (A)"); 
  
  // Site Detail Column Mapping
  var colSDFuze = getCol(sdHeaders, "Fuze Project ID");
  var colSDTech = getCol(sdHeaders, "Tech");
  var colSDSent = getCol(sdHeaders, "Handoff Sent Date");
  var colSDBuildYear = getCol(sdHeaders, "Build Plan Year"); 

  // 1. Build Tech, History & Year Map
  var techMap = {};
  var sentHistory = {};
  var yearMap = {};
  for (var i = 1; i < sdData.length; i++) {
    var fz = String(sdData[i][colSDFuze]).trim();
    if (fz) {
      techMap[fz] = String(sdData[i][colSDTech]).trim();
      sentHistory[fz] = (sdData[i][colSDSent] instanceof Date);
      yearMap[fz] = colSDBuildYear !== -1 ? String(sdData[i][colSDBuildYear]).trim() : "";
    }
  }

  // 2. Group by BBU (E0 filter)
  var bbuGroups = {};
  rows.forEach(function(row) {
    var bbuVal = String(row[colBBU] || "").trim();
    if (bbuVal.indexOf("E0") !== -1) {

      // ---> THE FIX: Clean the BBU string to force identical grouping <---
      // This splits the string at the hyphen/space and only keeps the core "E000..." number
      var cleanBbu = bbuVal.split("-")[0].trim();

      if (!bbuGroups[cleanBbu]) bbuGroups[cleanBbu] = [];
      bbuGroups[cleanBbu].push(row);
    }
  });
  
  var draftsCreated = 0;
  var today = new Date();
  var currentYearStr = String(today.getFullYear()); 

  function hasData(val) { return val !== "" && val !== null && val !== undefined && val !== "Data Missing"; }
  function getCell(row, colIndex) {
    return (colIndex === -1 || row[colIndex] === undefined || row[colIndex] === "") ? "Data Missing" : row[colIndex];
  }
  function formatDate(val) {
    if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "MM/dd/yyyy");
    return val;
  }

  // 3. Process Groups and Create Drafts
  for (var bbu in bbuGroups) {
    var rawGroupRows = bbuGroups[bbu];
    var processedGroup = [];
    var hasReadyNode = false;
    
    rawGroupRows.forEach(function(row) {
      var fz = String(row[colFuzeId]).trim();
      
      var buildYearVal = yearMap[fz] || "";
      var isCurrentYear = (buildYearVal === currentYearStr);
      
      var isInservice = (colInservice !== -1 && hasData(row[colInservice]));
      
      var isDoneInDump = (colBBReady !== -1 && hasData(row[colBBReady])) || 
                         (colGroundEO !== -1 && hasData(row[colGroundEO])) || 
                         isInservice;
      var isDoneInHistory = (sentHistory[fz] === true);
      
      var isCompleted = isDoneInDump || isDoneInHistory;
      var hasValidCQ = (row[colCQ] instanceof Date);
      
      var isReady = !isCompleted && isCurrentYear && hasValidCQ;
      if (isReady) hasReadyNode = true;
      
      processedGroup.push({ 
        data: row, 
        completed: isCompleted, 
        wrongYear: !isCurrentYear, 
        inService: isInservice, 
        ready: isReady,
        fuzeId: fz 
      });
    });
    
    if (hasReadyNode) {
      
      // Sorts the group numerically based on their "6672_Sector" (Hub Scope of Work)
      processedGroup.sort(function(a, b) {
        var sectorA = String(getCell(a.data, colSector)).trim();
        var sectorB = String(getCell(b.data, colSector)).trim();
        return sectorA.localeCompare(sectorB, undefined, {numeric: true, sensitivity: 'base'});
      });

      var activeNames = processedGroup.filter(p => p.ready).map(p => String(p.data[colNodeName])).join(" & ");
      var subject = "[ACTION REQUIRED] Ready for Install / Integration - " + activeNames;
      
      var htmlBody = "<div style='font-family: Arial, sans-serif;'>";
      htmlBody += "<div style='background-color: #fff3cd; color: #856404; padding: 10px; border: 1px solid #ffeeba; margin-bottom: 15px;'><strong>🛑 ACTION REQUIRED:</strong> Attach CQs & Paste Rack Layout.</div>";
      htmlBody += "<p>Work for <strong>" + activeNames + "</strong>.<br><span style='color: #274e13; font-size: 13px;'><em>*Green = Previously Completed Reference Sites. <br>*Grey = Outside Current Build Year.</em></span></p>";
      
      htmlBody += "<table cellpadding='5' cellspacing='0' style='border-collapse: collapse; font-size: 14px; border: 1px solid black;'>";
      
      htmlBody += "<tr><th style='background-color: #4f81bd; color: #ffffff; border: 1px solid black; text-align: left;'>Site Name</th>";
      processedGroup.forEach(function(p) { htmlBody += "<th style='background-color: cyan; color: #000000; font-weight: bold; border: 1px solid black;'>" + getCell(p.data, colNodeName) + "</th>"; });
      htmlBody += "</tr>";
      
      function buildRow(label, valFunc) {
        var tr = "<tr><td style='background-color: #4f81bd; color: #ffffff; font-weight: bold; white-space: nowrap; border: 1px solid black;'>" + label + "</td>";
        processedGroup.forEach(function(p) {
          
          var bg = "#ffffff"; // Default (Ready / White)
          if (p.inService) {
            bg = "#d9ead3"; // GREEN
          } else if (p.wrongYear) {
            bg = "#e0e0e0"; // GREY
          } else if (p.completed) {
            bg = "#d9ead3"; // GREEN
          }
          
          tr += "<td style='background-color: " + bg + "; border: 1px solid black;'>" + valFunc(p.data) + "</td>";
        });
        return tr + "</tr>";
      }
      
      htmlBody += buildRow("MDG Location ID", r => getCell(r, colMDG));
      htmlBody += "<tr><td style='background-color: #4f81bd; color: #ffffff; font-weight: bold; border: 1px solid black;'>Hub Site</td><td colspan='"+processedGroup.length+"' style='background-color:cyan; text-align:center; border: 1px solid black;'>"+getCell(processedGroup[0].data, colHubSite)+"</td></tr>";
      htmlBody += buildRow("Fuze ID", r => getCell(r, colFuzeId));
      htmlBody += buildRow("Tech", r => techMap[String(getCell(r, colFuzeId)).trim()] || "Data Missing");
      htmlBody += buildRow("Project Type", r => getCell(r, colProjType));
      htmlBody += buildRow("6672_Sector", r => getCell(r, colSector));
      htmlBody += buildRow("4G eNB ID", r => {
          var v = getCell(r, col4GeNB); var t = (techMap[String(getCell(r, colFuzeId)).trim()] || "").toUpperCase();
          return (v === "Data Missing" && t === "5G") ? "N/A" : v;
      });
      htmlBody += buildRow("5G gNB ID", r => {
          var v = getCell(r, col5GgNB); var t = (techMap[String(getCell(r, colFuzeId)).trim()] || "").toUpperCase();
          return (v === "Data Missing" && t === "4G") ? "N/A" : v;
      });
      htmlBody += "<tr><td style='background-color: #4f81bd; color: #ffffff; font-weight: bold; border: 1px solid black;'>BBU EO #</td><td colspan='"+processedGroup.length+"' style='text-align:center; border: 1px solid black;'><strong>"+bbu+"</strong></td></tr>";
      htmlBody += buildRow("DWDM EO", r => hasData(r[colDWDM]) ? r[colDWDM] : "Use Stock");
      htmlBody += buildRow("Installation PO", r => hasData(r[colPO]) ? r[colPO] : "PO Requested");
      htmlBody += buildRow("CQ Verified", r => r[colCQ] instanceof Date ? formatDate(r[colCQ]) : "CQ Pending");
      htmlBody += buildRow("Construction Started (F)", r => r[colConstStart] instanceof Date ? formatDate(r[colConstStart]) : "Data Missing");
      
      htmlBody += buildRow("Fiber FH", r => {
        var projType = String(getCell(r, colProjType)).trim();
        if (projType === "Initial Build") {
          return r[colSpliceTest] instanceof Date ? formatDate(r[colSpliceTest]) : (hasData(r[colSpliceTest]) ? r[colSpliceTest] : "Data Missing");
        }
        return "N/A";
      });

      htmlBody += "</table><br><p><strong>[PASTE RACK LAYOUT SCREENSHOT HERE]</strong></p></div>";

      GmailApp.createDraft("Ronan.Tito@parsons.com, ron.doering@ericsson.com, frantz.khan@ericsson.com, narendra.singh.bist.bist@ericsson.com, carlos.javier.rios.castrejon@ericsson.com", subject, '', { 
        htmlBody: htmlBody, 
        cc: "enis.orahovac@verizonwireless.com, matt.dubowski@verizonwireless.com, sean.mcgrath1@verizonwireless.com" 
      });

      // Site Detail Update Memory
      processedGroup.forEach(function(p) {
        if (p.ready) {
          for (var j = 1; j < sdData.length; j++) {
            if (String(sdData[j][colSDFuze]).trim() === p.fuzeId) {
              siteDetailSheet.getRange(j + 1, colSDSent + 1).setValue(today);
              break;
            }
          }
        }
      });
      draftsCreated++;
    }
  }
  
  // Logging for trigger verification
  console.log("Handoff automation finished. Drafts created: " + draftsCreated);

  // UI Alert for manual runs
  try {
    var ui = SpreadsheetApp.getUi();
    if (draftsCreated === 0) {
      ui.alert("Handoff Status", "No Sites Ready for Handoff.", ui.ButtonSet.OK);
    } else {
      ui.alert("Handoff Status", "Success! " + draftsCreated + " handoff draft(s) created.", ui.ButtonSet.OK);
    }
  } catch (e) {
    // Safely catches the error and does nothing if the script is running automatically
  }
}