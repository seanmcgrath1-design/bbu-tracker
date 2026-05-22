// ==========================================
// EO NOTIFICATION TRACKER
// ==========================================
// Searches Gmail for "ESA EO NOTIFICATION - UNeFI Request Submitted" emails,
// parses key fields, and logs them to the "EO's" tab.

function checkEONotifications() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eoSheet = ensureEOSheet(ss);

  // Build set of already-logged EO#s to avoid duplicates
  var existingData = eoSheet.getDataRange().getValues();
  var existingEOs = {};
  for (var i = 1; i < existingData.length; i++) {
    existingEOs[String(existingData[i][1]).trim()] = true;
  }

  // Search only since last run; first run goes back 90 days
  var props = PropertiesService.getScriptProperties();
  var lastCheck = props.getProperty('lastEOCheckDate');
  var searchQuery = 'subject:"ESA EO NOTIFICATION - UNeFI Request Submitted" ' +
    (lastCheck ? 'after:' + lastCheck : 'newer_than:90d');

  var threads = GmailApp.search(searchQuery);
  var rowsAdded = 0;

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var subject = msg.getSubject();

      // Extract EO# from subject line
      var eoMatch = subject.match(/E\d{9}/);
      if (!eoMatch) return;
      var eoNumber = eoMatch[0];
      if (existingEOs[eoNumber]) return;

      var parsed = parseEONotificationEmail(msg.getPlainBody(), eoNumber);
      if (!parsed) return;

      eoSheet.appendRow([
        msg.getDate(),
        eoNumber,
        parsed.installLocation,
        parsed.installLocationDesc,
        parsed.mpn,
        parsed.requestedQty
      ]);

      existingEOs[eoNumber] = true;
      rowsAdded++;
    });
  });

  props.setProperty('lastEOCheckDate', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd'));
  console.log("EO notification check complete. Rows added: " + rowsAdded);

  try {
    var ui = SpreadsheetApp.getUi();
    if (rowsAdded > 0) {
      ui.alert("EO Notifications", rowsAdded + " new EO(s) added to the EO's tab.", ui.ButtonSet.OK);
    } else {
      ui.alert("EO Notifications", "No new EO notifications found.", ui.ButtonSet.OK);
    }
  } catch(e) {}
}

function parseEONotificationEmail(body, eoNumber) {
  var installLocation = "";
  var installLocationDesc = "";
  var mpn = "";
  var requestedQty = "";

  // --- Main table: find the data row that starts with the EO number ---
  // Columns: EO# | Type | Requisitioner | Capital WBS | Install Location | Install Location Desc | ...
  var mainRowMatch = body.match(
    new RegExp(eoNumber + '\\t[^\\t]+\\t[^\\t]+\\t[^\\t]+\\t([^\\t\\n\\r]+)\\t([^\\t\\n\\r]+)')
  );
  if (mainRowMatch) {
    installLocation     = mainRowMatch[1].trim();
    installLocationDesc = mainRowMatch[2].trim();
  }

  // --- Line items table: find the first line item row (starts with 0010) ---
  // Columns: EO Line Number | MMID ZHER | MPN | Material Description | Requested Qty
  var lineRowMatch = body.match(/0010\t[^\t]+\t([^\t\n\r]+)\t[^\t\n\r]+\t([\d.]+)/);
  if (lineRowMatch) {
    mpn          = lineRowMatch[1].trim();
    requestedQty = lineRowMatch[2].trim();
  }

  if (!installLocation && !mpn) return null;

  return {
    installLocation:     installLocation,
    installLocationDesc: installLocationDesc,
    mpn:                 mpn,
    requestedQty:        requestedQty
  };
}

// Run once to reset the date filter and re-scan the last 90 days
function resetAndRescanEONotifications() {
  PropertiesService.getScriptProperties().deleteProperty('lastEOCheckDate');
  checkEONotifications();
}

// Debug: shows the plain body of the most recent matching email
function debugEOEmailParsing() {
  var threads = GmailApp.search('subject:"ESA EO NOTIFICATION - UNeFI Request Submitted" newer_than:90d');
  if (threads.length === 0) {
    try { SpreadsheetApp.getUi().alert("EO Debug", "No matching emails found.", SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
    return;
  }

  var msg = threads[0].getMessages()[0];
  var body = msg.getPlainBody();
  var subject = msg.getSubject();
  var eoMatch = subject.match(/E\d{9}/);
  var eoNumber = eoMatch ? eoMatch[0] : "NOT FOUND";
  var parsed = eoNumber !== "NOT FOUND" ? parseEONotificationEmail(body, eoNumber) : null;

  var output = "Subject: " + subject +
    "\nEO#: " + eoNumber +
    "\n\nParsed:" +
    "\n  Install Location: "     + (parsed ? parsed.installLocation     : "—") +
    "\n  Install Location Desc: " + (parsed ? parsed.installLocationDesc : "—") +
    "\n  MPN: "                  + (parsed ? parsed.mpn                 : "—") +
    "\n  Requested Qty: "        + (parsed ? parsed.requestedQty        : "—") +
    "\n\n--- PLAIN BODY (first 800 chars) ---\n" +
    body.substring(0, 800);

  console.log(output);
  try {
    SpreadsheetApp.getUi().alert("EO Debug", output, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

function ensureEOSheet(ss) {
  var sheet = ss.getSheetByName("EO's");
  if (!sheet) {
    sheet = ss.insertSheet("EO's");
    var headers = ["Date Received", "EO#", "Install Location", "Install Location Desc", "MPN", "Requested Qty"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#4f81bd").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function setupEONotificationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkEONotifications') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkEONotifications')
    .timeBased().everyDays(1).atHour(9).create();

  console.log("EO notification trigger set for 9 AM daily.");
  try {
    SpreadsheetApp.getUi().alert("EO Notifications", "Daily trigger set for 9 AM.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}
