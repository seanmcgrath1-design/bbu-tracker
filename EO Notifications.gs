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
  var searchQuery = 'from:donotreply@verizon.com subject:"EO NOTIFICATION" ' +
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

  // Plain body separates table cells with spaces/newlines (not tabs).
  // Data row pattern: ...VZ-XXXXXXXX.X.XXXX [Install Location] [Install Location Desc] [Amount.00]...

  // Install Location (10-digit number starting with 5) and Install Location Desc
  // Anchor on Capital WBS (VZ-XXXXXXXX.X.XXXX) since it uniquely precedes these fields
  var locMatch = body.match(/VZ-[\d]+\.[A-Z]+\.[\d]+\s+(5\d{9})\s+([^\n]+?)\s+[\d,]{4,}\.\d{2}/);
  if (locMatch) {
    installLocation     = locMatch[1].trim();
    installLocationDesc = locMatch[2].trim();
  }

  // MPN: alphanumeric part number format (e.g. RDH102409/1-PLV or RDH102409/1)
  var mpnMatch = body.match(/\b([A-Z][A-Z0-9]{3,}\/\d+(?:-[A-Z0-9-]+)?)\b/);
  if (mpnMatch) mpn = mpnMatch[1].trim();

  // Requested Qty: number with exactly 3 decimal places (e.g. 50.000)
  var qtyMatch = body.match(/\b(\d+\.\d{3})\b/);
  if (qtyMatch) requestedQty = qtyMatch[1].trim();

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
  var results = [];

  // Try progressively broader searches to find the email
  var searches = [
    'subject:"ESA EO NOTIFICATION - UNeFI Request Submitted" newer_than:90d',
    'subject:"ESA EO NOTIFICATION" newer_than:90d',
    'from:donotreply@verizon.com subject:"EO NOTIFICATION" newer_than:90d',
    'from:donotreply@verizon.com "UNeFI Request Submitted" newer_than:90d',
    'from:donotreply@verizon.com newer_than:7d'
  ];

  var foundMsg = null;
  for (var s = 0; s < searches.length; s++) {
    var threads = GmailApp.search(searches[s]);
    results.push('Search ' + (s+1) + ': [' + searches[s] + '] → ' + threads.length + ' thread(s)');
    if (threads.length > 0 && !foundMsg) {
      foundMsg = threads[0].getMessages()[0];
      results.push('  ✅ Found: ' + foundMsg.getSubject());
    }
  }

  if (!foundMsg) {
    results.push("\n❌ No matching emails found with any search.");
  } else {
    var body = foundMsg.getPlainBody();
    var subject = foundMsg.getSubject();
    var eoMatch = subject.match(/E\d{9}/);
    var eoNumber = eoMatch ? eoMatch[0] : "NOT FOUND";
    var parsed = eoNumber !== "NOT FOUND" ? parseEONotificationEmail(body, eoNumber) : null;

    results.push("\nEO#: " + eoNumber);
    results.push("Parsed:");
    results.push("  Install Location: "      + (parsed ? parsed.installLocation     : "❌ not found"));
    results.push("  Install Location Desc: " + (parsed ? parsed.installLocationDesc : "❌ not found"));
    results.push("  MPN: "                   + (parsed ? parsed.mpn                 : "❌ not found"));
    results.push("  Requested Qty: "         + (parsed ? parsed.requestedQty        : "❌ not found"));
    results.push("\n--- PLAIN BODY (first 800 chars) ---\n" + body.substring(0, 800));
  }

  var output = results.join("\n");
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
