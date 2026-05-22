// ==========================================
// EO NOTIFICATION TRACKER
// ==========================================
// Columns: A=Date Received, B=EO#, C=Receiving EO#, D=Install Location,
//          E=Install Location Desc, F=MPN, G=Requested Qty

function getEOSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet() ||
         SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
}

function checkEONotifications() {
  var ss = getEOSpreadsheet();
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
  var searchQuery = 'from:donotreply@verizon.com subject:"unefi request" ' +
    (lastCheck ? 'after:' + lastCheck : 'newer_than:180d');

  var threads = GmailApp.search(searchQuery);
  var rowsAdded = 0;

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var subject = msg.getSubject();

      // Only process UNeFI notification emails (Submitted or Approved); skip P2P, etc.
      var isUnefi = subject.indexOf('UNeFI Request') !== -1 ||
                    subject.indexOf('UNeFi Request') !== -1;
      if (!isUnefi) return;

      var eoMatch = subject.match(/E\d{9}/);
      if (!eoMatch) return;
      var eoNumber = eoMatch[0];
      if (existingEOs[eoNumber]) return;

      var parsed = parseEONotificationEmail(msg.getPlainBody(), eoNumber);
      if (!parsed) return;

      eoSheet.appendRow([
        msg.getDate(),
        eoNumber,
        "",                        // Col C: Receiving EO# (filled by checkP2PTransfers)
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

// Searches for P2P Transfer emails and updates col C (Receiving EO#)
// for any row whose EO# matches the Donating EO in a transfer.
function checkP2PTransfers() {
  var ss = getEOSpreadsheet();
  var eoSheet = ss.getSheetByName("EO's");
  if (!eoSheet) return;

  var eoData = eoSheet.getDataRange().getValues();

  // Build map: Donating EO# → row index (1-based offset handled on write)
  var eoRowMap = {};
  for (var i = 1; i < eoData.length; i++) {
    var eoNum = String(eoData[i][1]).trim();
    if (eoNum) eoRowMap[eoNum] = i;
  }

  var props = PropertiesService.getScriptProperties();
  var lastCheck = props.getProperty('lastP2PCheckDate');
  var searchQuery = 'from:donotreply@verizon.com subject:"Project to Project Transfer" ' +
    (lastCheck ? 'after:' + lastCheck : 'newer_than:90d');

  var threads = GmailApp.search(searchQuery);
  var updatesApplied = 0;

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var body = msg.getPlainBody();

      // Data row: [8-digit Donating SPM] [Donating E#] [long Receiving SPM] [Receiving E#]
      var allEOs = body.match(/E\d{9}/g) || [];
      var uniqueEOs = [];
      allEOs.forEach(function(e) { if (uniqueEOs.indexOf(e) === -1) uniqueEOs.push(e); });
      var dataMatch = uniqueEOs.length >= 2 ? [null, uniqueEOs[0], uniqueEOs[1]] : null;
      if (!dataMatch) return;

      var donatingEO  = dataMatch[1];
      var receivingEO = dataMatch[2];

      if (!eoRowMap.hasOwnProperty(donatingEO)) return;

      var rowIdx = eoRowMap[donatingEO];
      var currentReceiving = String(eoData[rowIdx][2]).trim(); // Col C
      if (currentReceiving) return; // already recorded

      eoSheet.getRange(rowIdx + 1, 3).setValue(receivingEO);
      eoData[rowIdx][2] = receivingEO;
      updatesApplied++;
    });
  });

  props.setProperty('lastP2PCheckDate', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd'));
  console.log("P2P Transfer check complete. Updates applied: " + updatesApplied);

  try {
    var ui = SpreadsheetApp.getUi();
    if (updatesApplied > 0) {
      ui.alert("P2P Transfers", updatesApplied + " row(s) updated with Receiving EO#.", ui.ButtonSet.OK);
    } else {
      ui.alert("P2P Transfers", "No new P2P transfers matched existing EOs.", ui.ButtonSet.OK);
    }
  } catch(e) {}
}

function parseEONotificationEmail(body, eoNumber) {
  var installLocation = "";
  var installLocationDesc = "";
  var mpn = "";
  var requestedQty = "";

  // Plain body separates table cells with spaces/newlines (not tabs).
  // Anchor on Capital WBS (VZ-XXXXXXXX.X.XXXX) which uniquely precedes Install Location fields.
  var wbsMatch = body.match(/VZ-[\d]+\.[A-Z]+\.[\d]+\s+(5\d{9})([ \t]+([^\n\d][^\n\d]*?))?[\t ]*(?:\n|[\d,]+\.\d{2})/);
  if (wbsMatch) {
    installLocation     = wbsMatch[1].trim();
    installLocationDesc = wbsMatch[3] ? wbsMatch[3].trim() : "";
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

// Run once to reset date filters and do a full chunked backfill + P2P scan
function resetAndRescanEONotifications() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('lastEOCheckDate');
  props.deleteProperty('lastP2PCheckDate');
  backfillEONotificationsChunked();
  checkP2PTransfers();
}

// Lighter reset — only re-scans P2P transfers without re-running the full EO backfill
function resetAndRescanP2POnly() {
  PropertiesService.getScriptProperties().deleteProperty('lastP2PCheckDate');
  checkP2PTransfers();
}

// Searches in 60-day chunks going back 180 days to stay under Gmail's 500-result cap
function backfillEONotificationsChunked() {
  var ss = getEOSpreadsheet();
  var eoSheet = ensureEOSheet(ss);

  var existingData = eoSheet.getDataRange().getValues();
  var existingEOs = {};
  for (var i = 1; i < existingData.length; i++) {
    existingEOs[String(existingData[i][1]).trim()] = true;
  }

  var today = new Date();
  var tz = Session.getScriptTimeZone();
  var rowsAdded = 0;
  var chunkDays = 60;
  var maxDays = 180;

  for (var start = 0; start < maxDays; start += chunkDays) {
    var dateFrom = new Date(today.getTime() - (start + chunkDays) * 86400000);
    var dateTo   = new Date(today.getTime() - start * 86400000);
    var fromStr  = Utilities.formatDate(dateFrom, tz, 'yyyy/MM/dd');
    var toStr    = Utilities.formatDate(dateTo,   tz, 'yyyy/MM/dd');

    var query = 'from:donotreply@verizon.com subject:"unefi request" after:' + fromStr + ' before:' + toStr;
    var threads = GmailApp.search(query);
    console.log('Chunk ' + fromStr + ' → ' + toStr + ': ' + threads.length + ' threads');

    threads.forEach(function(thread) {
      thread.getMessages().forEach(function(msg) {
        var subject = msg.getSubject();
        var isUnefi = subject.indexOf('UNeFI Request') !== -1 ||
                      subject.indexOf('UNeFi Request') !== -1;
        if (!isUnefi) return;

        var eoMatch = subject.match(/E\d{9}/);
        if (!eoMatch) return;
        var eoNumber = eoMatch[0];
        if (existingEOs[eoNumber]) return;

        var parsed = parseEONotificationEmail(msg.getPlainBody(), eoNumber);
        if (!parsed) return;

        eoSheet.appendRow([
          msg.getDate(),
          eoNumber,
          "",
          parsed.installLocation,
          parsed.installLocationDesc,
          parsed.mpn,
          parsed.requestedQty
        ]);

        existingEOs[eoNumber] = true;
        rowsAdded++;
      });
    });
  }

  // Set lastEOCheckDate so the daily trigger picks up from today
  PropertiesService.getScriptProperties().setProperty(
    'lastEOCheckDate',
    Utilities.formatDate(today, tz, 'yyyy/MM/dd')
  );

  console.log("Chunked backfill complete. Rows added: " + rowsAdded);
  try {
    SpreadsheetApp.getUi().alert("EO Backfill", rowsAdded + " new EO(s) added from 180-day backfill.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

// Debug: shows the plain body of the first failing P2P email to diagnose parse issues
function debugP2PSearch() {
  var ss = getEOSpreadsheet();
  var eoSheet = ss.getSheetByName("EO's");
  var results = [];
  var plainBodySample = null;

  var eoRowMap = {};
  if (eoSheet) {
    var eoData = eoSheet.getDataRange().getValues();
    for (var i = 1; i < eoData.length; i++) {
      var eoNum = String(eoData[i][1]).trim();
      if (eoNum) eoRowMap[eoNum] = i;
    }
    results.push("EO's tab has " + (eoData.length - 1) + " rows.\n");
  } else {
    results.push("EO's tab not found.\n");
  }

  var threads = GmailApp.search('from:donotreply@verizon.com subject:"Project to Project Transfer" newer_than:90d');
  results.push("P2P emails found: " + threads.length);
  var parsed = 0, failed = 0;

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var body = msg.getPlainBody();
      var allEOs = body.match(/E\d{9}/g) || [];
      var uniqueEOs = [];
      allEOs.forEach(function(e) { if (uniqueEOs.indexOf(e) === -1) uniqueEOs.push(e); });
      var dataMatch = uniqueEOs.length >= 2 ? [null, uniqueEOs[0], uniqueEOs[1]] : null;
      if (!dataMatch) {
        if (msg.getSubject().indexOf('Failed') !== -1) {
          results.push("⏭️ skipped (transfer failed, no Receiving EO)");
        } else {
          failed++;
          if (!plainBodySample) plainBodySample = body;
        }
        return;
      }
      parsed++;
      var donating = dataMatch[1];
      var receiving = dataMatch[2];
      var tracked = eoRowMap.hasOwnProperty(donating);
      results.push((tracked ? "✅" : "❌ not tracked") + "  " + donating + " → " + receiving);
    });
  });

  results.push("\nParsed: " + parsed + "  |  Failed: " + failed);

  if (plainBodySample) {
    results.push("\n--- PLAIN BODY OF FIRST FAILING EMAIL (first 800 chars) ---");
    results.push(plainBodySample.substring(0, 800));
  }

  var output = results.join("\n");
  console.log(output);
  try {
    SpreadsheetApp.getUi().alert("P2P Debug", output, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

// Debug: shows plain body + parse results for the most recent UNeFI Submitted email
function debugEOEmailParsing() {
  var results = [];

  var searches = [
    'from:donotreply@verizon.com subject:"unefi request" newer_than:180d',
    'from:donotreply@verizon.com subject:"unefi request" newer_than:90d',
    'subject:"ESA EO NOTIFICATION - UNeFI Request Submitted" newer_than:90d',
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
  var headers = ["Date Received", "EO#", "Receiving EO#", "Install Location", "Install Location Desc", "MPN", "Requested Qty"];
  var sheet = ss.getSheetByName("EO's");

  if (!sheet) {
    sheet = ss.insertSheet("EO's");
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#4f81bd").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    return sheet;
  }

  // Migrate existing sheet: insert "Receiving EO#" column if not present
  var firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (firstRow[2] !== "Receiving EO#") {
    sheet.insertColumnAfter(2); // Insert after col B
    sheet.getRange(1, 3).setValue("Receiving EO#")
      .setBackground("#4f81bd").setFontColor("#ffffff").setFontWeight("bold");
  }

  return sheet;
}

function setupEONotificationTrigger() {
  var handlersToSetup = ['checkEONotifications', 'checkP2PTransfers'];
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (handlersToSetup.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkEONotifications').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('checkP2PTransfers').timeBased().everyDays(1).atHour(9).create();

  console.log("EO and P2P triggers set for 9 AM daily.");
  try {
    SpreadsheetApp.getUi().alert("EO Notifications", "Daily triggers set for 9 AM:\n• EO Notifications\n• P2P Transfers", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}
