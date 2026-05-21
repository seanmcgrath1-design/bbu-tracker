// ==========================================
// PO CONFIRMATION TRACKER
// ==========================================

// Searches for handoff emails from Enis and logs them to the "PO Triggers" sheet
// as a to-do list for Sean to know which sites need a PO request.
function checkForEnisHandoffs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var triggersSheet = ss.getSheetByName("PO Triggers");
  if (!triggersSheet) {
    triggersSheet = ss.insertSheet("PO Triggers");
    var header = ["Date Received", "Hub Site", "Active Sites", "Email Subject", "Status"];
    triggersSheet.appendRow(header);
    triggersSheet.getRange(1, 1, 1, header.length)
      .setBackground("#e06666").setFontColor("#ffffff").setFontWeight("bold");
    triggersSheet.setFrozenRows(1);
  }

  // Build set of already-logged subjects to avoid duplicates
  var existingData = triggersSheet.getDataRange().getValues();
  var existingSubjects = {};
  for (var i = 1; i < existingData.length; i++) {
    existingSubjects[String(existingData[i][3]).trim()] = true;
  }

  // Search only since last run; first run goes back 90 days to catch everything
  var props = PropertiesService.getScriptProperties();
  var lastCheck = props.getProperty('lastEnisCheckDate');
  var searchQuery = 'from:enis.orahovac@verizonwireless.com subject:"Ready for" ' +
    (lastCheck ? 'after:' + lastCheck : 'newer_than:90d');

  var threads = GmailApp.search(searchQuery);
  var rowsAdded = 0;

  threads.forEach(function(thread) {
    var msg = thread.getMessages()[0];
    var subject = msg.getSubject();
    if (existingSubjects[subject]) return;
    if (subject.indexOf("Small Cell Mod Ready for Network Assurance:") !== -1) return;
    if (subject.indexOf("48 Hour Review Document") !== -1) return;

    var body = msg.getPlainBody();

    // Extract active sites from "Work for" line
    var allSites = "";
    var workForMatch = body.match(/Work for\s+([^\.\n\r<]+)/i);
    if (workForMatch) allSites = workForMatch[1].replace(/\*/g, '').trim();

    // Extract hub site from "Hub Site" line
    var hubSite = "";
    var hubMatch = body.match(/Hub Site[^\n]*\n([^\n]+)/i);
    if (hubMatch) hubSite = hubMatch[1].replace(/\*/g, '').trim();

    triggersSheet.appendRow([
      msg.getDate(),
      hubSite,
      allSites || subject,
      subject,
      "Awaiting PO Request"
    ]);

    existingSubjects[subject] = true;
    rowsAdded++;
  });

  // Save today as the last check date so tomorrow only scans new emails
  props.setProperty('lastEnisCheckDate', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd'));

  console.log("Enis handoff check complete. New rows added: " + rowsAdded);

  // Also check if any pending rows have since been sent by Sean
  var sentUpdates = checkSentHandoffs(ss);

  try {
    var ui = SpreadsheetApp.getUi();
    var msg = "";
    if (rowsAdded > 0) msg += rowsAdded + " new Enis handoff(s) added.\n";
    if (sentUpdates > 0) msg += sentUpdates + " row(s) updated to 'Handoff Sent'.\n";
    if (!msg) msg = "No new Enis handoffs found.";
    ui.alert("PO Triggers", msg.trim(), ui.ButtonSet.OK);
  } catch(e) {}
}

// Scans PO Triggers for "Awaiting PO Request" rows and marks them "Handoff Sent"
// if a matching email is found in Sean's sent mail.
function checkSentHandoffs(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggersSheet = ss.getSheetByName("PO Triggers");
  if (!triggersSheet) return 0;

  var data = triggersSheet.getDataRange().getValues();
  var COL_SITES = 2, COL_STATUS = 4;
  var updated = 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_STATUS]).trim() !== "Awaiting PO Request") continue;

    var activeSites = String(data[i][COL_SITES]).trim();
    if (!activeSites) continue;

    // Use the first site name as the search anchor
    var firstSite = activeSites.split(/\s*&\s*/)[0].trim();
    if (!firstSite) continue;

    var threads = GmailApp.search('in:sent subject:"Ready for" "' + firstSite + '"');
    if (threads.length > 0) {
      var sentDate = threads[0].getMessages()[0].getDate();
      var sentDateStr = Utilities.formatDate(sentDate, Session.getScriptTimeZone(), "M/d/yyyy");
      triggersSheet.getRange(i + 1, COL_STATUS + 1).setValue("Handoff Sent " + sentDateStr);
      updated++;
    }
  }

  console.log("Sent handoff check complete. Rows updated: " + updated);
  return updated;
}

function checkPOConfirmationsAndDraft() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ensureTrackingSheet(ss);

  var trackingData = trackingSheet.getDataRange().getValues();

  var COL_DATE = 0, COL_HUB = 1, COL_ALL_SITES = 2, COL_SITE = 3;
  var COL_RTN = 4, COL_SUBJECT = 5, COL_PO = 6, COL_AMOUNT = 7;
  var COL_TYPE = 8, COL_STATUS = 9;

  // Build index of all known RTN/EOs → row index
  var knownRtns = {};
  for (var i = 1; i < trackingData.length; i++) {
    knownRtns[String(trackingData[i][COL_RTN]).trim()] = i;
  }

  // 1. Process all PO confirmation emails
  var unmatchedPOs = [];
  var threads = GmailApp.search('subject:"PURCHASE ORDER PENDING" newer_than:90d');

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var body = msg.getPlainBody();
      var po = parsePOConfirmationEmail(body);
      if (!po) return;

      if (knownRtns.hasOwnProperty(po.rtnEO)) {
        // RTN/EO is tracked — update if still Pending
        var rowIdx = knownRtns[po.rtnEO];
        if (String(trackingData[rowIdx][COL_STATUS]).trim() === "Pending") {
          trackingSheet.getRange(rowIdx + 1, COL_PO + 1).setValue(po.poNumber);
          trackingSheet.getRange(rowIdx + 1, COL_AMOUNT + 1).setValue(po.totalAmount);
          trackingSheet.getRange(rowIdx + 1, COL_TYPE + 1).setValue(po.poType);
          trackingSheet.getRange(rowIdx + 1, COL_STATUS + 1).setValue("Confirmed");

          trackingData[rowIdx][COL_PO] = po.poNumber;
          trackingData[rowIdx][COL_AMOUNT] = po.totalAmount;
          trackingData[rowIdx][COL_TYPE] = po.poType;
          trackingData[rowIdx][COL_STATUS] = "Confirmed";
        }
      } else {
        // RTN/EO not tracked — collect for auto-registration (handles Enis emails)
        unmatchedPOs.push({ po: po, emailDate: msg.getDate() });
      }
    });
  });

  // 2a. Repair any existing rows stuck with "Unmatched - " subjects
  for (var i = 1; i < trackingData.length; i++) {
    var rowSubj = String(trackingData[i][COL_SUBJECT]).trim();
    if (rowSubj.indexOf("Unmatched - ") !== 0) continue;

    var poRef = rowSubj.replace("Unmatched - ", "").trim();
    var rtnEO = String(trackingData[i][COL_RTN]).trim();
    var fixedThreads = GmailApp.search('subject:"Ready for" "' + poRef + '"');

    if (fixedThreads.length === 0 && poRef.indexOf('&') !== -1) {
      var fp = poRef.split('&')[0].trim();
      fixedThreads = GmailApp.search('subject:"Ready for" "' + fp + '"');
    }
    if (fixedThreads.length === 0) {
      fixedThreads = GmailApp.search('subject:"Ready for" "' + rtnEO + '"');
    }

    if (fixedThreads.length > 0) {
      var fixedSubj = fixedThreads[0].getFirstMessageSubject();
      var fixedBody = fixedThreads[0].getMessages()[0].getPlainBody();
      var fixedSites = "";
      var fixedMatch = fixedBody.match(/Work for\s+([^\.\n\r<]+)/i);
      if (fixedMatch) fixedSites = fixedMatch[1].replace(/\*/g, '').trim();

      trackingSheet.getRange(i + 1, COL_SUBJECT + 1).setValue(fixedSubj);
      if (fixedSites) trackingSheet.getRange(i + 1, COL_ALL_SITES + 1).setValue(fixedSites);
      trackingData[i][COL_SUBJECT] = fixedSubj;
      if (fixedSites) trackingData[i][COL_ALL_SITES] = fixedSites;
      console.log("Repaired unmatched row: " + poRef + " → " + fixedSubj);
    }
  }

  // 2b. Auto-register any unmatched POs (searches both Sean's and Enis's emails)
  if (unmatchedPOs.length > 0) {
    var newRows = buildNewTrackingRows(unmatchedPOs);
    newRows.forEach(function(row) {
      trackingSheet.appendRow(row);
      trackingData.push(row);
      knownRtns[String(row[COL_RTN]).trim()] = trackingData.length - 1;
    });
  }

  // 3. Group all rows by email subject, skip already-drafted groups
  var subjectGroups = {};
  for (var i = 1; i < trackingData.length; i++) {
    var subj = String(trackingData[i][COL_SUBJECT]).trim();
    var status = String(trackingData[i][COL_STATUS]).trim();
    if (!subj || status === "Draft Created") continue;

    if (!subjectGroups[subj]) subjectGroups[subj] = [];
    subjectGroups[subj].push({ rowIndex: i, data: trackingData[i] });
  }

  // 4. Create reply draft for any fully confirmed group
  var draftsCreated = 0;
  for (var subj in subjectGroups) {
    var group = subjectGroups[subj];
    var allConfirmed = group.every(function(r) {
      return String(r.data[COL_STATUS]).trim() === "Confirmed";
    });

    if (allConfirmed) {
      var success = createPOReplyDraft(subj, group, COL_ALL_SITES, COL_SITE, COL_PO, COL_TYPE);
      if (success) {
        group.forEach(function(r) {
          trackingSheet.getRange(r.rowIndex + 1, COL_STATUS + 1).setValue("Draft Created");
        });
        draftsCreated++;
      }
    }
  }

  console.log("PO check complete. Reply drafts created: " + draftsCreated);

  try {
    var ui = SpreadsheetApp.getUi();
    if (draftsCreated > 0) {
      ui.alert("PO Tracker", draftsCreated + " reply draft(s) created. Check your Gmail drafts.", ui.ButtonSet.OK);
    } else {
      ui.alert("PO Tracker", "No new POs ready for reply. Check back later.", ui.ButtonSet.OK);
    }
  } catch(e) {}
}

// One-time backfill for existing PO emails not yet in the tracking sheet
function backfillPoTracking() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ensureTrackingSheet(ss);

  var existingData = trackingSheet.getDataRange().getValues();
  var existingRtns = {};
  for (var i = 1; i < existingData.length; i++) {
    existingRtns[String(existingData[i][4]).trim()] = true;
  }

  var unmatchedPOs = [];
  var threads = GmailApp.search('subject:"PURCHASE ORDER PENDING" newer_than:90d');

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var body = msg.getPlainBody();
      var po = parsePOConfirmationEmail(body);
      if (!po || existingRtns[po.rtnEO]) return;
      unmatchedPOs.push({ po: po, emailDate: msg.getDate() });
    });
  });

  if (unmatchedPOs.length === 0) {
    try { SpreadsheetApp.getUi().alert("PO Tracker", "No new PO emails found to backfill.", SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
    return;
  }

  var newRows = buildNewTrackingRows(unmatchedPOs);
  newRows.forEach(function(row) { trackingSheet.appendRow(row); });

  console.log("Backfill complete. Running PO check to create drafts...");
  checkPOConfirmationsAndDraft();
}

// Shared helper: takes a list of unmatched POs, finds original emails (Sean or Enis),
// groups by subject, builds All Active Sites, and returns rows ready to append.
function buildNewTrackingRows(unmatchedPOs) {
  var items = [];

  unmatchedPOs.forEach(function(item) {
    var po = item.po;
    var emailSubject = "";
    var emailBodySites = "";

    // Search all mail (not just sent) so Enis's emails are found too.
    // Fallback searches handle site names containing "&" which Gmail misinterprets as a search operator.
    var origThreads = GmailApp.search('subject:"Ready for" "' + po.poRef + '"');

    if (origThreads.length === 0 && po.poRef.indexOf('&') !== -1) {
      var firstPart = po.poRef.split('&')[0].trim();
      origThreads = GmailApp.search('subject:"Ready for" "' + firstPart + '"');
    }

    if (origThreads.length === 0) {
      origThreads = GmailApp.search('subject:"Ready for" "' + po.rtnEO + '"');
    }

    if (origThreads.length > 0) {
      emailSubject = origThreads[0].getFirstMessageSubject();
      if (emailSubject.indexOf("Small Cell Mod Ready for Network Assurance:") !== -1) return;
      if (emailSubject.indexOf("48 Hour Review Document") !== -1) return;
      var bodyText = origThreads[0].getMessages()[0].getPlainBody();
      var workForMatch = bodyText.match(/Work for\s+([^\.\n\r<]+)/i);
      if (workForMatch) emailBodySites = workForMatch[1].replace(/\*/g, '').trim();
    } else {
      emailSubject = "Unmatched - " + po.poRef;
    }

    items.push({
      emailSubject:   emailSubject,
      emailBodySites: emailBodySites,
      emailDate:      item.emailDate,
      po:             po
    });
  });

  // Group by email subject, primary PO first in each group
  var subjectGroups = {};
  items.forEach(function(item) {
    if (!subjectGroups[item.emailSubject]) subjectGroups[item.emailSubject] = [];
    subjectGroups[item.emailSubject].push(item);
  });

  var rows = [];
  for (var subj in subjectGroups) {
    var groupItems = subjectGroups[subj];
    groupItems.sort(function(a, b) {
      if (a.po.poType === b.po.poType) return 0;
      return a.po.poType === "DWDM Only" ? 1 : -1;
    });

    // Prefer body extraction, then subject, then PO Refs
    var allSites = groupItems[0].emailBodySites || "";
    if (!allSites) {
      var dashIdx = subj.indexOf(" - ");
      allSites = dashIdx !== -1 ? subj.substring(dashIdx + 3).trim() : "";
    }
    if (!allSites) {
      allSites = groupItems.map(function(item) { return item.po.poRef; }).join(" & ");
    }

    groupItems.forEach(function(item) {
      rows.push([
        item.emailDate,
        "",
        allSites,
        item.po.poRef,
        item.po.rtnEO,
        subj,
        item.po.poNumber,
        item.po.totalAmount,
        item.po.poType,
        "Confirmed"
      ]);
    });
  }

  return rows;
}

function parsePOConfirmationEmail(body) {
  // Plain body wraps bold fields with asterisks: "* Purchase Order: * 3002673065"
  var poMatch     = body.match(/Purchase Order:\s*\*?\s*(\d+)/);
  var rtnMatch    = body.match(/RTN\/EO:\s*\*?\s*(S\d+)/i);
  var poRefMatch  = body.match(/PO Ref:\s*\*?\s*([^\n\r*]+)/);
  var amountMatch = body.match(/Total Amount:\s*\*?\s*([\d,]+\.\d+)/);

  if (!poMatch || !rtnMatch) return null;

  var totalAmount  = amountMatch ? amountMatch[1].trim() : "0";
  var totalNumeric = parseFloat(totalAmount.replace(/,/g, ''));

  var poType;
  if (totalNumeric === 950) {
    poType = "DWDM Only";
  } else if (totalNumeric > 950) {
    poType = "Install/Integration + DWDM";
  } else {
    return null; // Not a recognized PO type — skip
  }

  return {
    poNumber:    poMatch[1].trim(),
    rtnEO:       rtnMatch[1].trim(),
    poRef:       poRefMatch ? poRefMatch[1].trim() : "",
    totalAmount: totalAmount,
    poType:      poType
  };
}

function createPOReplyDraft(subject, group, COL_ALL_SITES, COL_SITE, COL_PO, COL_TYPE) {
  var threads = GmailApp.search('subject:"' + subject + '" in:sent');
  if (threads.length === 0) threads = GmailApp.search('subject:"' + subject + '"');

  if (threads.length === 0) {
    console.error("Could not find original thread for: " + subject);
    return false;
  }

  // Build MDG Location ID lookup from Daily Data Dump (site name → MDG ID)
  var mdgMap = {};
  var dumpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Data Dump");
  if (dumpSheet) {
    var dumpData = dumpSheet.getDataRange().getValues();
    var dumpHeaders = dumpData[1];
    var colDumpSite = -1, colDumpMDG = -1;
    for (var h = 0; h < dumpHeaders.length; h++) {
      var hName = String(dumpHeaders[h]).trim().toLowerCase();
      if (hName === "site name") colDumpSite = h;
      if (hName === "mdg location id") colDumpMDG = h;
    }
    if (colDumpSite !== -1 && colDumpMDG !== -1) {
      for (var d = 2; d < dumpData.length; d++) {
        var siteName = String(dumpData[d][colDumpSite]).trim();
        var mdgId = String(dumpData[d][colDumpMDG]).trim();
        if (siteName && mdgId && mdgId !== "") mdgMap[siteName] = mdgId;
      }
    }
  }

  var thread = threads[0];
  var allActiveSites = String(group[0].data[COL_ALL_SITES]);

  var primaryRows = [];
  var dwdmRows = [];
  group.forEach(function(r) {
    if (String(r.data[COL_TYPE]) === "DWDM Only") {
      dwdmRows.push(r);
    } else {
      primaryRows.push(r);
    }
  });

  var thStyle = "background-color: #4f81bd; color: #ffffff; border: 1px solid black; padding: 8px; text-align: left;";
  var tdStyle = "border: 1px solid black; padding: 8px;";

  var htmlBody = "<div style='font-family: Arial, sans-serif; font-size: 14px;'>";
  htmlBody += "<p>Team,</p>";
  htmlBody += "<p>The Purchase Order(s) have been issued for <strong>" + allActiveSites + "</strong>. Please see details below:</p>";
  htmlBody += "<table cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-size: 14px; border: 1px solid black;'>";
  htmlBody += "<tr>";
  htmlBody += "<th style='" + thStyle + "'>Description</th>";
  htmlBody += "<th style='" + thStyle + "'>Site(s)</th>";
  htmlBody += "<th style='" + thStyle + "'>PO #</th>";
  htmlBody += "<th style='" + thStyle + "'>MDG Location ID</th>";
  htmlBody += "</tr>";

  // One Install/Integration row per primary PO, each showing only its own site
  primaryRows.forEach(function(pr) {
    var siteName = String(pr.data[COL_SITE]);
    var mdg = mdgMap[siteName] || "";
    htmlBody += "<tr><td style='" + tdStyle + "'>Install / Integration</td>";
    htmlBody += "<td style='" + tdStyle + "'>" + siteName + "</td>";
    htmlBody += "<td style='" + tdStyle + "'>" + String(pr.data[COL_PO]) + "</td>";
    htmlBody += "<td style='" + tdStyle + "'>" + mdg + "</td></tr>";

    // DWDM is bundled with this PO (type = "Install/Integration + DWDM")
    htmlBody += "<tr><td style='" + tdStyle + "'>DWDM Install</td>";
    htmlBody += "<td style='" + tdStyle + "'>" + siteName + "</td>";
    htmlBody += "<td style='" + tdStyle + "'>" + String(pr.data[COL_PO]) + "</td>";
    htmlBody += "<td style='" + tdStyle + "'>" + mdg + "</td></tr>";
  });

  // Standalone DWDM-only rows
  dwdmRows.forEach(function(r) {
    var siteName = String(r.data[COL_SITE]);
    var mdg = mdgMap[siteName] || "";
    htmlBody += "<tr><td style='" + tdStyle + "'>DWDM Install</td>";
    htmlBody += "<td style='" + tdStyle + "'>" + siteName + "</td>";
    htmlBody += "<td style='" + tdStyle + "'>" + String(r.data[COL_PO]) + "</td>";
    htmlBody += "<td style='" + tdStyle + "'>" + mdg + "</td></tr>";
  });

  htmlBody += "</table></div>";

  try {
    thread.createDraftReplyAll('', { htmlBody: htmlBody });
    return true;
  } catch(e) {
    console.error("Failed to create reply draft: " + e.message);
    return false;
  }
}

function ensureTrackingSheet(ss) {
  var sheet = ss.getSheetByName("PO Tracking");
  if (!sheet) {
    sheet = ss.insertSheet("PO Tracking");
    var header = ["Date Created", "Hub Site", "All Active Sites", "Site (PO Ref)", "RTN/EO", "Email Subject", "PO #", "Amount", "PO Type", "Status"];
    sheet.appendRow(header);
    sheet.getRange(1, 1, 1, header.length)
      .setBackground("#4f81bd").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function debugPOEmailSearch() {
  var results = [];
  var threads = GmailApp.search('subject:"PURCHASE ORDER PENDING" newer_than:90d');
  results.push('Found ' + threads.length + ' threads.\n');

  if (threads.length > 0) {
    var msg = threads[0].getMessages()[0];
    results.push('Subject: ' + msg.getSubject());
    results.push('From: ' + msg.getFrom());
    results.push('\n--- PLAIN BODY (first 800 chars) ---');
    results.push(msg.getPlainBody().substring(0, 800));
    results.push('\n--- PARSE RESULT ---');
    var po = parsePOConfirmationEmail(msg.getPlainBody());
    results.push(po ? JSON.stringify(po) : 'NULL - regex did not match');
  }

  var output = results.join('\n');
  console.log(output);
  try {
    SpreadsheetApp.getUi().alert("PO Email Debug", output, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

function setupPOTrackerTrigger() {
  var handlersToSetup = ['checkPOConfirmationsAndDraft', 'checkForEnisHandoffs'];

  // Remove existing triggers for these functions
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (handlersToSetup.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // PO confirmation check at 7 AM daily
  ScriptApp.newTrigger('checkPOConfirmationsAndDraft')
    .timeBased().everyDays(1).atHour(7).create();

  // Enis handoff check at 8 AM daily
  ScriptApp.newTrigger('checkForEnisHandoffs')
    .timeBased().everyDays(1).atHour(8).create();

  console.log("PO Tracker daily triggers set.");
  try {
    SpreadsheetApp.getUi().alert("PO Tracker", "Daily triggers set:\n• PO confirmations checked at 7 AM\n• Enis handoffs checked at 8 AM", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}
