// Sheet Lookup Web App
// Deployed as a Google Apps Script web app.
// Returns data from the Daily Data Dump tab for a given Fuze Project ID.
//
// Deploy: Extensions → Apps Script → Deploy → New deployment
//   Type: Web app | Execute as: Me | Who has access: Anyone
// After deploying, paste the /exec URL into run-shelter-bom.js as SHEET_LOOKUP_URL.

function doGet(e) {
  var projectId = e && e.parameter && e.parameter.projectId;

  if (!projectId) {
    return jsonResponse({ error: 'Missing projectId parameter' });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Daily Data Dump');

  if (!sheet) {
    return jsonResponse({ error: 'Daily Data Dump tab not found' });
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ error: 'Daily Data Dump tab appears empty' });
  }

  // Read only column A (Fuze Project ID) and column CD (Construction Start (F))
  // Column CD = column number 82 (A=1 ... Z=26, AA=27 ... AZ=52, BA=53 ... BZ=78, CA=79, CB=80, CC=81, CD=82)
  var projectIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var constDates  = sheet.getRange(2, 82, lastRow - 1, 1).getValues();

  for (var i = 0; i < projectIds.length; i++) {
    if (String(projectIds[i][0]).trim() === String(projectId).trim()) {
      var val = constDates[i][0];
      var dateStr = '';

      if (val instanceof Date) {
        var mm = String(val.getMonth() + 1).padStart(2, '0');
        var dd = String(val.getDate()).padStart(2, '0');
        dateStr = mm + '/' + dd + '/' + val.getFullYear();
      } else {
        dateStr = String(val).trim();
      }

      if (!dateStr) {
        return jsonResponse({ error: 'Construction Start (F) is blank for project ' + projectId });
      }

      return jsonResponse({ date: dateStr });
    }
  }

  return jsonResponse({ error: 'Project ID ' + projectId + ' not found in Daily Data Dump' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
