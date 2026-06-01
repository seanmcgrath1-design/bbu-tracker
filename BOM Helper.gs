// BOM Helper — web app endpoint for the Shelter BOM Node.js tool
// Deploy as: Execute as Me, Anyone can access
// Returns Construction Start (F) - 30 days for a given Fuze Project ID

function doGet(e) {
  var fuzeId = e && e.parameter && e.parameter.fuzeId;
  if (!fuzeId) {
    return jsonResponse({ error: 'Missing fuzeId parameter' });
  }

  try {
    var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Site Detail');
    if (!sheet) return jsonResponse({ error: 'Site Detail tab not found' });

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var fuzeCol = -1, constStartCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim().toLowerCase();
      if (h === 'fuze project id') fuzeCol = i;
      if (h.indexOf('construction') !== -1 && h.indexOf('(f)') !== -1) constStartCol = i;
    }

    if (fuzeCol === -1) return jsonResponse({ error: '"Fuze Project ID" column not found in Site Detail' });
    if (constStartCol === -1) return jsonResponse({ error: '"Construction Start (F)" column not found in Site Detail' });

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][fuzeCol]).trim() === fuzeId) {
        var dateVal = data[r][constStartCol];
        if (!dateVal) return jsonResponse({ error: 'No Construction Start (F) date for project ' + fuzeId });

        var constStart = dateVal instanceof Date ? dateVal : new Date(dateVal);
        var tz = Session.getScriptTimeZone();
        var constStartStr = Utilities.formatDate(constStart, tz, 'MM/dd/yyyy');

        var forecast = new Date(constStart.getTime() - 30 * 86400000);
        var forecastStr = Utilities.formatDate(forecast, tz, 'MM/dd/yyyy');

        return jsonResponse({ fuzeId: fuzeId, constructionStart: constStartStr, forecastDate: forecastStr });
      }
    }

    return jsonResponse({ error: 'Project ' + fuzeId + ' not found in Site Detail' });

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
