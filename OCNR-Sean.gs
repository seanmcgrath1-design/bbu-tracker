// OCNR-Sean.gs
// Checks every 15 minutes for unprocessed OCNR_*.xlsx files in the OCNR Drive folder.
// Converts each one to a Google Sheet, adds a Month column, and builds a pivot table.
//
// ONE-TIME SETUP:
//   1. In the Apps Script editor, enable the Drive Advanced Service
//      (Services → Drive API)
//   2. Run setupOCNRTrigger() once from the editor to install the time-driven trigger

var OCNR_FOLDER_NAME = 'OCNR';

// ─── Setup ───────────────────────────────────────────────────────────────────

function setupOCNRTrigger() {
  // Remove any existing trigger so we don't create duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkForNewOCNRFiles') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('checkForNewOCNRFiles')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('OCNR 15-minute trigger installed.');
}

// ─── Trigger handler ─────────────────────────────────────────────────────────

function checkForNewOCNRFiles() {
  // Search Drive directly by file name pattern — avoids picking the wrong "OCNR" folder
  var files = DriveApp.searchFiles(
    'title contains "OCNR_" and mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" and trashed = false'
  );

  while (files.hasNext()) {
    var file = files.next();
    if (!/^OCNR_\d+_\d+_\d+\.xlsx$/.test(file.getName())) continue;

    // Skip if a Google Sheet with the same name already exists in the same folder
    var gsName = file.getName().replace('.xlsx', '');
    var folder = file.getParents().next();
    var existing = folder.getFilesByName(gsName);
    if (existing.hasNext()) continue;

    processOCNRFile(file);
  }
}

// ─── Main processing ─────────────────────────────────────────────────────────

function processOCNRFile(file) {
  var folder = file.getParents().next();
  var gsName = file.getName().replace('.xlsx', '');

  // Convert xlsx → Google Sheet (requires Drive Advanced Service)
  var copied = Drive.Files.copy(
    { title: gsName, mimeType: MimeType.GOOGLE_SHEETS },
    file.getId()
  );

  var spreadsheet = SpreadsheetApp.openById(copied.id);
  var sheet = spreadsheet.getSheets()[0];
  sheet.setName('Sheet1');

  addMonthColumn(spreadsheet, sheet);
  SpreadsheetApp.flush(); // commit all writes before pivot reads the sheet
  createPivotTable(spreadsheet, sheet);

  Logger.log('OCNR processed → ' + gsName);
}

// ─── Month column ─────────────────────────────────────────────────────────────

function addMonthColumn(spreadsheet, sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var deliveryDateCol = headers.indexOf('Item Delivery Date') + 1;

  if (deliveryDateCol === 0) {
    Logger.log('Column "Item Delivery Date" not found — skipping Month column.');
    return;
  }

  var lastRow = sheet.getLastRow();
  var monthCol = sheet.getLastColumn() + 1;

  sheet.getRange(1, monthCol).setValue('Month');

  // Compute month values directly in JS so they're available immediately for the pivot table.
  // Format: "01 January", "02 February" ... so alphabetical sort = chronological order.
  var MONTH_NAMES = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];
  var dates = sheet.getRange(2, deliveryDateCol, lastRow - 1, 1).getValues();
  var monthValues = dates.map(function(row) {
    var d = row[0];
    if (!d || !(d instanceof Date)) return [''];
    var m = d.getMonth();
    return [(m < 9 ? '0' : '') + (m + 1) + ' ' + MONTH_NAMES[m]];
  });
  // Use Sheets API with RAW valueInputOption — the only reliable way to store
  // strings like "07 July" without Google Sheets auto-converting them to dates.
  var rangeA1 = "'" + sheet.getName() + "'!" +
    sheet.getRange(2, monthCol, lastRow - 1, 1).getA1Notation();
  Sheets.Spreadsheets.Values.update(
    { values: monthValues },
    spreadsheet.getId(),
    rangeA1,
    { valueInputOption: 'RAW' }
  );
}

// ─── Pivot table ──────────────────────────────────────────────────────────────

function createPivotTable(spreadsheet, sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var supplierCol = headers.indexOf('Supplier Name') + 1;
  var monthCol    = headers.indexOf('Month') + 1;
  var valueCol    = headers.indexOf('Total Value not received(W/O Taxes)') + 1;
  var itemCatCol  = headers.indexOf('Item Category Descr') + 1;

  var missing = [];
  if (!supplierCol) missing.push('Supplier Name');
  if (!monthCol)    missing.push('Month');
  if (!valueCol)    missing.push('Total Value not received(W/O Taxes)');
  if (missing.length) {
    Logger.log('Pivot skipped — columns not found: ' + missing.join(', '));
    return;
  }

  // Recreate the Pivot sheet fresh each run
  var pivotSheet = spreadsheet.getSheetByName('Pivot');
  if (pivotSheet) spreadsheet.deleteSheet(pivotSheet);
  pivotSheet = spreadsheet.insertSheet('Pivot');

  var sourceRange = sheet.getDataRange();
  var pivot = pivotSheet.getRange('A1').createPivotTable(sourceRange);

  // Rows: Supplier Name (A → Z)
  var rowGroup = pivot.addRowGroup(supplierCol);
  rowGroup.sortAscending();

  // Columns: Month
  pivot.addColumnGroup(monthCol);

  // Values: Sum of Total Value not received(W/O Taxes)
  pivot.addPivotValue(valueCol, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
}

// ─── Manual test ─────────────────────────────────────────────────────────────

function testOCNRManual() {
  var files = DriveApp.searchFiles(
    'title contains "OCNR_" and mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" and trashed = false'
  );
  while (files.hasNext()) {
    var f = files.next();
    if (/^OCNR_\d+_\d+_\d+\.xlsx$/.test(f.getName())) {
      Logger.log('Processing: ' + f.getName());
      processOCNRFile(f);
      return;
    }
  }
  Logger.log('No matching OCNR xlsx found in Drive.');
}
