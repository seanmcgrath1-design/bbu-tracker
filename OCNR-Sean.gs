// OCNR-Sean.gs
// Fires when an OCNR_*.xlsx file syncs to Google Drive.
// Converts it to a Google Sheet, adds a Month column, and builds a pivot table.
//
// ONE-TIME SETUP:
//   1. In the Apps Script editor, enable the Drive Advanced Service
//      (Extensions → Apps Script → Services → Drive API)
//   2. Run setupOCNRTrigger() once from the editor to install the Drive trigger

var OCNR_FOLDER_NAME = 'OCNR';

// ─── Setup ───────────────────────────────────────────────────────────────────

function setupOCNRTrigger() {
  // Remove any existing trigger so we don't create duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onOCNRDriveChange') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('onOCNRDriveChange')
    .forDrive()
    .onchange()
    .create();
  Logger.log('OCNR Drive trigger installed.');
}

// ─── Trigger handler ─────────────────────────────────────────────────────────

function onOCNRDriveChange(e) {
  if (!e || e.changeType !== 'ADDED') return;

  var file;
  try {
    file = DriveApp.getFileById(e.fileId);
  } catch (err) {
    return; // File not accessible yet — skip
  }

  // Only process OCNR_M_D_YY.xlsx files
  if (!/^OCNR_\d+_\d+_\d+\.xlsx$/.test(file.getName())) return;

  // Confirm it is inside the OCNR folder
  var parents = file.getParents();
  var inOCNR = false;
  while (parents.hasNext()) {
    if (parents.next().getName() === OCNR_FOLDER_NAME) { inOCNR = true; break; }
  }
  if (!inOCNR) return;

  processOCNRFile(file);
}

// ─── Main processing ─────────────────────────────────────────────────────────

function processOCNRFile(file) {
  var folder  = file.getParents().next();
  var gsName  = file.getName().replace('.xlsx', '');

  // Remove any previously converted Google Sheet with the same name
  var existing = folder.getFilesByName(gsName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  // Convert xlsx → Google Sheet (requires Drive Advanced Service)
  var copied = Drive.Files.copy(
    { title: gsName, mimeType: MimeType.GOOGLE_SHEETS },
    file.getId()
  );

  var spreadsheet = SpreadsheetApp.openById(copied.id);
  var sheet = spreadsheet.getSheets()[0];
  sheet.setName('Sheet1');

  addMonthColumn(sheet);
  createPivotTable(spreadsheet, sheet);

  Logger.log('OCNR processed → ' + gsName);
}

// ─── Month column ─────────────────────────────────────────────────────────────

function addMonthColumn(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var deliveryDateCol = headers.indexOf('Item Delivery Date') + 1;

  if (deliveryDateCol === 0) {
    Logger.log('Column "Item Delivery Date" not found — skipping Month column.');
    return;
  }

  var lastRow = sheet.getLastRow();
  var monthCol = sheet.getLastColumn() + 1;

  sheet.getRange(1, monthCol).setValue('Month');

  // Prefix with zero-padded month number so alphabetical sort = chronological order
  // e.g. "01 January", "02 February" ... "12 December"
  var offset = deliveryDateCol - monthCol; // relative column offset (negative)
  sheet
    .getRange(2, monthCol, lastRow - 1, 1)
    .setFormulaR1C1('=TEXT(RC[' + offset + '],"mm")&" "&TEXT(RC[' + offset + '],"mmmm")');
}

// ─── Pivot table ──────────────────────────────────────────────────────────────

function createPivotTable(spreadsheet, sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var supplierCol     = headers.indexOf('Supplier Name') + 1;
  var monthCol        = headers.indexOf('Month') + 1;
  var valueCol        = headers.indexOf('Total Value not received(W/O Taxes)') + 1;
  var itemCatCol      = headers.indexOf('Item Category Descr') + 1;

  var missing = [];
  if (!supplierCol)  missing.push('Supplier Name');
  if (!monthCol)     missing.push('Month');
  if (!valueCol)     missing.push('Total Value not received(W/O Taxes)');
  if (!itemCatCol)   missing.push('Item Category Descr');
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

  // Filter: Item Category Descr
  pivot.addFilter(itemCatCol, SpreadsheetApp.newFilterCriteria().build());

  // Rows: Supplier Name (A → Z)
  var rowGroup = pivot.addRowGroup(supplierCol);
  rowGroup.sortAscending();

  // Columns: Month
  pivot.addColumnGroup(monthCol);

  // Values: Sum of Total Value not received(W/O Taxes)
  pivot.addPivotValue(valueCol, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
}
