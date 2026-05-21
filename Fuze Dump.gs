// Apps Script Code: Code.gs

/**
 * ============================================================================
 * GOOGLE APPS AUTOMATION: DAILY FUZE DUMP IMPORT (LIVE PLAN)
 * ============================================================================
 * Description: Searches Gmail for the daily "Fuze Dump" email with a CSV attachment,
 * parses the data, imports it into the specified Google Sheet, and strictly forces
 * the "Daily Data Dump" tab into the 2nd position using UI flushing.
 * ============================================================================
 */

// --- CONFIGURATION ---
// IMPORTANT: Replace this with the ID of your live "6672 Plan" tracker.
const TARGET_SPREADSHEET_ID = '1Ada_FMW6YmE25puTyjYA4QYWUxmlFLpTptEdgviUif0';

// Optimized Gmail search query to find the correct unread email with an attachment
const SEARCH_QUERY = 'subject:"Fuze Dump" is:unread has:attachment filename:csv';
const TARGET_TAB_NAME = 'Daily Data Dump';

/**
 * Main function designed to be run via a Time-Driven Trigger.
 */
function dailyFuzeDumpAutomation() {
  console.log('Starting daily Fuze Dump automation...');
  
  // 1. Find the relevant email thread in Gmail
  const threads = GmailApp.search(SEARCH_QUERY, 0, 1);
  
  if (threads.length === 0) {
    console.log('No unread emails found matching the search criteria. Exiting.');
    return;
  }
  
  // 2. Get the latest message and its attachments
  const message = threads[0].getMessages()[0];
  const attachments = message.getAttachments();
  
  if (attachments.length === 0) {
    console.log('Email found, but no attachments were present. Exiting.');
    return;
  }
  
  // Filter for the CSV attachment specifically
  const csvAttachment = attachments.find(att => att.getContentType() === 'text/csv' || att.getName().endsWith('.csv'));
  
  if (!csvAttachment) {
    console.log('No CSV attachment found in the email. Exiting.');
    return;
  }
  
  // 3. Parse the CSV data
  let csvData;
  try {
    const csvString = csvAttachment.getDataAsString();
    csvData = Utilities.parseCsv(csvString);
    console.log(`Successfully parsed CSV attachment. Found ${csvData.length} rows.`);
  } catch (error) {
    console.error('Error parsing CSV attachment: ' + error.message);
    return;
  }
  
  // 4. Connect to the target Google Sheet
  const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(TARGET_TAB_NAME);
  
  // 5. Create or Reset the Sheet, and Force Position
  if (!sheet) {
    // insertSheet(name, index) uses 0-based indexing. 1 strictly places it 2nd.
    sheet = spreadsheet.insertSheet(TARGET_TAB_NAME, 1);
    console.log(`Created new tab named "${TARGET_TAB_NAME}" strictly at the 2nd position.`);
  } else {
    // Clear existing data
    sheet.clear();
    console.log(`Cleared existing tab named "${TARGET_TAB_NAME}".`);
    
    // Activate the sheet and FORCE the system to sync changes before moving
    sheet.activate();
    SpreadsheetApp.flush(); 
    
    if (spreadsheet.getNumSheets() >= 2) {
      // moveActiveSheet uses 1-based indexing. 2 places it strictly as the second tab.
      spreadsheet.moveActiveSheet(2);
      SpreadsheetApp.flush(); // Force sync again to confirm the move
      console.log(`Forced the tab to the 2nd position and flushed UI.`);
    }
  }
  
  // 6. Apply the timestamp to Cell A1
  const timeStamp = Utilities.formatDate(new Date(), "America/Chicago", "MM/dd/yyyy HH:mm:ss z");
  const timestampRange = sheet.getRange('A1');
  timestampRange.setValue(`Data Imported On: ${timeStamp}`);
  timestampRange.setFontWeight('bold');
  
  // 7. Output the CSV Data starting at Row 2
  if (csvData.length > 0 && csvData[0].length > 0) {
    const numRows = csvData.length;
    const numCols = csvData[0].length;
    
    // Efficiently write the data array to the sheet in a single API call
    const dataRange = sheet.getRange(2, 1, numRows, numCols);
    dataRange.setValues(csvData);
    console.log('Data successfully written to the sheet.');
  }
  
  // 8. Mark the email as read so we don't process it again next time
  message.markRead();
  console.log('Email marked as read. Automation complete!');
}

/**
 * Run this function ONCE to programmatically set up your daily trigger.
 */
function createDailyTrigger() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (let trigger of existingTriggers) {
    if (trigger.getHandlerFunction() === 'dailyFuzeDumpAutomation') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create a new trigger to run every day between 6 AM and 7 AM
  ScriptApp.newTrigger('dailyFuzeDumpAutomation')
    .timeBased()
    .everyDays(1)
    .atHour(6) // Triggers between 6:00 AM and 7:00 AM
    .create();
    
  console.log('Time-driven trigger successfully configured for 6 AM - 7 AM daily.');
}