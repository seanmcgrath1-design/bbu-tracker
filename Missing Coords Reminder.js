/**
 * AUTOMATION: Missing GPS Coordinates — RE Specialist Reminder Emails
 *
 * Scans Site Detail for sites flagged "Missing Coords" in the Distance column,
 * looks up the responsible Real Estate Specialist from the Daily Data Dump,
 * then sends each specialist one grouped reminder email listing their sites.
 *
 * Send rules:
 *   - No prior reminder → send (First Notice)
 *   - Reminded < 2 weeks ago → skip
 *   - Reminded 2+ weeks ago → send again (Follow-Up)
 *
 * Prerequisites:
 *   - "RE Contacts" tab must exist in the spreadsheet with:
 *       Column A: Name  (must match exactly what appears in Daily Data Dump)
 *       Column B: Email (full address, e.g. jane.doe@verizonwireless.com)
 *       Row 1 = headers, data starts row 2
 *   - A "Reminder Sent" column in Site Detail is created automatically on first run.
 */

const TWO_WEEKS_MS = 14 * 24 * 60 * 60 * 1000;

function sendMissingCoordsReminders() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const siteSheet = ss.getSheetByName("Site Detail");
  const dumpSheet = ss.getSheetByName("Daily Data Dump");
  const reSheet   = ss.getSheetByName("RE Contacts");

  if (!siteSheet || !dumpSheet || !reSheet) {
    const missing = [!siteSheet && "'Site Detail'", !dumpSheet && "'Daily Data Dump'", !reSheet && "'RE Contacts'"].filter(Boolean).join(", ");
    SpreadsheetApp.getUi().alert("Missing Tabs", `Could not find required tab(s): ${missing}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // ── 1. Build name → email map from RE Contacts tab ────────────────────────
  const reEmailMap = {};
  reSheet.getDataRange().getValues().slice(1).forEach(row => {
    const name  = String(row[0]).trim();
    const email = String(row[1]).trim();
    if (name && email) reEmailMap[name.toLowerCase()] = email;
  });

  if (Object.keys(reEmailMap).length === 0) {
    SpreadsheetApp.getUi().alert("RE Contacts Empty", "The 'RE Contacts' tab has no entries. Please add specialist names and emails.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // ── 2. Locate columns in Site Detail ──────────────────────────────────────
  const siteLastCol = siteSheet.getLastColumn();
  const siteRaw     = siteSheet.getRange(1, 1, 10, siteLastCol).getValues();
  const mIdData     = findCol(siteRaw, "fuze", "id");
  const mHeaderRow  = mIdData.row;
  const mFuzeCol    = mIdData.col;
  const mDistCol    = findCol(siteRaw, "distance").col;
  const mSiteCol    = findCol(siteRaw, "site").col;
  const mHubCol     = findCol(siteRaw, "hub").col;

  if (mFuzeCol === -1 || mDistCol === -1) {
    SpreadsheetApp.getUi().alert("Column Error", "Could not find 'Fuze ID' or 'Distance' columns in Site Detail.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // ── 3. Find or auto-create "Reminder Sent" column ─────────────────────────
  let mReminderCol = findCol(siteRaw, "reminder").col;
  if (mReminderCol === -1) {
    const newColNum = siteLastCol + 1;
    siteSheet.getRange(mHeaderRow + 1, newColNum).setValue("Reminder Sent");
    SpreadsheetApp.flush();
    mReminderCol = newColNum - 1; // store as 0-based to match row array indexing
    console.log(`Auto-created 'Reminder Sent' column at position ${newColNum}.`);
  }

  // ── 4. Scan Site Detail — collect eligible missing-coords sites ───────────
  const siteData     = siteSheet.getDataRange().getValues().slice(mHeaderRow + 1);
  const now          = new Date();
  const missingSites = [];
  let   skippedCount = 0;

  siteData.forEach((row, idx) => {
    if (String(row[mDistCol]).trim() !== "Missing Coords") return;

    const sentVal    = row[mReminderCol];
    let   isFollowUp = false;

    if (sentVal instanceof Date && !isNaN(sentVal.getTime())) {
      if ((now - sentVal) < TWO_WEEKS_MS) {
        skippedCount++; // reminded too recently — skip
        return;
      }
      isFollowUp = true; // 2+ weeks ago — eligible for follow-up
    }

    missingSites.push({
      fuzeId:     String(row[mFuzeCol]).replace(/\D/g, ""),
      siteName:   mSiteCol !== -1 ? String(row[mSiteCol]).trim() : "(Unknown)",
      hub:        mHubCol  !== -1 ? String(row[mHubCol]).trim()  : "",
      sheetRow:   idx + mHeaderRow + 2, // 1-based row number in the sheet
      isFollowUp: isFollowUp
    });
  });

  if (missingSites.length === 0) {
    let msg = "No sites are eligible for a reminder right now.";
    if (skippedCount > 0) msg += `\n\n⏳ ${skippedCount} site(s) were skipped — already reminded within the last 2 weeks.`;
    SpreadsheetApp.getUi().alert("Nothing to Send", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // ── 5. Look up RE Specialists from Daily Data Dump ────────────────────────
  const dumpRaw    = dumpSheet.getRange(1, 1, 10, dumpSheet.getLastColumn()).getValues();
  const dIdData    = findCol(dumpRaw, "fuze", "id");
  const dFuzeCol   = dIdData.col;
  const dHeaderRow = dIdData.row;
  const dReCol     = findCol(dumpRaw, "real", "estate").col;

  if (dFuzeCol === -1 || dReCol === -1) {
    SpreadsheetApp.getUi().alert("Column Error", "Could not find 'Fuze ID' or 'Real Estate Specialist' columns in Daily Data Dump.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const dumpMap = {};
  dumpSheet.getDataRange().getValues().slice(dHeaderRow + 1).forEach(row => {
    const id = String(row[dFuzeCol]).replace(/\D/g, "");
    if (id) dumpMap[id] = String(row[dReCol]).trim();
  });

  // ── 6. Group eligible sites by specialist ─────────────────────────────────
  const bySpecialist = {};
  const unmapped     = [];

  missingSites.forEach(site => {
    const reName = dumpMap[site.fuzeId];
    if (!reName) { unmapped.push({ ...site, reason: "Not found in Daily Data Dump" }); return; }

    const reEmail = reEmailMap[reName.toLowerCase()];
    if (!reEmail) { unmapped.push({ ...site, reName, reason: "Name not in RE Contacts tab" }); return; }

    if (!bySpecialist[reName]) bySpecialist[reName] = { email: reEmail, sites: [] };
    bySpecialist[reName].sites.push(site);
  });

  // ── 7. Send one email per specialist ──────────────────────────────────────
  const today      = Utilities.formatDate(now, "America/Chicago", "MM/dd/yyyy");
  let   emailsSent = 0;
  const stampRows  = []; // track which sheet rows to update after sending

  for (const specialist in bySpecialist) {
    const { email, sites } = bySpecialist[specialist];
    const hasFollowUp      = sites.some(s => s.isFollowUp);

    const siteRows = sites.map(s => {
      const noticeLabel = s.isFollowUp
        ? `<span style="color:#b45309; font-weight:bold;">Follow-Up</span>`
        : `<span style="color:#1565c0; font-weight:bold;">First Notice</span>`;
      return `<tr>
        <td style="padding:8px; border:1px solid #ccc;">${s.siteName}</td>
        <td style="padding:8px; border:1px solid #ccc;">${s.fuzeId}</td>
        <td style="padding:8px; border:1px solid #ccc;">${s.hub}</td>
        <td style="padding:8px; border:1px solid #ccc; text-align:center;">${noticeLabel}</td>
      </tr>`;
    }).join("");

    const heading  = hasFollowUp ? "Follow-Up: GPS Coordinates Required — Action Needed" : "GPS Coordinates Required — Action Needed";
    const subject  = hasFollowUp
      ? `[Follow-Up][Action Required] GPS Coordinates Missing — ${sites.length} Site(s) Need Update`
      : `[Action Required] GPS Coordinates Missing — ${sites.length} Site(s) Need Update`;

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 720px;">
        <h2 style="color: #d52b1e; margin-bottom: 4px;">${heading}</h2>
        <p>Hi ${specialist},</p>
        <p>The following small cell site(s) in your portfolio are <strong>missing GPS coordinates</strong> in the Fuze system.
           These sites cannot be assigned to a BBU cluster until latitude/longitude coordinates are entered.</p>
        <p>Please update the GPS coordinates for the site(s) listed below at your earliest convenience.</p>
        <table cellpadding="0" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 14px;">
          <tr style="background-color: #d52b1e; color: #ffffff;">
            <th style="padding:10px; border:1px solid #ccc; text-align:left;">Site Name</th>
            <th style="padding:10px; border:1px solid #ccc; text-align:left;">Fuze Project ID</th>
            <th style="padding:10px; border:1px solid #ccc; text-align:left;">Hub</th>
            <th style="padding:10px; border:1px solid #ccc; text-align:center;">Notice Type</th>
          </tr>
          ${siteRows}
        </table>
        <br>
        <p>If you have any questions, please reach out to <strong>Sean McGrath</strong> or <strong>Enis Orahovac</strong>.</p>
        <p style="font-size: 11px; color: #999;">This is an automated reminder generated by the BBU Mapping Tool on ${today}.</p>
      </div>`;

    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody, cc: "sean.mcgrath1@verizonwireless.com, enis.orahovac@verizonwireless.com" });
    sites.forEach(s => stampRows.push(s.sheetRow));
    emailsSent++;
  }

  // ── 8. Stamp "Reminder Sent" date for all successfully emailed sites ───────
  if (stampRows.length > 0) {
    const reminderColNum = mReminderCol + 1; // convert 0-based → 1-based for getRange
    stampRows.forEach(rowNum => siteSheet.getRange(rowNum, reminderColNum).setValue(now));
  }

  // ── 9. Log unmapped sites ─────────────────────────────────────────────────
  if (unmapped.length > 0) {
    console.warn("Missing Coords Reminder — sites that could not be emailed:", JSON.stringify(unmapped, null, 2));
  }

  // ── 10. Summary alert (only visible on manual runs) ───────────────────────
  try {
    const sentSites     = missingSites.filter(s => !unmapped.find(u => u.fuzeId === s.fuzeId));
    const firstCount    = sentSites.filter(s => !s.isFollowUp).length;
    const followUpCount = sentSites.filter(s =>  s.isFollowUp).length;

    let msg = `Sent ${emailsSent} email(s):\n  • ${firstCount} first notice(s)\n  • ${followUpCount} follow-up(s)`;
    if (skippedCount > 0) msg += `\n\n⏳ ${skippedCount} site(s) skipped — reminded within the last 2 weeks.`;
    if (unmapped.length > 0) {
      msg += `\n\n⚠️ ${unmapped.length} site(s) could not be emailed:\n`;
      msg += unmapped.map(u => `  • ${u.siteName} (${u.fuzeId}) — ${u.reason}`).join("\n");
    }
    SpreadsheetApp.getUi().alert("Missing Coords Reminder", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // Swallow UI error when running from a time-based trigger
  }
}

/**
 * Run once to schedule sendMissingCoordsReminders() every Monday at 8 AM.
 */
function setupMissingCoordsWeeklyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "sendMissingCoordsReminders") ScriptApp.deleteTrigger(t);
  }

  ScriptApp.newTrigger("sendMissingCoordsReminders")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert("Trigger Set", "Weekly reminder scheduled! It will run every Monday morning at 8 AM.", SpreadsheetApp.getUi().ButtonSet.OK);
}
