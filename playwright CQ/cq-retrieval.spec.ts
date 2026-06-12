import { test, Page } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';

// --- Config (edit paths here if your folders differ) ---
const DOWNLOADS_DIR = 'C:\\Users\\mcgr1se\\Downloads';
// Synced Google Drive folder: bbu-tracker/CQ (mirrors locally to G:\My Drive\bbu-tracker\CQ)
const CQ_DRIVE_DIR = 'G:\\My Drive\\bbu-tracker\\CQ';

// SPM tracker grid — used only to trigger the SSO login. Fresh tokens each run.
const TRACKER_URL =
  'https://fuze.verizon.com/tracker/app/jsp/spmTracker/spm-tracker-grid.jsp?spm_tracker_id=3572789';
// Project detail page — navigated to directly by Fuze ID (bypasses the grid filter).
const PROJECT_URL = (fuzeId: string) =>
  `https://fuze.verizon.com/spm/projects.jsp?projectId=${fuzeId}`;

// Strip Windows-illegal filename characters but keep spaces (site names like "DENSE _26"
// stay readable). Matching in Handoff.gs is keyed on the Fuze ID prefix, not this part.
function fsSafe(value: string): string {
  return value.replace(/[\\/:*?"<>|]/g, '').replace(/\s+/g, ' ').trim();
}

// Download one CQ for a Fuze ID (assumes already logged in). Saves to Downloads then moves to the
// Drive CQ folder. Throws on failure so the caller can record it per-site.
async function retrieveCq(page: Page, fuzeId: string): Promise<string> {
  await page.goto(PROJECT_URL(fuzeId));
  await page.waitForLoadState('networkidle');

  // Site name from the breadcrumb link (href = /spm/sites.jsp?siteId=...)
  let siteName = '';
  try {
    siteName = (await page.locator('a[href*="/spm/sites.jsp?siteId="]').first().innerText({ timeout: 15000 })).trim();
  } catch {
    console.warn(`  (${fuzeId}) could not read site name — using Fuze ID only.`);
  }

  await page.getByRole('button', { name: 'ConQuest (CQ) Milestone' }).click();
  // Recorded name was "Create CQ  COMPLETED 08/22..." — date is site-specific; match the prefix.
  await page.getByRole('button', { name: 'Create CQ' }).first().click();

  const editPromise = page.waitForEvent('popup');
  await page.getByRole('link', { name: 'Edit CQ' }).click();
  const editCq = await editPromise;

  const base = siteName ? `CQ_${fuzeId}_${siteName}` : `CQ_${fuzeId}`;
  const filename = `${fsSafe(base)}.xlsx`;
  const downloadPath = path.join(DOWNLOADS_DIR, filename);
  if (!fs.existsSync(DOWNLOADS_DIR)) fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });

  const downloadPromise = editCq.waitForEvent('download');
  await editCq.getByRole('button', { name: 'Generate CQ' }).click();
  const download = await downloadPromise;
  await download.saveAs(downloadPath);
  await editCq.close();

  // Move from Downloads into the synced Google Drive CQ folder.
  if (!fs.existsSync(CQ_DRIVE_DIR)) fs.mkdirSync(CQ_DRIVE_DIR, { recursive: true });
  const destPath = path.join(CQ_DRIVE_DIR, filename);
  try {
    fs.renameSync(downloadPath, destPath);
  } catch (err: any) {
    if (err && err.code === 'EXDEV') {
      fs.copyFileSync(downloadPath, destPath);
      fs.unlinkSync(downloadPath);
    } else {
      throw err;
    }
  }
  return destPath;
}

test('CQ Retrieval', async ({ page }) => {
  // Batch (FUZE_IDS=comma,list) or single (FUZE_ID). Batch scales the timeout per site.
  const fuzeIds = (process.env.FUZE_IDS ?? process.env.FUZE_ID ?? '')
    .split(',').map(s => s.trim()).filter(Boolean);
  if (fuzeIds.length === 0) throw new Error('FUZE_ID or FUZE_IDS env var is required.');
  test.setTimeout(120000 + fuzeIds.length * 90000);

  // --- Login (Verizon SSO) ---
  await page.goto(TRACKER_URL);
  await page.getByRole('textbox', { name: 'Enter your User Name' }).fill('mcgr1se');
  await page.getByRole('textbox', { name: 'Enter your Password' }).fill(process.env.VZ_PASSWORD ?? '');
  await page.getByRole('button', { name: 'Log in' }).click();
  await page.waitForLoadState('networkidle');

  // --- Retrieve each CQ (continue on per-site failure) ---
  const failures: string[] = [];
  for (const fuzeId of fuzeIds) {
    try {
      const dest = await retrieveCq(page, fuzeId);
      console.log(`OK ${fuzeId} -> ${dest}`);
    } catch (err: any) {
      failures.push(fuzeId);
      console.error(`FAIL ${fuzeId}: ${err && err.message ? err.message : err}`);
    }
  }

  console.log(`\nCQ Retrieval done: ${fuzeIds.length - failures.length}/${fuzeIds.length} succeeded.`);
  if (failures.length) console.log(`Failed: ${failures.join(', ')}`);
  // Fail the test only if every site failed (e.g. login/navigation broke); partial success is fine —
  // the orchestrator confirms which CQs landed via the cqStatus poll.
  if (failures.length === fuzeIds.length) throw new Error('All CQ downloads failed.');
});
