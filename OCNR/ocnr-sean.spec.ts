import { test } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';

test('OCNR-Sean', async ({ page }) => {
  test.setTimeout(180000);

  // --- Login ---
  // Go directly to Qlik hub — SSO redirects to login with fresh tokens each run
  await page.goto('https://fscqlik-tpa.verizon.com/analytics/');
  await page.waitForURL('**/ssologin.verizon.com/**', { timeout: 30000 });
  await page.getByRole('textbox', { name: 'Enter your User Name' }).fill('mcgr1se');
  await page.getByRole('textbox', { name: 'Enter your Password' }).fill(process.env.VZ_PASSWORD ?? '');
  await page.getByRole('button', { name: 'Log in' }).click();
  await page.waitForURL('**/fscqlik-tpa.verizon.com/**', { timeout: 60000 });
  await page.waitForLoadState('networkidle');

  // --- Navigate to app ---
  // Open the app directly in a new page — same browser context carries the session.
  // The original popup approach relied on a font-icon link that is no longer reliable.
  const page1 = await page.context().newPage();
  await page1.goto('https://fscqlik-tpa.verizon.com/analytics/sense/app/c9ba5d4c-8d56-45c7-ad96-9a1eedeaf7fb/overview');
  await page1.waitForLoadState('networkidle');

  // --- Open report ---
  await page1.getByRole('button', { name: 'Open Commitments Not Received' }).click();

  // Purchasing Organization — Select all
  await page1.getByTestId('collapsed-title-Purchasing Organization').locator('div').filter({ hasText: /^Purchasing Organization$/ }).click();
  await page1.getByTestId('actions-toolbar-more').click();
  await page1.getByText('Select all').click();
  await page1.getByTestId('actions-toolbar-confirm').click();

  // Document Type — Select all
  await page1.getByRole('heading', { name: 'Document Type' }).click();
  await page1.getByTestId('actions-toolbar-more').click();
  await page1.getByText('Select all').click();
  await page1.getByTestId('actions-toolbar-confirm').click();

  // Purchasing Group — Select all
  await page1.getByRole('heading', { name: 'Purchasing Group' }).click();
  await page1.getByTestId('actions-toolbar-more').click();
  await page1.getByRole('menuitem', { name: 'Select all' }).click();
  await page1.getByTestId('actions-toolbar-confirm').click();

  // Filter by buyer name
  await page1.locator('th:nth-child(44) > .lui-icon').click();
  await page1.getByTestId('search-input-field').click();
  await page1.getByTestId('search-input-field').fill('sean mcgrath');
  await page1.locator('span').filter({ hasText: 'SEAN MCGRATH' }).first().click();
  await page1.getByTestId('actions-toolbar-confirm').click();

  // --- Export as xlsx ---
  await page1.getByRole('columnheader', { name: 'Purchasing Group. Press space' }).click({ button: 'right' });
  await page1.locator('.qv-objectmenu-item-container > .lui-icon').first().click();
  await page1.getByText('Download as...').click();
  await page1.getByText('Data', { exact: true }).click();
  await page1.getByRole('button', { name: 'Export' }).click();

  // --- Download & rename ---
  const today = new Date();
  const dateStr = `${today.getMonth() + 1}_${today.getDate()}_${String(today.getFullYear()).slice(-2)}`;
  const filename = `OCNR_${dateStr}.xlsx`;
  const destFolder = 'C:\\Users\\mcgr1se\\Downloads\\OCNR';
  const destPath = path.join(destFolder, filename);

  if (!fs.existsSync(destFolder)) fs.mkdirSync(destFolder, { recursive: true });

  const page2Promise = page1.waitForEvent('popup');
  const downloadPromise = page1.waitForEvent('download');
  await page1.getByRole('link', { name: 'Click here to download your' }).click();
  const page2 = await page2Promise;
  const download = await downloadPromise;
  await page2.close();
  await download.saveAs(destPath);

  await page1.getByRole('button', { name: 'Close' }).click();

  console.log(`\nSaved: ${destPath}`);
});
