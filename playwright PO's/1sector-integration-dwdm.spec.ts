import { test, Page } from '@playwright/test';

const PROJECT_NUM = process.env.PROJECT_NUM ?? '';
const DUE_DATE    = process.env.DUE_DATE    ?? '';  // MM/DD/YYYY

async function pickDate(page: Page, mmddyyyy: string) {
  const [month, day, year] = mmddyyyy.split('/').map(Number);
  const target  = new Date(year, month - 1, 1);
  const now     = new Date();
  const current = new Date(now.getFullYear(), now.getMonth(), 1);
  const monthsAhead = Math.max(
    0,
    (target.getFullYear() - current.getFullYear()) * 12 + (target.getMonth() - current.getMonth())
  );

  for (let i = 0; i < monthsAhead; i++) {
    await page.getByRole('button', { name: 'Next' }).click();
  }

  await page.getByText(String(day), { exact: true }).first().click();
}

test('Create Service PO — 1 Sector Integration DWDM', async ({ page }) => {
  test.setTimeout(180000);

  if (!PROJECT_NUM) throw new Error('Missing PROJECT_NUM — run via: npm run 1sector-integration-dwdm');
  if (!DUE_DATE)    throw new Error('Missing DUE_DATE — run via: npm run 1sector-integration-dwdm');

  // --- Login ---
  await page.goto('https://vz1erpwp1.verizon.com/sap/bc/ui5_ui5/sap/zmim_ord_dash/index.html?sap-theme=custom_horizon@/sap/public/bc/themes/~client-400&sap-client=400#');
  await page.getByRole('textbox', { name: 'Enter your User Name' }).fill('mcgr1se');
  await page.getByRole('textbox', { name: 'Enter your Password' }).fill(process.env.VZ_PASSWORD ?? '');
  await page.getByRole('button', { name: 'Log in' }).click();

  await page.getByRole('searchbox', { name: 'Search Order Number, SPM ID,' }).waitFor({ timeout: 60000 });

  // --- New Purchase Order ---
  await page.getByRole('button', { name: 'Start New Order ' }).click();
  await page.getByText('New Purchase Order').click();

  // --- Search for project ---
  await page.getByRole('searchbox', { name: 'Search SPM ID, Site ID, Site' }).fill(PROJECT_NUM);
  const spmCells = page.getByRole('gridcell', { name: 'SPM ID Fixed Column' });
  await spmCells.first().waitFor({ timeout: 15000 });
  const cellCount = await spmCells.count();
  let clickIndex = 0;
  let nonEmptyIdx = 0;
  let found = false;
  for (let i = 0; i < cellCount; i++) {
    const text = (await spmCells.nth(i).textContent() ?? '').trim();
    if (!text) continue;
    if (text === PROJECT_NUM) { clickIndex = nonEmptyIdx; found = true; break; }
    nonEmptyIdx++;
  }
  if (!found) throw new Error(`SPM ID ${PROJECT_NUM} not found in search results`);
  await page.getByRole('gridcell', { name: 'Click to Select' }).nth(clickIndex).click();
  await page.getByRole('button', { name: 'Save Selection to Start Order' }).click();

  // --- Requisitioner and due date ---
  await page.getByRole('textbox', { name: 'Requisitioner ID' }).fill('sean mcgrath');
  await page.getByText('4546987335').click();
  await page.getByRole('button', { name: 'Open Picker' }).click();
  await pickDate(page, DUE_DATE);

  // --- Select SPM row and start order ---
  await page.locator('[id="__xmlview2--idPOTabSelSPM-rowsel0"]').click();
  await page.getByRole('button', { name: 'Start Order(All)' }).click();

  // --- Load 1 Sector DWDM template ---
  await page.getByText('Templates').click();
  await page.getByRole('row').filter({ hasText: /E\/\/\/ Install.*5G.*1 sector.*DWDM/i })
    .locator('button[id*="idEditTempBtn"]')
    .click();

  // Select $2,200 (index 1) — skip index 0 which is $1,651
  await page.getByRole('gridcell', { name: 'Click to Select' }).nth(1).click();
  // After first click SAP renames all selection cells — wait for selection to register,
  // then click $950 which is now the last unselected cell matching /click to select/i
  await page.getByRole('row', { selected: true }).first().waitFor({ timeout: 5000 });
  await page.getByRole('gridcell', { name: /click to select/i }).last().click();
  await page.getByRole('button', { name: 'Import To Order' }).click();

  // Wait for imported items panel, select all 2, then add to cart
  await page.getByRole('heading', { name: /Imported Line Items/ }).waitFor({ timeout: 30000 });
  await page.getByRole('checkbox', { name: 'Select All' }).click();
  await page.waitForTimeout(1500);
  await page.getByRole('button', { name: 'Add To Cart' }).click();

  // --- Review cart and submit ---
  await page.getByText('Review Cart (2)').click();
  await page.getByRole('checkbox', { name: 'Select All' }).click();
  await page.getByRole('button', { name: 'Review Orders' }).click();
  await page.getByRole('button', { name: 'Proceed' }).click();
  await page.getByRole('button', { name: 'OK' }).click();
  await page.getByRole('button', { name: 'Edit' }).click();
  await page.getByRole('button', { name: 'Review Order' }).click();
  await page.getByRole('button', { name: 'Submit Order' }).click();

  // --- Capture Order ID ---
  await page.waitForLoadState('networkidle');
  const orderLink = page.getByRole('link', { name: /^[A-Z]\d{9}$/ }).first();
  const orderText = page.getByText(/[A-Z]\d{9}/).first();

  let orderId = '';
  try {
    await orderLink.waitFor({ timeout: 10000 });
    orderId = (await orderLink.textContent() ?? '').trim();
  } catch {
    try {
      await orderText.waitFor({ timeout: 5000 });
      const raw = (await orderText.textContent() ?? '').trim();
      const match = raw.match(/[A-Z]\d{9}/);
      orderId = match ? match[0] : raw;
    } catch {
      orderId = 'could not be extracted — check browser';
    }
  }

  console.log(`\n${'='.repeat(40)}`);
  console.log(`  Order ID: ${orderId}`);
  console.log(`${'='.repeat(40)}\n`);
});
