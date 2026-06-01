import { test, Page } from '@playwright/test';

const PROJECT_NUM = process.env.PROJECT_NUM ?? '';
const DUE_DATE    = process.env.DUE_DATE    ?? '';  // MM/DD/YYYY

// Navigates the SAP calendar picker to the correct month and clicks the day.
// The picker opens showing the current month, so we advance by however many
// months lie between today and the target date.
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

test('Create Service PO — 4G/5G 2 Sector', async ({ page }) => {
  test.setTimeout(180000);

  if (!PROJECT_NUM) throw new Error('Missing PROJECT_NUM — run via: node playwright-tests/run-service-po-4g5g-2sector.js');
  if (!DUE_DATE)    throw new Error('Missing DUE_DATE — run via: node playwright-tests/run-service-po-4g5g-2sector.js');

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
  await page.getByRole('gridcell', { name: 'Click to Select' }).first().click();
  await page.getByRole('button', { name: 'Save Selection to Start Order' }).click();

  // --- Requisitioner and due date ---
  await page.getByRole('textbox', { name: 'Requisitioner ID' }).fill('sean mcgrath');
  await page.getByText('4546987335').click(); // select autocomplete result
  await page.getByRole('button', { name: 'Open Picker' }).click();
  await pickDate(page, DUE_DATE);

  // --- Select SPM row and start order ---
  await page.locator('[id="__xmlview2--idPOTabSelSPM-rowsel0"]').click();
  await page.getByRole('button', { name: 'Start Order(All)' }).click();

  // --- Load 4G/5G 2 Sector template ---
  await page.getByText('Templates').click();
  await page.getByRole('row').filter({ hasText: /E\/\/\/ Install.*4G\/5G.*2 Sector/ })
    .locator('button[id*="idEditTempBtn"]')
    .click();

  // Template has 3 items — use Select All to capture all of them, then import
  await page.getByRole('checkbox', { name: 'Select All' }).click();
  await page.getByRole('button', { name: 'Import To Order' }).click();

  // Wait for imported items panel, select all 3, then add to cart
  await page.getByRole('heading', { name: /Imported Line Items/ }).waitFor({ timeout: 30000 });
  await page.getByRole('checkbox', { name: 'Select All' }).click();
  await page.waitForTimeout(1500); // SAP selection state settles before any network call fires
  await page.getByRole('button', { name: 'Add To Cart' }).click();

  // --- Review cart and submit ---
  await page.getByText('Review Cart (3)').click();
  await page.getByRole('checkbox', { name: 'Select All' }).click();
  await page.getByRole('button', { name: 'Review Orders' }).click();
  await page.getByRole('button', { name: 'Proceed' }).click();
  await page.getByRole('button', { name: 'OK' }).click();
  await page.getByRole('button', { name: 'Edit' }).click();
  await page.getByRole('button', { name: 'Review Order' }).click();
  await page.getByRole('button', { name: 'Submit Order' }).click();

  // --- Capture Order ID ---
  await page.waitForLoadState('networkidle');
  // Order ID format: letter + 9 digits, e.g. S000755208
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
