import { test, expect, Page } from '@playwright/test';

// Source EO to clone from
const SOURCE_EO = 'E000274573';

// Target date for all cloned EOs (MM/DD/YYYY and YYYYMMDD for calendar cell ID)
const TARGET_DATE = '07/17/2026';
const TARGET_DATE_ID = '20260717';

// SPM Project IDs — provided at runtime (1–20 IDs)
const SPM_COUNT = parseInt(process.env.SPM_COUNT ?? '0', 10);
const SPM_IDS = Array.from({ length: SPM_COUNT }, (_, i) => process.env[`SPMID${i + 1}`] ?? '');

// SAP UI5 auto-generates clone row IDs starting at 700, stepping by 16
const CLONE_BASE = 700;
const CLONE_STEP = 16;

async function submitEO(eoPage: Page) {
  await eoPage.getByRole('button', { name: 'Edit' }).click();
  await eoPage.waitForLoadState('networkidle');
  await eoPage.getByRole('button', { name: 'Review' }).click();
  await eoPage.waitForLoadState('networkidle');
  // Wait for Submit to become enabled after Review processing
  await eoPage.getByRole('button', { name: 'Submit' }).waitFor({ state: 'visible' });
  await expect(eoPage.getByRole('button', { name: 'Submit' })).toBeEnabled({ timeout: 30000 });
  await eoPage.getByRole('button', { name: 'Submit' }).click();
}

test('Clone and submit EOs - Ericsson V1 SFP', async ({ page }) => {
  test.setTimeout(180000); // 3 minutes — SSO redirect can be slow

  if (SPM_IDS.some((id) => !id)) {
    throw new Error('Missing SPM IDs — run via: node run-eo-clone.js');
  }

  // --- Login ---
  await page.goto('https://vz1erpwp1.verizon.com/sap/bc/ui5_ui5/sap/zmim_ord_dash/index.html?sap-theme=custom_horizon@/sap/public/bc/themes/~client-400&sap-client=400#');
  await page.getByRole('textbox', { name: 'Enter your User Name' }).fill('mcgr1se');
  await page.getByRole('textbox', { name: 'Enter your Password' }).fill(process.env.VZ_PASSWORD ?? '');
  await page.getByRole('button', { name: 'Log in' }).click();

  // Wait for SSO redirect to complete and the app to load
  await page.getByRole('searchbox', { name: 'Search Order Number, SPM ID,' }).waitFor({ timeout: 60000 });

  // --- Search for source EO ---
  await page.getByRole('searchbox', { name: 'Search Order Number, SPM ID,' }).fill(SOURCE_EO);
  await page.getByRole('searchbox', { name: 'Search Order Number, SPM ID,' }).press('Enter');

  const page1Promise = page.waitForEvent('popup');
  await page.getByRole('link', { name: `Order # ${SOURCE_EO}` }).click();
  const page1 = await page1Promise;

  // --- Clone ---
  await page1.bringToFront();
  await page1.waitForLoadState('networkidle');
  await page1.locator('[id="__toolbar14-overflowButton"]').waitFor({ timeout: 30000 });
  const cloneVisible = await page1.getByRole('button', { name: 'Clone' }).isVisible();
  if (!cloneVisible) {
    await page1.locator('[id="__toolbar14-overflowButton"]').click();
  }

  await page1.getByRole('button', { name: 'Clone' }).click();

  // Wait for Clone dialog to fully load
  await page1.getByRole('button', { name: 'Add SPM ID' }).waitFor({ timeout: 30000 });

  // One row exists by default — add the rest
  for (let i = 0; i < SPM_IDS.length - 1; i++) {
    await page1.getByRole('button', { name: 'Add SPM ID' }).click();
  }

  // Fill each row: SPM ID + target date
  for (let i = 0; i < SPM_IDS.length; i++) {
    const base = CLONE_BASE + i * CLONE_STEP;
    const inputSel   = `[id="__input2-__clone${base}-inner"]`;
    const dateInputSel = `[id="__picker0-__clone${base + 1}-inner"]`;

    await page1.locator(inputSel).fill(SPM_IDS[i]);

    // Type date directly into the date picker input field
    await page1.locator(dateInputSel).fill(TARGET_DATE);
    await page1.locator(dateInputSel).press('Enter');
  }

  // Check "Include GC and Shipping" for all 5 rows
  const checkboxes = await page1.getByRole('checkbox', { name: 'Include GC and Shipping' }).all();
  for (const checkbox of checkboxes) {
    await checkbox.click();
  }

  // Select DC-GC for all 5 rows
  for (let i = 0; i < SPM_IDS.length; i++) {
    const selectId = `__select3-__clone${CLONE_BASE + i * CLONE_STEP + 2}-label`;
    await page1.locator(`[id="${selectId}"]`).click();
    await page1.getByRole('option', { name: 'DC-GC' }).click();
  }

  // Submit the clone form
  await page1.getByRole('button', { name: 'Clone' }).click();

  // Wait for the cloned EO links to appear (e.g. "E000286104;")
  await page1.getByRole('link', { name: /^E\d+;?$/ }).first().waitFor({ timeout: 30000 });

  // --- Open each cloned EO and release it ---
  const eoLinks = await page1.getByRole('link', { name: /^E\d+;?$/ }).all();

  for (const link of eoLinks) {
    const popupPromise = page1.waitForEvent('popup');
    await link.click();
    const eoPage = await popupPromise;
    await eoPage.waitForLoadState('domcontentloaded');
    await submitEO(eoPage);
  }

  // --- Collect and log all created EO IDs ---
  const eoIds = await Promise.all(
    eoLinks.map(async link => (await link.textContent() ?? '').trim().replace(/;$/, ''))
  );

  console.log(`\n${'='.repeat(40)}`);
  console.log(`  Created ${eoIds.length} EO(s):`);
  eoIds.forEach(id => console.log(`    ${id}`));
  console.log(`${'='.repeat(40)}\n`);
});
