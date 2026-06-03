import { test } from '@playwright/test';

const PROJECT_ID = process.env.FUZE_PROJECT_ID ?? '';

// Suppress Chrome's "wants to access other devices on your local network" prompt
test.use({
  launchOptions: {
    args: [
      '--disable-features=PrivateNetworkAccessPermissionPrompt,PrivateNetworkAccessSendPreflights,BlockInsecurePrivateNetworkRequests',
      '--disable-web-security',
      '--allow-insecure-localhost',
    ],
  },
});

test('Shelter BOM', async ({ page }) => {
  test.setTimeout(180000);

  if (!PROJECT_ID) {
    throw new Error('Missing project ID — run via: npm run shelter-bom');
  }

  // FORECAST_DATE env var is an optional manual override
  let TARGET_DATE = process.env.FORECAST_DATE ?? '';

  // --- Login ---
  await page.goto(`https://fuze.verizon.com/spm/projects.jsp?projectId=${PROJECT_ID}`);
  await page.getByRole('textbox', { name: 'Enter your User Name' }).fill('mcgr1se');
  await page.getByRole('textbox', { name: 'Enter your Password' }).fill(process.env.VZ_PASSWORD ?? '');
  await page.getByRole('button', { name: 'Log in' }).click();

  // Wait for SSO to complete and redirect back to fuze.verizon.com before navigating
  await page.waitForURL('**/fuze.verizon.com/**', { timeout: 60000 });
  await page.waitForLoadState('load');
  await page.goto(`https://fuze.verizon.com/spm/projects.jsp?projectId=${PROJECT_ID}`);
  await page.waitForLoadState('load');

  // Auto-calculate forecast date from "Construction Commenced → Construction Started (F)"
  // on the Overview tab. Uses waitForFunction to poll textContent directly, bypassing
  // aria-hidden and CSS visibility which block getByText from finding the section.
  if (!TARGET_DATE) {
    // Two nested tabpanels must be opened to reach "Construction Started":
    //   Construction Milestone → Construction Commenced → Construction Started row
    const constMilestoneBtn = page.getByRole('button', { name: /^Construction Milestone\b/ }).first();
    await constMilestoneBtn.waitFor({ timeout: 15000 });
    await constMilestoneBtn.click();

    const commencedBtn = page.getByRole('button', { name: /Construction Commenced\b/ }).first();
    await commencedBtn.waitFor({ timeout: 30000 });
    await commencedBtn.scrollIntoViewIfNeeded();
    await commencedBtn.click();

    // Wait for the Construction Started row to appear, then read the F date.
    // Anchor on the "Construction Started" label and search downward for the date —
    // more reliable than traversing upward from the textbox.
    await page.getByText('Construction Started', { exact: true }).first().waitFor({ timeout: 15000 });

    const constStart = await page.evaluate((): string | null => {
      const dateRe = /^\d{2}\/\d{2}\/\d{4}$/;
      for (const el of document.querySelectorAll('*')) {
        if (el.textContent?.trim() !== 'Construction Started') continue;
        // Found the label — look for a date in the parent container (the row)
        const parent = el.parentElement;
        if (!parent) continue;
        for (const tb of parent.querySelectorAll('input, [role="textbox"]')) {
          const v = ((tb as HTMLInputElement).value || tb.textContent || '').trim();
          if (dateRe.test(v)) return v;
        }
      }
      return null;
    }) as string | null;

    if (!constStart) throw new Error(
      'Construction Start (F) date not found in the Construction Commenced section.\n' +
      'Override: $env:FORECAST_DATE="MM/DD/YYYY"; npm run shelter-bom'
    );

    const [cm, cd, cy] = constStart.split('/').map(Number);
    const forecast = new Date(cy, cm - 1, cd);
    forecast.setDate(forecast.getDate() - 30);
    const today = new Date(); today.setHours(0, 0, 0, 0);
    if (forecast < today) throw new Error(
      `Construction Start (F) is ${constStart} — forecast date is in the past.\n` +
      'Override: $env:FORECAST_DATE="MM/DD/YYYY"; npm run shelter-bom'
    );

    TARGET_DATE = `${String(forecast.getMonth() + 1).padStart(2, '0')}/${String(forecast.getDate()).padStart(2, '0')}/${forecast.getFullYear()}`;
    console.log(`  Construction Start (F): ${constStart}  →  Forecast Date: ${TARGET_DATE}`);
  }

  // Set up PNA interceptor AFTER login so it doesn't interfere with the SSO flow
  await page.route('**/*', async (route) => {
    if (route.request().method() === 'OPTIONS') {
      await route.fulfill({
        status: 200,
        headers: {
          'Access-Control-Allow-Origin': 'https://fuze.verizon.com',
          'Access-Control-Allow-Private-Network': 'true',
          'Access-Control-Allow-Methods': '*',
          'Access-Control-Allow-Headers': '*',
        },
      });
    } else {
      await route.continue();
    }
  });

  // Wait for project page to load
  await page.getByRole('button', { name: /Equipment Milestone/ }).waitFor({ timeout: 60000 });

  // --- Navigate to BOM ---
  await page.getByRole('button', { name: /Equipment Milestone/ }).click();
  await page.getByRole('button', { name: /GENERATE BOM/ }).click();

  // The RFDS BOM URL redirects to the project page when loaded directly — it must be
  // opened as a popup from the project page to carry the correct session context.
  const page1Promise = page.waitForEvent('popup');
  await page.getByRole('link', { name: 'Navigate To RFDS BOM' }).click();
  const page1 = await page1Promise;

  // Set up PNA interceptor BEFORE waiting for load so preflight requests are handled
  await page1.route('**/*', async (route) => {
    if (route.request().method() === 'OPTIONS') {
      await route.fulfill({
        status: 200,
        headers: {
          'Access-Control-Allow-Origin': 'https://fuze.verizon.com',
          'Access-Control-Allow-Private-Network': 'true',
          'Access-Control-Allow-Methods': '*',
          'Access-Control-Allow-Headers': '*',
        },
      });
    } else {
      await route.continue();
    }
  });
  await page1.waitForLoadState('domcontentloaded');

  // --- Import Template ---
  await page1.getByRole('button', { name: 'Import Template' }).click();
  await page1.locator('tr:nth-child(8) > .importTemplateSelector > .k-checkbox-label').first().click();
  await page1.getByRole('button', { name: 'Continue' }).click();
  // "Select all" checkbox in the Kendo grid header — stable selector that doesn't use the generated GUID
  await page1.locator('#template-modal th .k-checkbox-label:visible').first().click({ timeout: 15000 });
  await page1.getByRole('button', { name: 'Import Equipment Records' }).click();

  // Some projects open a cart-mapping modal after import — click through it if it appears
  try {
    await page1.locator('#equip-cart-map-modal').waitFor({ state: 'visible', timeout: 8000 });
    await page1.getByRole('button', { name: 'Map Cart and Import' }).click();
    await page1.locator('#equip-cart-map-modal').waitFor({ state: 'hidden', timeout: 30000 });
  } catch {
    // Modal didn't appear — this project doesn't require cart mapping
  }

  // --- EE Submit ---
  await page1.getByRole('button', { name: 'EE Submit' }).click();
  await page1.locator('#shelterBomCheckbox').check();

  // Fill the Shelter Equipment Need By date — type character by character to fire
  // all keyboard events the field's validation listener expects
  const needByInput = page1.getByText('Shelter Equipment Need By')
    .locator('xpath=following::input[@type="text" and not(@disabled)][1]');
  await needByInput.waitFor({ state: 'visible', timeout: 10000 });
  await needByInput.click();
  await needByInput.pressSequentially(TARGET_DATE, { delay: 50 });
  await needByInput.press('Tab');

  // Download BOM
  const downloadPromise = page1.waitForEvent('download');
  await page1.locator('#eesubmitPost').click();
  const download = await downloadPromise;

  console.log(`\nBOM downloaded: ${download.suggestedFilename()}`);
});
