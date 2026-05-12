// Capture the new cleaned-up UI flow:
//   1. Empty state
//   2. After upload — PreparePanel loaded (detection + plan)
//   3. PreparePanel with Refinar expanded
//   4. AdvancedDrawer open
import { chromium, devices } from 'playwright-core';
import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';

const URL = 'http://localhost:3001/';
const EXCEL = 'Comisiones V1.xlsx';
const OUTDIR = resolve('C:\\Users\\cpinzon\\AppData\\Local\\Temp\\ui_cleanup');

await mkdir(OUTDIR, { recursive: true });
const browser = await chromium.launch({ headless: true });

async function captureFlow(deviceCfg, prefix) {
  const ctx = await browser.newContext(deviceCfg);
  const page = await ctx.newPage();
  page.on('console', msg => {
    if (msg.type() === 'error') console.error(`  [${prefix}] console.error:`, msg.text());
  });
  page.on('pageerror', err => console.error(`  [${prefix}] pageerror:`, err.message));

  await page.goto(URL, { waitUntil: 'networkidle' });
  await page.waitForTimeout(400);

  // 1. Empty state — viewport-clip (not fullPage) so off-screen drawer
  // doesn't leak into the capture. Drawer is position:fixed.
  await page.screenshot({ path: resolve(OUTDIR, `${prefix}-1-empty.png`), fullPage: false });
  console.log(`  ${prefix}-1-empty → captured`);

  // Upload Excel
  const fi = page.locator('input[type="file"]').first();
  await fi.setInputFiles(EXCEL);

  // Wait for the PreparePanel detection grid (first signal)
  await page.waitForSelector('.prep-stats, .prep-banner.is-warn', { timeout: 60_000 });
  await page.waitForTimeout(1500);

  // 2. After upload — both detection + plan ideally loaded
  // Wait for slides list OR plan error (don't block forever on AI)
  await Promise.race([
    page.waitForSelector('.prep-slides', { timeout: 240_000 }),
    page.waitForSelector('.prep-banner.is-error', { timeout: 240_000 }),
  ]).catch(() => { /* ignore — capture whatever state we have */ });
  await page.waitForTimeout(800);
  await page.screenshot({ path: resolve(OUTDIR, `${prefix}-2-loaded.png`), fullPage: false });
  console.log(`  ${prefix}-2-loaded → captured`);

  // 3. Open Refinar
  await page.locator('.prep-refine-toggle').click();
  await page.waitForTimeout(500);
  await page.screenshot({ path: resolve(OUTDIR, `${prefix}-3-refine.png`), fullPage: false });
  console.log(`  ${prefix}-3-refine → captured`);

  // 4. Open Advanced drawer
  await page.locator('.upl-adv-trigger').click();
  await page.waitForSelector('.adv-drawer.is-open', { timeout: 5_000 });
  await page.waitForTimeout(800);
  await page.screenshot({ path: resolve(OUTDIR, `${prefix}-4-drawer.png`), fullPage: false });
  console.log(`  ${prefix}-4-drawer → captured`);

  // 5. Toggle off some slides — close drawer first via ESC (avoids overlap)
  await page.keyboard.press('Escape');
  await page.waitForTimeout(400);
  const toggles = page.locator('.prep-slide:not(.is-mandatory)');
  const count = await toggles.count();
  if (count >= 2) {
    await toggles.nth(1).click();
    await page.waitForTimeout(200);
    await toggles.nth(3).click();
    await page.waitForTimeout(400);
    await page.screenshot({ path: resolve(OUTDIR, `${prefix}-5-toggled.png`), fullPage: false });
    console.log(`  ${prefix}-5-toggled → captured`);
  }

  await ctx.close();
}

try {
  await captureFlow(
    { viewport: { width: 1440, height: 1100 }, deviceScaleFactor: 1 },
    'desktop'
  );
  await captureFlow(
    { ...devices['iPhone 13'], viewport: { width: 390, height: 1000 } },
    'mobile'
  );
} finally {
  await browser.close();
}
console.log('Done.');
