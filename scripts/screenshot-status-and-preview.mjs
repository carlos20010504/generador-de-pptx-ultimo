// Capture: status page, empty state, loaded state with PII alert.
import { chromium } from 'playwright-core';
import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';

const URL_BASE = 'http://localhost:3001';
const EXCEL = 'Comisiones V1.xlsx';
const OUTDIR = resolve('C:\\Users\\cpinzon\\AppData\\Local\\Temp\\ui_v3');

await mkdir(OUTDIR, { recursive: true });
const browser = await chromium.launch({ headless: true });

async function shoot(page, name) {
  await page.screenshot({ path: resolve(OUTDIR, `${name}.png`), fullPage: false });
  console.log(`  ${name} → captured`);
}

// Desktop: status page
{
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 900 }, deviceScaleFactor: 1,
  });
  const page = await ctx.newPage();
  await page.goto(`${URL_BASE}/status`, { waitUntil: 'domcontentloaded', timeout: 90_000 });
  await page.waitForTimeout(400);
  await shoot(page, 'desktop-status');
  await ctx.close();
}

// Desktop: empty + loaded with PII warning
{
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 900 }, deviceScaleFactor: 1,
  });
  const page = await ctx.newPage();
  page.on('console', msg => {
    if (msg.type() === 'error') console.error('  desktop console.error:', msg.text());
  });
  await page.goto(URL_BASE, { waitUntil: 'domcontentloaded', timeout: 90_000 });
  await page.waitForTimeout(400);
  await shoot(page, 'desktop-empty');

  await page.locator('input[type="file"]').first().setInputFiles(EXCEL);
  await page.waitForSelector('.prep-stats, .prep-banner.is-warn', { timeout: 60000 });
  await Promise.race([
    page.waitForSelector('.prep-slides', { timeout: 240_000 }),
    page.waitForSelector('.prep-banner.is-error', { timeout: 240_000 }),
  ]).catch(() => {});
  await page.waitForTimeout(800);
  await shoot(page, 'desktop-loaded');
  await ctx.close();
}

// Mobile: status + loaded
{
  const ctx = await browser.newContext({
    viewport: { width: 390, height: 1000 }, deviceScaleFactor: 3,
    isMobile: true, hasTouch: true,
    userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15',
  });
  const page = await ctx.newPage();
  await page.goto(`${URL_BASE}/status`, { waitUntil: 'domcontentloaded', timeout: 90_000 });
  await page.waitForTimeout(400);
  await shoot(page, 'mobile-status');

  await page.goto(URL_BASE, { waitUntil: 'domcontentloaded', timeout: 90_000 });
  await page.waitForTimeout(400);
  await page.locator('input[type="file"]').first().setInputFiles(EXCEL);
  await page.waitForSelector('.prep-stats, .prep-banner.is-warn', { timeout: 60000 });
  await Promise.race([
    page.waitForSelector('.prep-slides', { timeout: 240_000 }),
    page.waitForSelector('.prep-banner.is-error', { timeout: 240_000 }),
  ]).catch(() => {});
  await page.waitForTimeout(800);
  await shoot(page, 'mobile-loaded');
  await ctx.close();
}

await browser.close();
console.log('Done.');
