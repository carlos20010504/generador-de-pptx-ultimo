// Screenshot the UI at multiple device sizes via Playwright (Chromium).
// Usage: node scripts/screenshot-ui.mjs [url] [outdir]
import { chromium, devices } from 'playwright-core';
import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';

const URL = process.argv[2] || 'http://localhost:3001/';
const OUTDIR = resolve(process.argv[3] || 'C:\\Users\\cpinzon\\AppData\\Local\\Temp\\ui_shots\\real');

const TARGETS = [
  // name, width, height, isMobile, deviceScaleFactor, userAgent
  { name: 'mobile-iphone',  ...devices['iPhone 13'] },
  { name: 'mobile-android', ...devices['Pixel 5'] },
  { name: 'tablet-ipad',    ...devices['iPad Mini'] },
  { name: 'desktop-1280',   viewport: { width: 1280, height: 900 }, deviceScaleFactor: 1, isMobile: false, hasTouch: false },
  { name: 'desktop-1920',   viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 1, isMobile: false, hasTouch: false },
];

await mkdir(OUTDIR, { recursive: true });
const browser = await chromium.launch({ headless: true });

for (const target of TARGETS) {
  const ctx = await browser.newContext({
    viewport: target.viewport,
    deviceScaleFactor: target.deviceScaleFactor,
    isMobile: target.isMobile,
    hasTouch: target.hasTouch,
    userAgent: target.userAgent,
  });
  const page = await ctx.newPage();
  await page.goto(URL, { waitUntil: 'networkidle', timeout: 30000 });
  // Tiny wait so background animations settle into a stable frame
  await page.waitForTimeout(400);
  const out = resolve(OUTDIR, `${target.name}.png`);
  await page.screenshot({ path: out, fullPage: false });
  await ctx.close();
  console.log(`  ${target.name.padEnd(18)} ${target.viewport.width}x${target.viewport.height} → ${out}`);
}

await browser.close();
console.log('Done.');
