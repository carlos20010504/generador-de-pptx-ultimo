import { chromium } from 'playwright-core';
import { fileURLToPath } from 'url';
import path from 'path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fileUrl = `file:///${path.join(__dirname, 'lightbox_test.html').replace(/\\/g, '/')}`;

const browser = await chromium.launch({ headless: true });
const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
const page = await ctx.newPage();
await page.goto(fileUrl);
await page.waitForLoadState('networkidle');

// Screenshot modo medium
await page.screenshot({ path: path.join(__dirname, 'lightbox_medium.png'), fullPage: false });

// Click Ampliar
await page.click('#btn-zoom');
await page.waitForTimeout(400);

// Screenshot modo big (sin scroll)
await page.screenshot({ path: path.join(__dirname, 'lightbox_big_top.png'), fullPage: false });

// Scroll abajo a la mitad
await page.evaluate(() => window.__test.scroll(300, 400));
await page.waitForTimeout(200);

// Screenshot modo big con scroll aplicado (verifica que toolbar siga sticky)
await page.screenshot({ path: path.join(__dirname, 'lightbox_big_scrolled.png'), fullPage: false });

await browser.close();
console.log('Screenshots saved to __tests__/stress_fixtures/');
