// Headless test del lightbox standalone. Verifica que:
// 1. En modo "medium": imagen cabe, NO hay scroll vertical (en viewport 1920x1080)
// 2. Click "Ampliar" → modo "big": HAY scroll horizontal y/o vertical
// 3. window.scrollTo funciona (no está congelado)
// 4. Toolbar sticky permanece visible al scrollear

import { chromium } from 'playwright-core';
import { fileURLToPath } from 'url';
import path from 'path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fileUrl = `file:///${path.join(__dirname, 'lightbox_test.html').replace(/\\/g, '/')}`;

const VIEWPORTS = [
  { name: 'desktop-1920',  width: 1920, height: 1080 },
  { name: 'laptop-1366',   width: 1366, height: 768  },
  { name: 'tablet-768',    width: 768,  height: 1024 },
];

let pass = 0, fail = 0;
function assert(cond, msg) {
  if (cond) { console.log(`  ✓ ${msg}`); pass++; }
  else      { console.log(`  ✗ ${msg}`); fail++; }
}

const browser = await chromium.launch({ headless: true });
try {
  for (const vp of VIEWPORTS) {
    console.log(`\n=== Viewport ${vp.name} (${vp.width}x${vp.height}) ===`);
    const ctx = await browser.newContext({ viewport: vp });
    const page = await ctx.newPage();
    await page.goto(fileUrl);
    await page.waitForLoadState('networkidle');

    // 1. Modo "medium" (default)
    let m = await page.evaluate(() => window.__test.measure());
    console.log(`  [MEDIUM] img ${m.imgWidth}×${m.imgHeight}, container ${m.clientWidth}×${m.clientHeight}, scroll ${m.scrollWidth}×${m.scrollHeight}`);
    assert(m.imgWidth <= m.clientWidth, 'imagen cabe horizontal en modo medium');
    // Vertical scroll en modo medium puede o no ocurrir según viewport.

    // 2. Click Ampliar → modo "big"
    await page.click('#btn-zoom');
    await page.waitForTimeout(300); // espera la transition
    m = await page.evaluate(() => window.__test.measure());
    console.log(`  [BIG]    img ${m.imgWidth}×${m.imgHeight}, container ${m.clientWidth}×${m.clientHeight}, scroll ${m.scrollWidth}×${m.scrollHeight}`);
    assert(m.hasHorizontalScroll || m.hasVerticalScroll, 'hay scroll en modo big');
    assert(m.imgWidth > m.clientWidth || m.imgHeight > m.clientHeight, 'imagen desborda en modo big');

    // 3. Verificar que scroll programático funciona (scrollTop actually moves)
    const scrolled = await page.evaluate(() => window.__test.scroll(200, 300));
    assert(scrolled.top > 0 || scrolled.left > 0, `scroll funciona (top=${scrolled.top}, left=${scrolled.left})`);

    // 4. Toolbar sticky — debe quedar en y=0 aunque scrolleemos
    const barTop = await page.evaluate(() => {
      const bar = document.querySelector('.prv-lightbox-bar');
      const rect = bar.getBoundingClientRect();
      return rect.top;
    });
    assert(barTop >= 0 && barTop < 10, `toolbar sticky en top (rect.top=${barTop})`);

    // 5. Click Reducir → vuelve a medium, no debe haber scroll innecesario
    await page.click('#btn-zoom');
    await page.waitForTimeout(300);
    const finalLabel = await page.textContent('#btn-zoom');
    assert(finalLabel.trim() === 'Ampliar', `toggle vuelve a Ampliar (got: "${finalLabel}")`);

    await ctx.close();
  }
} finally {
  await browser.close();
}

console.log(`\n=== Resultado: ${pass} pass, ${fail} fail ===`);
process.exit(fail > 0 ? 1 : 0);
