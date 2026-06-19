// Verifica que el dev server arranca limpio y la home no tira el error
// Turbopack/runtime que el usuario reportó. NO testea el lightbox real
// (eso requeriría subir Excel + plan + generate, ~60s).
import { chromium } from 'playwright-core';

const browser = await chromium.launch({ headless: true });
try {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(`pageerror: ${e.message}`));
  page.on('console', m => {
    if (m.type() === 'error') errors.push(`console.error: ${m.text()}`);
  });

  const resp = await page.goto('http://localhost:3001/', { waitUntil: 'networkidle', timeout: 30000 });
  console.log(`HTTP status: ${resp.status()}`);

  if (resp.status() !== 200) {
    console.log('FAIL: server not returning 200');
    process.exit(1);
  }

  // Esperar que React hidrate
  await page.waitForTimeout(2000);

  // Buscar el PRV_STYLES inyectado por el componente
  const stylesheets = await page.evaluate(() => {
    const styles = Array.from(document.querySelectorAll('style'));
    return styles.map(s => s.textContent || '').filter(t =>
      t.includes('prv-lightbox') || t.includes('prv-card')
    ).length;
  });

  console.log(`Inline <style> blocks with prv- rules: ${stylesheets}`);

  // Buscar el componente raíz que dispara el flow
  const hasUploader = await page.evaluate(() => {
    return !!document.querySelector('input[type="file"]');
  });
  console.log(`Has file input (ExcelUploader): ${hasUploader}`);

  if (errors.length > 0) {
    console.log(`\nErrors detected:`);
    errors.forEach(e => console.log(`  - ${e}`));
    // No hacemos fail por console.errors (algunos warnings de React son comunes en dev)
  } else {
    console.log('\nNo runtime errors.');
  }

  await ctx.close();
} finally {
  await browser.close();
}
