import { test, expect, type ConsoleMessage } from '@playwright/test';

// These tests run against a real production build (npm run start). They guard
// the things tsc + next build cannot catch: hydration mismatches, runtime
// console errors, broken initial render, mobile overflow.

test.describe('Home page smoke', () => {
  test('renders without console / page errors', async ({ page }) => {
    const errors: string[] = [];
    page.on('console', (msg: ConsoleMessage) => {
      if (msg.type() === 'error') errors.push(`Console: ${msg.text()}`);
    });
    page.on('pageerror', (err) => {
      errors.push(`PageError: ${err.message}`);
    });

    await page.goto('/');
    await expect(page.locator('main.page-shell')).toBeVisible();
    await page.waitForLoadState('networkidle');

    // Filter out 3rd-party noise (favicon 404 in dev etc.).
    const real = errors.filter((e) => !/favicon|sourcemap/i.test(e));
    expect(real, `Errors detected:\n${real.join('\n')}`).toEqual([]);
  });

  test('uploader area is mounted with a file input', async ({ page }) => {
    await page.goto('/');
    // The hidden file input that ExcelUploader controls.
    const fileInput = page.locator('input[type="file"]').first();
    await expect(fileInput).toBeAttached();
    // It accepts Excel-ish mime types.
    const accept = await fileInput.getAttribute('accept');
    expect(accept ?? '').toMatch(/xlsx|excel|spreadsheet|\*/i);
  });

  test('no horizontal overflow at the current viewport', async ({ page }) => {
    await page.goto('/');
    await page.waitForLoadState('networkidle');
    const { scroll, inner } = await page.evaluate(() => ({
      scroll: document.documentElement.scrollWidth,
      inner: window.innerWidth,
    }));
    // Allow 1px rounding tolerance.
    expect(scroll, `scrollWidth=${scroll} > innerWidth=${inner}`).toBeLessThanOrEqual(inner + 1);
  });
});
