import { test, expect } from '@playwright/test';
import path from 'path';
import fs from 'fs';

const COMISIONES_PATH = path.resolve(__dirname, '../fixtures/comisiones.xlsx');
const HAS_FIXTURE = fs.existsSync(COMISIONES_PATH);

test.describe('Prompt obedience UI', () => {
  test.skip(!HAS_FIXTURE, 'No fixture: copy a comisiones-style xlsx to tests/fixtures/comisiones.xlsx');

  test('user can write prompt in PreparePanel and see intent banner', async ({ page }) => {
    await page.goto('/');
    await page.waitForLoadState('networkidle');

    const fileInput = page.locator('input[type="file"]').first();
    await fileInput.setInputFiles(COMISIONES_PATH);

    // PreparePanel appears with the new label
    await expect(page.getByText('Tu instrucción para la IA')).toBeVisible({ timeout: 30_000 });

    // Refinar opens by default — placeholder visible immediately
    const promptArea = page.getByPlaceholder(/9 slides incluyendo Riesgos/i);
    await expect(promptArea).toBeVisible();

    // Type the prompt
    await promptArea.fill('hazme 9 slides con Riesgos Core');

    // Wait for the intent banner (text "Detecté")
    await expect(page.getByText(/Detecté/i)).toBeVisible({ timeout: 240_000 });
  });

  test('panel-avanzado drawer has no prompt textarea', async ({ page }) => {
    await page.goto('/');
    await page.waitForLoadState('networkidle');

    const fileInput = page.locator('input[type="file"]').first();
    await fileInput.setInputFiles(COMISIONES_PATH);
    await expect(page.getByText('Tu instrucción para la IA')).toBeVisible({ timeout: 30_000 });

    // Open the renamed avanzado drawer
    await page.getByRole('button', { name: /Audiencia y tema/i }).click();

    // Drawer is open. Verify NO prompt textarea inside.
    // The placeholder we used in PreparePanel is "Ej: hazme 9 slides…"
    // — we want to confirm the drawer doesn't contain that placeholder either.
    const drawer = page.locator('[role="dialog"]').first();
    await expect(drawer).toBeVisible();
    // The prompt placeholder we added to PreparePanel is unique enough
    // to verify only ONE textarea with it exists in the page (not in the drawer).
    const promptCount = await page.getByPlaceholder(/9 slides incluyendo Riesgos/i).count();
    expect(promptCount).toBe(1);
  });
});
