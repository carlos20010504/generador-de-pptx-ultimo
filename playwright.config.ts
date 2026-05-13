import { defineConfig, devices } from '@playwright/test';

const PORT = 3001;
const BASE_URL = process.env.PLAYWRIGHT_BASE_URL || `http://localhost:${PORT}`;

export default defineConfig({
  testDir: './tests/e2e',
  timeout: 30_000,
  expect: { timeout: 5_000 },
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  // Reintenta una vez localmente también — con 4 workers en paralelo el pre-warm
  // de la página dispara N llamadas concurrentes a /api/health que spawnean Python
  // simultáneamente. En máquinas con poca CPU eso puede causar timeouts intermitentes.
  retries: 1,
  // Cap de 2 workers locales (CI ya usa 1). Evita la tormenta de Python spawns
  // concurrentes del pre-warm que hacían tests flakeys con default workers=4.
  workers: process.env.CI ? 1 : 2,
  reporter: process.env.CI ? [['github'], ['list']] : 'list',
  use: {
    baseURL: BASE_URL,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  projects: [
    { name: 'chromium-desktop', use: { ...devices['Desktop Chrome'] } },
    {
      // iPhone 13 viewport but on Chromium (we don't ship for WebKit and
      // installing WebKit doubles CI time).
      name: 'chromium-mobile',
      use: {
        ...devices['Desktop Chrome'],
        viewport: { width: 390, height: 844 },
        deviceScaleFactor: 3,
        isMobile: true,
        hasTouch: true,
      },
    },
  ],
  // When PLAYWRIGHT_BASE_URL is set (CI / external server), Playwright skips webServer.
  webServer: process.env.PLAYWRIGHT_BASE_URL ? undefined : {
    command: 'npm run start',
    url: BASE_URL,
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
