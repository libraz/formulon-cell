import { defineConfig, devices } from '@playwright/test';

/**
 * Visual-regression config. Runs ONLY `e2e/visual/**` specs and ONLY on
 * Chromium (WebKit's font rasterisation drifts too much to keep meaningful
 * snapshots without per-browser baselines we don't yet maintain).
 *
 * Baselines live in `e2e/visual/<spec>.spec.ts-snapshots/`. Generate them on
 * Linux — `yarn workspace @formulon-cell/playground exec playwright test
 * --config playwright.visual.config.ts --update-snapshots` from the CI image
 * is the canonical command.
 */
const baseURL = 'http://127.0.0.1:5173';

export default defineConfig({
  testDir: './e2e/visual',
  fullyParallel: false,
  workers: 1,
  forbidOnly: !!process.env.CI,
  retries: 0,
  reporter: process.env.CI ? [['github'], ['list']] : 'list',
  use: {
    baseURL,
    viewport: { width: 1280, height: 800 },
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
  webServer: {
    command: 'yarn workspace @formulon-cell/playground dev --host 127.0.0.1 --port 5173',
    url: baseURL,
    reuseExistingServer: !process.env.CI,
    stdout: 'pipe',
    stderr: 'pipe',
    timeout: 120_000,
  },
});
