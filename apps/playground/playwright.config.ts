import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: './e2e',
  testIgnore: ['**/visual/**'],
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: process.env.CI ? 1 : undefined,
  reporter: process.env.CI ? [['github'], ['list']] : 'list',
  use: {
    baseURL: 'http://127.0.0.1:5173',
    trace: 'on-first-retry',
    viewport: { width: 1280, height: 800 },
    // Mod+C/X/V are routed through navigator.clipboard (the host is a
    // non-editable, user-select:none div, so browsers never fire native
    // copy/paste events on it). Pre-grant the Chromium permission so specs
    // don't need per-test grantPermissions calls. WebKit doesn't honor
    // grantPermissions for the clipboard, so the override is project-level.
    permissions: ['clipboard-read', 'clipboard-write'],
  },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    {
      name: 'webkit',
      use: { ...devices['Desktop Safari'], permissions: [] },
    },
  ],
  webServer: {
    command: 'yarn workspace @formulon-cell/playground dev --host 127.0.0.1 --port 5173',
    url: 'http://127.0.0.1:5173',
    reuseExistingServer: !process.env.CI,
    stdout: 'pipe',
    stderr: 'pipe',
    timeout: 120_000,
  },
});
