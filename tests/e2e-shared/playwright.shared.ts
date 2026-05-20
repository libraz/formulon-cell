import { defineConfig, devices } from '@playwright/test';

import type { DemoApp } from './types.js';

const portFromEnv = (app: DemoApp): number => {
  const appSpecificKey = `FC_E2E_${app.id.replaceAll('-', '_').toUpperCase()}_PORT`;
  const raw = process.env[appSpecificKey] ?? process.env.FC_E2E_PORT;
  if (!raw) return app.port;
  const value = Number(raw);
  if (!Number.isInteger(value) || value < 1 || value > 65_535) {
    throw new Error(`${appSpecificKey} / FC_E2E_PORT must be a TCP port number.`);
  }
  return value;
};

/**
 * Builds a Playwright config for a demo app. Both framework demos share the same
 * config surface — only the dev-server port and workspace name change.
 *
 * Browsers: Chromium + WebKit. Firefox is skipped (cf. tidy-seeking-whisper
 * plan §1.3); the engine's pthread WASM doesn't require it, and the
 * grant-permissions calls for the clipboard scenarios don't cleanly apply.
 *
 * COOP/COEP: each demo app's `vite.config.ts` already injects the headers
 * required for crossOriginIsolated, so the engine boots into WASM (not the
 * JS stub). Specs that need real recalc rely on this; see `expectNoStub`
 * in `pages/SpreadsheetPage.ts`.
 */
export function defineDemoAppConfig(app: DemoApp) {
  const port = portFromEnv(app);
  const baseURL = `http://127.0.0.1:${port}`;
  const reuseExistingServer = process.env.FC_REUSE_E2E_SERVER === '1' && !process.env.CI;
  return defineConfig({
    // Playwright resolves testDir relative to the consuming config file
    // (apps/<id>/playwright.config.ts), so `./e2e` points at each demo's
    // own spec directory.
    testDir: './e2e',
    // Visual regression specs live under `e2e/visual/` and require a Linux
    // baseline. Normal e2e runs skip them; opt in with FC_VISUAL=1.
    testIgnore: process.env.FC_VISUAL ? [] : ['**/visual/**'],
    fullyParallel: true,
    forbidOnly: !!process.env.CI,
    retries: process.env.CI ? 2 : 0,
    workers: process.env.CI ? 1 : undefined,
    reporter: process.env.CI ? [['github'], ['list']] : 'list',
    use: {
      baseURL,
      trace: 'on-first-retry',
      // Stable viewport so canvas hit-tests / visual diffs don't shift.
      viewport: { width: 1280, height: 800 },
      // Mod+C/X/V are routed through navigator.clipboard because the host is
      // a non-editable, user-select:none div and browsers never fire native
      // copy/paste events on it. WebKit doesn't honor the permission, so
      // the project-level override below clears it for that engine.
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
      command: `yarn workspace ${app.workspace} dev --host 127.0.0.1 --port ${port} --strictPort`,
      url: baseURL,
      // Reuse only when explicitly requested. Otherwise Playwright can attach
      // to an unrelated app already listening on the demo port and report
      // misleading UI failures before the spreadsheet even mounts.
      reuseExistingServer,
      stdout: 'pipe',
      stderr: 'pipe',
      timeout: 120_000,
    },
  });
}
