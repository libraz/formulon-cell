import { defineConfig, devices } from '@playwright/test';

import { ADAPTER_SPECS, PRIMARY_DEMO_APP, runFullE2E } from './adapter-specs.js';
import { SYSTEM_CLIPBOARD } from './tags.js';
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

const SYSTEM_CLIPBOARD_GREP = new RegExp(SYSTEM_CLIPBOARD);

const BROWSER_PROJECTS = [
  { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
  { name: 'webkit', use: { ...devices['Desktop Safari'], permissions: [] } },
];

/**
 * Builds a Playwright config for a demo app. Both framework demos share the same
 * config surface — only the dev-server port and workspace name change.
 *
 * Browsers: Chromium + WebKit. Firefox is skipped: the engine's WASM doesn't
 * require it, and the grant-permissions calls for the clipboard scenarios
 * don't cleanly apply.
 *
 * The default serial WASM loads without COOP/COEP headers. Specs that need
 * real recalc verify WASM mode via `expectNoStub` in pages/SpreadsheetPage.ts.
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
    // Visual regression specs live under `e2e/visual/` and compare against macOS baselines kept
    // only on the release machine. Normal e2e runs skip them; opt in with FC_VISUAL=1.
    testIgnore: process.env.FC_VISUAL ? [] : ['**/visual/**'],
    // A missing baseline fails without being written; only an explicit --update-snapshots run
    // takes one.
    updateSnapshots: 'none',
    // A secondary app runs only the adapter-specific specs unless the full run or the visual walk
    // is requested.
    ...(app.id === PRIMARY_DEMO_APP || runFullE2E() || process.env.FC_VISUAL
      ? {}
      : { testMatch: ADAPTER_SPECS.map(({ spec }) => `**/${spec}`) }),
    fullyParallel: true,
    forbidOnly: !!process.env.CI,
    retries: process.env.CI ? 2 : 0,
    // Two workers: local machines are shared, and CI runners have four vCPUs.
    workers: 2,
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
    // Each browser has a `<browser>-clipboard` project that runs the system-clipboard tests one at a
    // time; select both with `--project='chromium*'`.
    projects: BROWSER_PROJECTS.flatMap((project) => [
      { ...project, grepInvert: SYSTEM_CLIPBOARD_GREP },
      { ...project, name: `${project.name}-clipboard`, grep: SYSTEM_CLIPBOARD_GREP, workers: 1 },
    ]),
    // A production build served by `vite preview`, so edits to the sources during a run cannot
    // hot-update the app under test. FC_E2E_DEV=1 serves the dev server instead, for debugging.
    webServer: {
      command:
        process.env.FC_E2E_DEV === '1'
          ? `yarn workspace ${app.workspace} dev --host 127.0.0.1 --port ${port} --strictPort`
          : `yarn workspace ${app.workspace} build && yarn workspace ${app.workspace} preview --host 127.0.0.1 --port ${port} --strictPort`,
      url: baseURL,
      // Reuse only when explicitly requested. Otherwise Playwright can attach
      // to an unrelated app already listening on the demo port and report
      // misleading UI failures before the spreadsheet even mounts.
      reuseExistingServer,
      stdout: 'pipe',
      stderr: 'pipe',
      // Covers the production build that precedes `vite preview`.
      timeout: 300_000,
    },
  });
}
