import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

/**
 * App-agnostic page object that hides the demo-app shell so the same
 * scenario can run against playground (vanilla), react-demo, and vue-demo.
 *
 * Selectors target the `.fc-host__*` chrome classes emitted by the core
 * mount, NOT the demo wrappers. That keeps the DSL robust as wrappers
 * evolve.
 */
export class SpreadsheetPage {
  constructor(public readonly page: Page) {}

  /** Open the app and wait for the engine to flip out of the "loading" state.
   *  Defaults to `?fixture=empty` so scenarios can rely on a blank workbook —
   *  the demo app's normal boot seed would otherwise leave cells under the
   *  click target and bleed into edit / undo assertions. Pass `fixture: null`
   *  to opt back into the default seed (e.g. for "with real data" scenarios). */
  async mount(opts: { fixture?: string | null } = {}): Promise<void> {
    const fixture = opts.fixture === undefined ? 'empty' : opts.fixture;
    const url = fixture ? `/?fixture=${encodeURIComponent(fixture)}` : '/';
    await this.page.goto(url);
    await this.waitForReady();
  }

  /** Wait until `[data-fc-engine-state]` settles on "ready" (real WASM) or
   *  "ready-stub" (fallback). Spec-side helpers can then assert no-stub. */
  async waitForReady(): Promise<void> {
    await this.page.waitForSelector('.fc-host', { state: 'attached', timeout: 30_000 });
    await this.page.waitForFunction(
      () => {
        const host = document.querySelector('.fc-host') as HTMLElement | null;
        const state = host?.dataset.fcEngineState;
        return state === 'ready' || state === 'ready-stub';
      },
      { timeout: 30_000 },
    );
  }

  /** Throw if the engine fell back to the JS stub — required for specs that
   *  exercise real recalc / xlsx round-trip. */
  async expectNoStub(): Promise<void> {
    const state = await this.page.evaluate(() => {
      const host = document.querySelector('.fc-host') as HTMLElement | null;
      return host?.dataset.fcEngineState ?? null;
    });
    expect(state, 'engine fell back to stub — check the demo app served COOP/COEP headers').toBe(
      'ready',
    );
  }

  /** crossOriginIsolated probe — useful for diagnostics on CI. */
  isCrossOriginIsolated(): Promise<boolean> {
    return this.page.evaluate(() => (window as Window).crossOriginIsolated === true);
  }

  /** Focus the host (canvas hit) and type into the active cell, committing
   *  with Enter. The canvas isn't typeable directly — the editor surface
   *  appears once focus is on the spreadsheet and a character is typed. */
  async typeIntoActiveCell(text: string, opts: { commit?: boolean } = {}): Promise<void> {
    await this.focusHost();
    await this.page.keyboard.type(text);
    if (opts.commit !== false) await this.page.keyboard.press('Enter');
  }

  async focusHost(): Promise<void> {
    // Click in the grid area to give the host keyboard focus, then jump to A1
    // so scenarios start from a deterministic cell. Without the Home step the
    // click position would land on whatever cell happens to be under the
    // pixel offset (varies with viewport width / column sizes / fixtures).
    await this.page
      .locator('.fc-host')
      .first()
      .click({ position: { x: 200, y: 200 } });
    const isMac = await this.page.evaluate(() => navigator.platform.toLowerCase().includes('mac'));
    await this.page.keyboard.press(`${isMac ? 'Meta' : 'Control'}+Home`);
  }

  /** Read the formula-bar textarea contents. */
  async formulaBarValue(): Promise<string> {
    return this.page.locator('.fc-host__formulabar-input').first().inputValue();
  }

  /** Cross-OS Mod+key. macOS uses Meta; everything else uses Control. */
  async shortcut(key: string): Promise<void> {
    const isMac = await this.page.evaluate(() => navigator.platform.toLowerCase().includes('mac'));
    await this.page.keyboard.press(`${isMac ? 'Meta' : 'Control'}+${key}`);
  }

  /** Fail the surrounding test if any console error fired during the run.
   *  Bind once near the start of a test, drain at the end. */
  collectConsoleErrors(): { read: () => string[] } {
    const errors: string[] = [];
    this.page.on('console', (msg) => {
      if (msg.type() === 'error') errors.push(msg.text());
    });
    this.page.on('pageerror', (err) => errors.push(err.message));
    return { read: () => [...errors] };
  }
}
