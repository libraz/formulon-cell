import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** D01 — `workbook.save()` round-trips the current workbook to xlsx bytes.
 *  We don't trigger a download (UA-specific); instead we exercise the same
 *  call path the demo's Save button uses and assert it returns a non-trivial
 *  Uint8Array. Playground-only because we need `window.__fcInst`. */
export async function runSaveXlsxScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  type SpreadsheetGlobal = {
    __fcInst?: {
      workbook: { save(): Uint8Array };
    };
  };

  const out = await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    const inst = w.__fcInst;
    if (!inst) return { ok: false as const };
    try {
      const bytes = inst.workbook.save();
      return { ok: true as const, length: bytes.byteLength, head: Array.from(bytes.slice(0, 4)) };
    } catch (e) {
      return { ok: false as const, error: (e as Error).message };
    }
  });

  expect(out.ok).toBe(true);
  if (!out.ok) return;
  // xlsx is a ZIP — the first two bytes are "PK" (0x50, 0x4b).
  expect(out.length).toBeGreaterThan(100);
  expect(out.head[0]).toBe(0x50);
  expect(out.head[1]).toBe(0x4b);
}
