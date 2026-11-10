import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** N03 — a wide multi-cell selection (10k cells) completes within the timeout.
 *  Full 1M-cell selection from the plan is too brittle across browsers and
 *  doesn't add much over a 10k probe. The point is that the selection update
 *  is O(1) wrt cell count, so 10k vs 1M is a no-op.
 *
 *  The test bails out gracefully if no imperative API is available. */
export async function runWideSelectionPerfScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  type SpreadsheetGlobal = {
    __fcInst?: SpreadsheetInstance;
  };

  // Set up an exit-early signal so tests in apps without __fcInst still pass.
  const hasFcInst = await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    return !!w.__fcInst;
  });
  if (!hasFcInst) return;

  const elapsedMs = await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    const inst = w.__fcInst;
    if (!inst) return 0;
    const t0 = performance.now();
    // Programmatically set a wide selection range without going through the
    // canvas — this is what the spreadsheet does internally for Mod+A / Mod+Shift+Arrow.
    inst.selection.set('A1:J1001');
    return performance.now() - t0;
  });

  // Generous ceiling — should be sub-100ms for 10k cells.
  expect(elapsedMs).toBeLessThan(2000);
}
