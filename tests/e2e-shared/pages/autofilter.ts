import type { Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';

type ProbeWindow = Window & { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };

/** Page point of the AutoFilter button in the header cell at (`row`, `col`), left-to-right sheet. */
export const filterButtonPoint = async (page: Page, row: number, col: number) => {
  const box = await page.locator('.fc-host__grid').first().boundingBox();
  if (!box) throw new Error('Missing grid');
  const point = await page.evaluate(
    ({ row, col }) => {
      const w = window as unknown as ProbeWindow;
      const state = w.__fcProbes['mount/internals']
        .spreadsheetInternals(w.__fcInst)
        .store.getState();
      const { cellRectUnclamped, layoutForView } = w.__fcProbes['render/geometry'];
      const rect = cellRectUnclamped(layoutForView(state), state.viewport, row, col);
      // The 17 px button sits one px inside the cell's trailing edge and one below its top.
      return { x: rect.x + rect.w - 1 - 17 / 2, y: rect.y + 1 + 17 / 2 };
    },
    { row, col },
  );
  return { x: box.x + point.x, y: box.y + point.y };
};

/** The rows the view hides, ascending. */
export const hiddenRows = (page: Page) =>
  page.evaluate(() => {
    const w = window as unknown as ProbeWindow;
    const state = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst).store.getState();
    return [...state.layout.hiddenRows].sort((a, b) => a - b);
  });

/** The sheet 0 AutoFilter range, or null when the sheet has none. */
export const readAutoFilterRange = (page: Page) =>
  page.evaluate(
    () =>
      (window as unknown as ProbeWindow).__fcInst.workbook.sheet(0).getAutoFilter()?.range ?? null,
  );
