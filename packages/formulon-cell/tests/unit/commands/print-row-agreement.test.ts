// @vitest-environment node

import { describe, expect, it } from 'vitest';
import { buildPrintDocument } from '../../../src/commands/print.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { syncPageSetupToEngine } from '../../../src/engine/print-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

type RawWorkbook = {
  getRowHeightPt(sheet: number, row: number): { status: { ok: boolean }; value: number };
};

describe('printed rows on a new default-profile workbook', () => {
  it('prints every row at the height the engine lays it out with', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'en' });
    const store = createSpreadsheetStore();
    try {
      hydrateLayoutFromEngine(wb, store, 0);
      for (let row = 0; row < 3; row += 1) {
        wb.setNumber({ sheet: 0, row, col: 0 }, row + 1);
        store.setState((s) => {
          const cells = new Map(s.data.cells);
          cells.set(addrKey({ sheet: 0, row, col: 0 }), {
            value: { kind: 'number', value: row + 1 },
            formula: null,
          });
          return { ...s, data: { ...s.data, cells } };
        });
      }
      // Title rows print at the native scale (100%), so their inline height
      // is the store's row height in points.
      mutators.setPageSetup(store, 0, { printArea: 'A1:A3', printTitleRows: '1:3' });
      syncPageSetupToEngine(wb, store, 0);

      const document = buildPrintDocument(wb, store, 0);
      const heights = [...document.html.matchAll(/<tr style="height:([0-9.]+)pt">/g)].map((match) =>
        Number(match[1]),
      );
      expect(heights.length).toBeGreaterThanOrEqual(3);
      const raw = (wb as unknown as { wb: RawWorkbook }).wb;
      for (let row = 0; row < 3; row += 1) {
        expect(heights[row]).toBeCloseTo(raw.getRowHeightPt(0, row).value, 8);
        expect(heights[row]).toBe(15);
      }
    } finally {
      wb.dispose();
    }
  });
});
