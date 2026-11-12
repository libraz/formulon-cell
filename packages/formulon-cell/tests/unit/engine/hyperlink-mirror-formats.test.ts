import { describe, expect, it } from 'vitest';
import { flushFormatToEngine } from '../../../src/engine/cell-format-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

describe.skipIf(!canLoadWasm())('hyperlinks mirrored to the engine', () => {
  // The store owns cell formats: writing the link set must neither restyle the linked cells
  // nor reset them when the set is written again.
  it('leaves the format of a linked cell as the store has it, flush after flush', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const key = addrKey({ sheet: 0, row: 0, col: 0 });
      store.setState((s) => ({
        ...s,
        format: {
          ...s.format,
          formats: new Map([[key, { bold: true, hyperlink: 'https://a.example/' }]]),
        },
      }));
      const bold = (): boolean | null => {
        const xf = wb.getCellXfIndex(0, 0, 0);
        const font = xf === null ? null : wb.getCellXf(xf)?.fontIndex;
        return font === null || font === undefined ? null : (wb.getFontRecord(font)?.bold ?? null);
      };
      flushFormatToEngine(wb, store, 0);
      expect(bold()).toBe(true);
      expect(wb.getHyperlinksFull(0)?.map((link) => link.target)).toEqual(['https://a.example/']);

      store.setState((s) => ({
        ...s,
        format: {
          ...s.format,
          formats: new Map([[key, { bold: true, hyperlink: 'https://b.example/' }]]),
        },
      }));
      flushFormatToEngine(wb, store, 0);
      expect(bold()).toBe(true);
      expect(wb.getHyperlinksFull(0)?.map((link) => link.target)).toEqual(['https://b.example/']);
    } finally {
      wb.dispose();
    }
  });
});
