import { describe, expect, it, vi } from 'vitest';
import { hydrateCommentsAndHyperlinksFromEngine } from '../../../src/engine/format-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

// A sheet whose link list the engine refuses to read says nothing about which links are stale.
describe.skipIf(!canLoadWasm())('re-reading a sheet when the engine cannot list its links', () => {
  it('keeps the links the grid holds', async () => {
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
      vi.spyOn(wb, 'getHyperlinksFull').mockReturnValue(null);
      hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
      expect(store.getState().format.formats.get(key)).toEqual({
        bold: true,
        hyperlink: 'https://a.example/',
      });
    } finally {
      wb.dispose();
    }
  });
});
