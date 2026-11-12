import { describe, expect, it } from 'vitest';
import { hydrateCommentsAndHyperlinksFromEngine } from '../../../src/engine/format-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

// A workbook write outside the grid (the sheet facade) is followed by a re-read of the sheet:
// what the engine no longer holds must leave the grid as well.
describe.skipIf(!canLoadWasm())('re-reading a sheet after the engine dropped a note', () => {
  it('drops the note from the grid, and a hidden note shows hidden', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const a1 = addrKey({ sheet: 0, row: 0, col: 0 });
      const b2 = addrKey({ sheet: 0, row: 1, col: 1 });
      const format = (key: string) => store.getState().format.formats.get(key);
      wb.setCommentEntry(0, 0, 0, 'Ann', 'first', true);
      wb.setCommentEntry(0, 1, 1, 'Ann', 'second', true);
      hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
      expect(format(a1)).toMatchObject({ comment: 'first', commentVisible: true });
      expect(format(b2)).toMatchObject({ comment: 'second', commentVisible: true });

      wb.setCommentEntry(0, 0, 0, null, '');
      wb.setCommentEntry(0, 1, 1, 'Ann', 'second', false);
      hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
      expect(format(a1)?.comment).toBeUndefined();
      expect(format(a1)?.commentAuthor).toBeUndefined();
      expect(format(b2)).toMatchObject({ comment: 'second' });
      expect(format(b2)?.commentVisible).toBeUndefined();
    } finally {
      wb.dispose();
    }
  });
});
