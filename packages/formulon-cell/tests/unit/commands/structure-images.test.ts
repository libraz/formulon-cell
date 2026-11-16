// @vitest-environment node

import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { deleteCols, insertCols, insertRows } from '../../../src/commands/structure.js';
import {
  hydrateImagesFromEngine,
  persistentImageFromState,
} from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { reloadDrawingsAfterSheetEdit } from '../../../src/mount/engine-binding.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// 1x1 transparent PNG.
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  ),
  (char) => char.charCodeAt(0),
);

const anchorOf = (store: SpreadsheetStore) => {
  const image = store
    .getState()
    .illustrations.illustrations.map(persistentImageFromState)
    .find((item) => item !== null);
  return image?.nativeAnchor;
};

describe('pictures follow row/column edits', () => {
  it('keeps the stored anchor in step with the engine through edit, undo and redo', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(
        wb.insertImage(0, PNG, { anchor: { fromRow: 4, fromCol: 2 }, name: 'logo' }),
      ).not.toBeNull();
      hydrateImagesFromEngine(wb, store, 0);
      expect(anchorOf(store)).toMatchObject({ fromRow: 4, fromCol: 2 });

      const history = new History();
      wb.attachHistory(history);
      insertRows(store, wb, history, 1, 2);
      expect(anchorOf(store)).toMatchObject({ fromRow: 6, fromCol: 2 });

      insertCols(store, wb, history, 0, 1);
      expect(anchorOf(store)).toMatchObject({ fromRow: 6, fromCol: 3 });

      history.undo();
      expect(anchorOf(store)).toMatchObject({ fromRow: 6, fromCol: 2 });
      history.undo();
      expect(anchorOf(store)).toMatchObject({ fromRow: 4, fromCol: 2 });
      history.redo();
      expect(anchorOf(store)).toMatchObject({ fromRow: 6, fromCol: 2 });

      deleteCols(store, wb, history, 0, 1);
      expect(anchorOf(store)).toMatchObject({ fromRow: 6, fromCol: 1 });
      history.undo();
      expect(anchorOf(store)).toMatchObject({ fromRow: 6, fromCol: 2 });
    } finally {
      wb.dispose();
    }
  });

  it('re-projects pictures with their new sheet index after a sheet move', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.addSheet('Second');
      expect(
        wb.insertImage(1, PNG, { anchor: { fromRow: 2, fromCol: 1 }, name: 'logo' }),
      ).not.toBeNull();
      mutators.setSheetIndex(store, 1);
      hydrateImagesFromEngine(wb, store, 1);
      expect(store.getState().illustrations.illustrations[0]?.sheet).toBe(1);

      wb.attachStore(store);
      // The mounted sheet tabs reload drawings through this entry point.
      const unsubscribe = wb.subscribe((event) => {
        if (event.kind === 'sheet-move') reloadDrawingsAfterSheetEdit(wb, store, event);
      });
      expect(wb.moveSheet(1, 0)).toBe(true);
      unsubscribe();
      mutators.setSheetIndex(store, 0);
      const items = store.getState().illustrations.illustrations;
      expect(items).toHaveLength(1);
      expect(items[0]?.sheet).toBe(0);
      expect(anchorOf(store)).toMatchObject({ fromRow: 2, fromCol: 1 });
    } finally {
      wb.dispose();
    }
  });
});
