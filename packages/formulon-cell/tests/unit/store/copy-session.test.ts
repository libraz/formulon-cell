import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { insertCols, insertRows } from '../../../src/commands/structure.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('internal clipboard session revision', () => {
  let store: ReturnType<typeof createSpreadsheetStore>;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault({ preferStub: true });
  });

  afterEach(() => wb.dispose());

  it('advances for every public copy-range replacement and clear', () => {
    expect(store.getState().ui.copyRevision).toBe(0);

    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(store.getState().ui.copyRevision).toBe(1);

    mutators.setCopyRange(store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    expect(store.getState().ui.copyRevision).toBe(2);

    mutators.setCopyRange(store, null);
    expect(store.getState().ui.copyRevision).toBe(3);

    mutators.setCopyRanges(store, [{ sheet: 0, r0: 2, c0: 2, r1: 2, c1: 3 }]);
    expect(store.getState().ui.copyRevision).toBe(4);

    mutators.setCopyRanges(store, null);
    expect(store.getState().ui.copyRevision).toBe(5);
  });

  it('keeps the revision while structural edits move the live marquee', () => {
    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const revision = store.getState().ui.copyRevision;

    insertRows(store, wb, null, 0, 1);
    insertCols(store, wb, null, 0, 1);

    expect(store.getState().ui.copyRevision).toBe(revision);
    expect(store.getState().ui.copyRange).toEqual({
      sheet: 0,
      r0: 1,
      c0: 1,
      r1: 1,
      c1: 1,
    });
  });
});
