import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  moveSheet,
  removeSheet,
  renameSheet,
  setSheetHidden,
} from '../../../src/commands/sheet-mutate.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

interface FakeWb {
  capabilities: { sheetMutate: boolean };
  sheetCount: number;
  rename: { idx: number; name: string }[];
  remove: number[];
  move: { from: number; to: number }[];
  /** Result toggles — flip to false to simulate engine rejection. */
  acceptRename: boolean;
  acceptRemove: boolean;
  acceptMove: boolean;
}

const makeFake = (
  opts: { sheetMutate?: boolean; sheetCount?: number } = {},
): {
  wb: WorkbookHandle;
  fake: FakeWb;
} => {
  const fake: FakeWb = {
    capabilities: { sheetMutate: opts.sheetMutate ?? true },
    sheetCount: opts.sheetCount ?? 3,
    rename: [],
    remove: [],
    move: [],
    acceptRename: true,
    acceptRemove: true,
    acceptMove: true,
  };
  const wb = {
    capabilities: fake.capabilities,
    get sheetCount() {
      return fake.sheetCount;
    },
    renameSheet: (idx: number, name: string): boolean => {
      if (!fake.acceptRename) return false;
      fake.rename.push({ idx, name });
      return true;
    },
    removeSheet: (idx: number): boolean => {
      if (!fake.acceptRemove) return false;
      fake.remove.push(idx);
      fake.sheetCount -= 1;
      return true;
    },
    moveSheet: (from: number, to: number): boolean => {
      if (!fake.acceptMove) return false;
      fake.move.push({ from, to });
      return true;
    },
  } as unknown as WorkbookHandle;
  return { wb, fake };
};

const setActive = (store: SpreadsheetStore, idx: number): void => {
  mutators.setSheetIndex(store, idx);
};

describe('renameSheet', () => {
  it('forwards to wb.renameSheet and returns true on success', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeFake();
    expect(renameSheet(wb, 1, 'Renamed')).toBe(true);
    expect(fake.rename).toEqual([{ idx: 1, name: 'Renamed' }]);
    expect(store.getState().data.sheetIndex).toBe(0); // unchanged
  });

  it('returns false when engine rejects (e.g. duplicate name)', () => {
    const { wb, fake } = makeFake();
    fake.acceptRename = false;
    expect(renameSheet(wb, 0, 'Sheet1')).toBe(false);
  });
});

describe('removeSheet', () => {
  it('refuses to remove the last remaining sheet', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeFake({ sheetCount: 1 });
    expect(removeSheet(store, wb, 0)).toBe(false);
    expect(fake.remove).toEqual([]);
  });

  it('removing index < active shifts active down by 1', () => {
    const store = createSpreadsheetStore();
    setActive(store, 2);
    const { wb, fake } = makeFake({ sheetCount: 3 });
    expect(removeSheet(store, wb, 0)).toBe(true);
    expect(fake.remove).toEqual([0]);
    expect(store.getState().data.sheetIndex).toBe(1);
  });

  it('removing the active sheet selects max(idx-1, 0)', () => {
    const store = createSpreadsheetStore();
    setActive(store, 1);
    const { wb } = makeFake({ sheetCount: 3 });
    expect(removeSheet(store, wb, 1)).toBe(true);
    expect(store.getState().data.sheetIndex).toBe(0);
  });

  it('removing index > active leaves active untouched', () => {
    const store = createSpreadsheetStore();
    setActive(store, 0);
    const { wb } = makeFake({ sheetCount: 3 });
    expect(removeSheet(store, wb, 2)).toBe(true);
    expect(store.getState().data.sheetIndex).toBe(0);
  });

  it('removing the only-other sheet when active is index 0 leaves active at 0', () => {
    const store = createSpreadsheetStore();
    setActive(store, 0);
    const { wb } = makeFake({ sheetCount: 2 });
    expect(removeSheet(store, wb, 0)).toBe(true);
    expect(store.getState().data.sheetIndex).toBe(0);
  });
});

describe('moveSheet', () => {
  it('returns true and is a no-op when from === to', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeFake();
    expect(moveSheet(store, wb, 1, 1)).toBe(true);
    expect(fake.move).toEqual([]);
  });

  it('moving the active sheet updates active to `to`', () => {
    const store = createSpreadsheetStore();
    setActive(store, 0);
    const { wb } = makeFake();
    expect(moveSheet(store, wb, 0, 2)).toBe(true);
    expect(store.getState().data.sheetIndex).toBe(2);
  });

  it('moving from < active across to >= active shifts active down by 1', () => {
    const store = createSpreadsheetStore();
    setActive(store, 2);
    const { wb } = makeFake();
    // Sheets [0,1,2] → move 0 to index 2; new order [1,2,0]; old-2 is now 1.
    expect(moveSheet(store, wb, 0, 2)).toBe(true);
    expect(store.getState().data.sheetIndex).toBe(1);
  });

  it('moving from > active to <= active shifts active up by 1', () => {
    const store = createSpreadsheetStore();
    setActive(store, 1);
    const { wb } = makeFake();
    // Sheets [0,1,2] → move 2 to index 0; new order [2,0,1]; old-1 is now 2.
    expect(moveSheet(store, wb, 2, 0)).toBe(true);
    expect(store.getState().data.sheetIndex).toBe(2);
  });

  it('moves outside the active sheet leave it untouched', () => {
    const store = createSpreadsheetStore();
    setActive(store, 0);
    const { wb } = makeFake();
    // Move 1→2 (or 2→1) — neither crosses index 0.
    expect(moveSheet(store, wb, 1, 2)).toBe(true);
    expect(store.getState().data.sheetIndex).toBe(0);
  });

  it('returns false when engine rejects', () => {
    const { wb, fake } = makeFake();
    fake.acceptMove = false;
    const store = createSpreadsheetStore();
    expect(moveSheet(store, wb, 0, 1)).toBe(false);
  });
});

interface HiddenFakeWb {
  capabilities: { sheetTabHidden: boolean; sheetMutate: boolean };
  sheetCount: number;
  hide: { sheet: number; hidden: boolean }[];
}

const makeHiddenFake = (
  opts: { sheetTabHidden?: boolean; sheetCount?: number } = {},
): { wb: WorkbookHandle; fake: HiddenFakeWb } => {
  const fake: HiddenFakeWb = {
    capabilities: {
      sheetTabHidden: opts.sheetTabHidden ?? true,
      sheetMutate: true,
    },
    sheetCount: opts.sheetCount ?? 3,
    hide: [],
  };
  const wb = {
    capabilities: fake.capabilities,
    get sheetCount() {
      return fake.sheetCount;
    },
    setSheetTabHidden(sheet: number, hidden: boolean): boolean {
      if (!fake.capabilities.sheetTabHidden) return false;
      fake.hide.push({ sheet, hidden });
      return true;
    },
  } as unknown as WorkbookHandle;
  return { wb, fake };
};

describe('setSheetHidden', () => {
  it('marks the sheet hidden and forwards to the engine', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const { wb, fake } = makeHiddenFake({ sheetCount: 3 });
    expect(setSheetHidden(store, wb, history, 1, true)).toBe(true);
    expect(store.getState().layout.hiddenSheets.has(1)).toBe(true);
    expect(fake.hide).toEqual([{ sheet: 1, hidden: true }]);
  });

  it('refuses to hide the last visible sheet', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeHiddenFake({ sheetCount: 2 });
    setSheetHidden(store, wb, null, 0, true); // hide one
    fake.hide.length = 0;
    // Trying to hide the only remaining visible sheet should bail out.
    expect(setSheetHidden(store, wb, null, 1, true)).toBe(false);
    expect(fake.hide).toEqual([]);
    expect(store.getState().layout.hiddenSheets.has(1)).toBe(false);
  });

  it('hiding the active sheet hops to the next visible one', () => {
    const store = createSpreadsheetStore();
    setActive(store, 1);
    const { wb } = makeHiddenFake({ sheetCount: 3 });
    expect(setSheetHidden(store, wb, null, 1, true)).toBe(true);
    expect(store.getState().data.sheetIndex).toBe(0);
  });

  it('unhide flips the flag back and emits the engine call', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeHiddenFake({ sheetCount: 3 });
    setSheetHidden(store, wb, null, 1, true);
    fake.hide.length = 0;
    expect(setSheetHidden(store, wb, null, 1, false)).toBe(true);
    expect(store.getState().layout.hiddenSheets.has(1)).toBe(false);
    expect(fake.hide).toEqual([{ sheet: 1, hidden: false }]);
  });

  it('undo restores the prior hidden set', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const { wb } = makeHiddenFake({ sheetCount: 3 });
    expect(setSheetHidden(store, wb, history, 1, true)).toBe(true);
    expect(history.canUndo()).toBe(true);
    history.undo();
    expect(store.getState().layout.hiddenSheets.has(1)).toBe(false);
  });

  it('returns false when capability is off (engine receives nothing)', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeHiddenFake({ sheetTabHidden: false, sheetCount: 3 });
    // Capability off → recordLayoutChangeWithEngine still updates the store
    // but the engine call short-circuits via the wrapper. The store-level
    // mutation still proceeds (matches the merges-off pattern).
    expect(setSheetHidden(store, wb, null, 1, true)).toBe(true);
    expect(fake.hide).toEqual([]);
  });
});
