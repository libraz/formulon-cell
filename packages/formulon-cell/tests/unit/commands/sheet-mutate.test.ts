import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  addAllowedEditRange,
  setWorkbookStructureProtected,
} from '../../../src/commands/protection.js';
import {
  addSheet,
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
  add: number;
  rename: { idx: number; name: string }[];
  remove: number[];
  move: { from: number; to: number }[];
  /** Result toggles — flip to false to simulate engine rejection. */
  acceptRename: boolean;
  acceptRemove: boolean;
  acceptMove: boolean;
  names: string[];
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
    add: 0,
    rename: [],
    remove: [],
    move: [],
    acceptRename: true,
    acceptRemove: true,
    acceptMove: true,
    names: Array.from({ length: opts.sheetCount ?? 3 }, (_, i) => `Sheet${i + 1}`),
  };
  const wb = {
    capabilities: fake.capabilities,
    get sheetCount() {
      return fake.sheetCount;
    },
    renameSheet: (idx: number, name: string): boolean => {
      if (!fake.acceptRename) return false;
      fake.rename.push({ idx, name });
      fake.names[idx] = name;
      return true;
    },
    sheetName: (idx: number): string => fake.names[idx] ?? `Sheet${idx + 1}`,
    addSheet: (name?: string): number => {
      fake.add += 1;
      const idx = fake.sheetCount;
      fake.sheetCount += 1;
      fake.names.push(name ?? `Sheet${idx + 1}`);
      return idx;
    },
    removeSheet: (idx: number): boolean => {
      if (!fake.acceptRemove) return false;
      fake.remove.push(idx);
      fake.sheetCount -= 1;
      fake.names.splice(idx, 1);
      return true;
    },
    moveSheet: (from: number, to: number): boolean => {
      if (!fake.acceptMove) return false;
      fake.move.push({ from, to });
      const [name] = fake.names.splice(from, 1);
      if (name) fake.names.splice(to, 0, name);
      return true;
    },
  } as unknown as WorkbookHandle;
  return { wb, fake };
};

const setActive = (store: SpreadsheetStore, idx: number): void => {
  mutators.setSheetIndex(store, idx);
};

describe('addSheet', () => {
  it('adds a sheet unless workbook structure is protected', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeFake({ sheetCount: 2 });
    expect(addSheet(store, wb)).toBe(2);
    expect(fake.add).toBe(1);
    expect(fake.sheetCount).toBe(3);

    setWorkbookStructureProtected(store, true);
    expect(addSheet(store, wb)).toBe(-1);
    expect(fake.add).toBe(1);
    expect(fake.sheetCount).toBe(3);
  });

  it('records sheet insertion as an undoable workbook action', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const { wb, fake } = makeFake({ sheetCount: 2 });
    const added = addSheet(store, wb, history);
    expect(added).toBe(2);
    mutators.setSheetIndex(store, added);

    expect(history.undo()).toBe(true);
    expect(fake.sheetCount).toBe(2);
    expect(fake.remove).toEqual([2]);
    expect(store.getState().data.sheetIndex).toBe(0);

    expect(history.redo()).toBe(true);
    expect(fake.sheetCount).toBe(3);
    expect(store.getState().data.sheetIndex).toBe(2);
  });
});

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

  it('returns false without touching the engine when workbook structure is protected', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeFake();
    setWorkbookStructureProtected(store, true);
    expect(renameSheet(wb, 0, 'Blocked', store)).toBe(false);
    expect(fake.rename).toEqual([]);
  });

  it('records sheet rename as an undoable workbook action', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const { wb, fake } = makeFake();
    expect(renameSheet(wb, 1, 'Renamed', store, history)).toBe(true);
    expect(fake.names[1]).toBe('Renamed');
    expect(history.undo()).toBe(true);
    expect(fake.names[1]).toBe('Sheet2');
    expect(history.redo()).toBe(true);
    expect(fake.names[1]).toBe('Renamed');
  });
});

describe('removeSheet', () => {
  it('refuses to remove the last remaining sheet', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeFake({ sheetCount: 1 });
    expect(removeSheet(store, wb, 0)).toBe(false);
    expect(fake.remove).toEqual([]);
  });

  it('refuses to remove sheets when workbook structure is protected', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeFake();
    setWorkbookStructureProtected(store, true);
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

  it('remaps workbook sheet layout metadata after removal', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        hiddenSheets: new Set([2]),
        sheetTabColors: new Map([
          [0, '#c00000'],
          [2, '#70ad47'],
        ]),
      },
    }));
    const { wb } = makeFake({ sheetCount: 3 });
    expect(removeSheet(store, wb, 1)).toBe(true);
    expect(Array.from(store.getState().layout.hiddenSheets)).toEqual([1]);
    expect(Array.from(store.getState().layout.sheetTabColors.entries())).toEqual([
      [0, '#c00000'],
      [1, '#70ad47'],
    ]);
  });

  it('remaps workbook protection metadata after removal', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeFake({ sheetCount: 3 });
    mutators.setSheetProtected(store, 1, true, { password: 'remove' });
    mutators.setSheetProtected(store, 2, true, { password: 'keep' });
    addAllowedEditRange(store, { sheet: 1, r0: 0, c0: 0, r1: 0, c1: 0 }, { title: 'Drop' });
    addAllowedEditRange(store, { sheet: 2, r0: 1, c0: 1, r1: 1, c1: 1 }, { title: 'Keep' });

    expect(removeSheet(store, wb, 1)).toBe(true);

    const protection = store.getState().protection;
    expect(Array.from(protection.protectedSheets.entries())).toEqual([[1, { password: 'keep' }]]);
    expect(protection.allowedEditRanges).toHaveLength(1);
    expect(protection.allowedEditRanges[0]?.title).toBe('Keep');
    expect(protection.allowedEditRanges[0]?.range.sheet).toBe(1);
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

  it('refuses to move sheets when workbook structure is protected', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeFake();
    setWorkbookStructureProtected(store, true);
    expect(moveSheet(store, wb, 0, 2)).toBe(false);
    expect(fake.move).toEqual([]);
  });

  it('moves workbook sheet layout metadata with the sheet', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        hiddenSheets: new Set([0]),
        sheetTabColors: new Map([
          [0, '#c00000'],
          [2, '#4472c4'],
        ]),
      },
    }));
    const { wb } = makeFake();
    expect(moveSheet(store, wb, 0, 2)).toBe(true);
    expect(Array.from(store.getState().layout.hiddenSheets)).toEqual([2]);
    expect(Array.from(store.getState().layout.sheetTabColors.entries())).toEqual([
      [2, '#c00000'],
      [1, '#4472c4'],
    ]);
  });

  it('moves workbook protection metadata with the sheet', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeFake();
    mutators.setSheetProtected(store, 0, true, { password: 'move' });
    mutators.setSheetProtected(store, 2, true, { password: 'shift' });
    addAllowedEditRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, { title: 'Move' });
    addAllowedEditRange(store, { sheet: 2, r0: 1, c0: 1, r1: 1, c1: 1 }, { title: 'Shift' });

    expect(moveSheet(store, wb, 0, 2)).toBe(true);

    const protection = store.getState().protection;
    expect(Array.from(protection.protectedSheets.entries())).toEqual([
      [2, { password: 'move' }],
      [1, { password: 'shift' }],
    ]);
    expect(protection.allowedEditRanges.map((entry) => [entry.title, entry.range.sheet])).toEqual([
      ['Move', 2],
      ['Shift', 1],
    ]);
  });

  it('records sheet move as an undoable workbook action', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    setActive(store, 0);
    const { wb, fake } = makeFake();
    expect(moveSheet(store, wb, 0, 2, history)).toBe(true);
    expect(fake.names).toEqual(['Sheet2', 'Sheet3', 'Sheet1']);
    expect(store.getState().data.sheetIndex).toBe(2);

    expect(history.undo()).toBe(true);
    expect(fake.names).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(store.getState().data.sheetIndex).toBe(0);

    expect(history.redo()).toBe(true);
    expect(fake.names).toEqual(['Sheet2', 'Sheet3', 'Sheet1']);
    expect(store.getState().data.sheetIndex).toBe(2);
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

  it('refuses to hide or unhide sheets when workbook structure is protected', () => {
    const store = createSpreadsheetStore();
    const { wb, fake } = makeHiddenFake({ sheetCount: 3 });
    setWorkbookStructureProtected(store, true);
    expect(setSheetHidden(store, wb, null, 1, true)).toBe(false);
    expect(setSheetHidden(store, wb, null, 1, false)).toBe(false);
    expect(fake.hide).toEqual([]);
  });
});
