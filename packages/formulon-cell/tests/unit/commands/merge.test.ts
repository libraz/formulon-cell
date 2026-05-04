import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  applyMerge,
  applyUnmerge,
  expandRangeWithMerges,
  mergeAnchorOf,
  mergeAt,
  stepWithMerge,
} from '../../../src/commands/merge.js';
import type { Range } from '../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const key = `${0}:${c.row}:${c.col}`;
      if (typeof c.value === 'number') {
        wb.setNumber({ sheet: 0, row: c.row, col: c.col }, c.value);
        map.set(key, { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText({ sheet: 0, row: c.row, col: c.col }, c.value);
        map.set(key, { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

describe('applyMerge', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    history = new History();
    wb.attachHistory(history);
  });

  it('returns false on a 1×1 range and writes nothing', () => {
    const r: Range = { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 };
    expect(applyMerge(store, wb, history, r)).toBe(false);
    expect(store.getState().merges.byAnchor.size).toBe(0);
  });

  it('records the merge in store.merges with anchor + reverse index', () => {
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    applyMerge(store, wb, history, r);
    const m = store.getState().merges;
    expect(m.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual(r);
    // 3 non-anchor cells map back to the anchor.
    expect(m.byCell.size).toBe(3);
    expect(m.byCell.get(addrKey({ sheet: 0, row: 0, col: 1 }))).toBe(
      addrKey({ sheet: 0, row: 0, col: 0 }),
    );
  });

  it('clears non-anchor cell values (Excel keeps only top-left)', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'keep' },
      { row: 0, col: 1, value: 'drop1' },
      { row: 1, col: 0, value: 999 },
    ]);
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    applyMerge(store, wb, history, r);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'keep' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 }).kind).toBe('blank');
  });

  it('skips writes on already-blank non-anchor cells (no superfluous history entries)', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'only' }]);
    const before = history.canUndo();
    expect(before).toBe(true); // seed writes pushed entries
    // Pop them so we have a clean slate.
    while (history.canUndo()) history.undo();
    while (history.canRedo()) history.redo();
    while (history.canUndo()) history.undo();
    history.clear();
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    applyMerge(store, wb, history, r);
    // Only the merges-state entry should be on the stack, not a setBlank for the
    // empty (0,1) cell.
    let count = 0;
    while (history.canUndo()) {
      history.undo();
      count += 1;
    }
    expect(count).toBe(1);
  });

  it('strips an existing merge that overlaps the new one', () => {
    applyMerge(store, wb, history, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    applyMerge(store, wb, history, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    const m = store.getState().merges;
    // First merge's anchor was deleted because the second range overlaps it.
    expect(m.byAnchor.has(addrKey({ sheet: 0, row: 0, col: 0 }))).toBe(false);
    expect(m.byAnchor.has(addrKey({ sheet: 0, row: 1, col: 1 }))).toBe(true);
  });

  it('is undoable as a single step (cells + merge state both reverted)', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'A' },
      { row: 0, col: 1, value: 'B' },
    ]);
    applyMerge(store, wb, history, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    expect(store.getState().merges.byAnchor.size).toBe(1);
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');

    // One transaction was committed for the merge — undo it.
    history.undo();

    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'B' });
  });

  it('redo re-applies both the cell clearing and the merge state', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 1, value: 'drop' }]);
    applyMerge(store, wb, history, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    history.undo();
    history.redo();
    expect(store.getState().merges.byAnchor.size).toBe(1);
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
  });
});

describe('applyUnmerge', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    history = new History();
    wb.attachHistory(history);
  });

  it('returns false when no merge is touched and pushes nothing', () => {
    const before = history.canUndo();
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    expect(applyUnmerge(store, wb, history, r)).toBe(false);
    expect(history.canUndo()).toBe(before);
  });

  it('removes the merge and is undoable', () => {
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    applyMerge(store, wb, history, r);
    expect(applyUnmerge(store, wb, history, r)).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    history.undo();
    expect(store.getState().merges.byAnchor.size).toBe(1);
  });
});

describe('mergeAt / mergeAnchorOf', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('mergeAt returns null when no merge covers the address', () => {
    expect(mergeAt(store.getState(), { sheet: 0, row: 5, col: 5 })).toBeNull();
  });

  it('mergeAt finds the merge for an anchor cell', () => {
    const r: Range = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 };
    applyMerge(store, wb, null, r);
    expect(mergeAt(store.getState(), { sheet: 0, row: 1, col: 1 })).toEqual(r);
  });

  it('mergeAt finds the merge for a body (non-anchor) cell', () => {
    const r: Range = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 };
    applyMerge(store, wb, null, r);
    expect(mergeAt(store.getState(), { sheet: 0, row: 2, col: 3 })).toEqual(r);
  });

  it('mergeAnchorOf returns the anchor for a body cell', () => {
    applyMerge(store, wb, null, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 });
    expect(mergeAnchorOf(store.getState(), { sheet: 0, row: 2, col: 2 })).toEqual({
      sheet: 0,
      row: 1,
      col: 1,
    });
  });

  it('mergeAnchorOf passes through addresses outside any merge', () => {
    expect(mergeAnchorOf(store.getState(), { sheet: 0, row: 5, col: 5 })).toEqual({
      sheet: 0,
      row: 5,
      col: 5,
    });
  });
});

describe('expandRangeWithMerges', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('returns the range unchanged when no merges intersect', () => {
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    expect(expandRangeWithMerges(store.getState(), r)).toEqual(r);
  });

  it('grows to cover a merge that the range partially overlaps', () => {
    applyMerge(store, wb, null, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 });
    // Selection only touches the top-left corner of the merge.
    const expanded = expandRangeWithMerges(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    expect(expanded).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 });
  });

  it('grows to cover multiple non-overlapping merges that the range spans', () => {
    applyMerge(store, wb, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    applyMerge(store, wb, null, { sheet: 0, r0: 2, c0: 3, r1: 3, c1: 4 });
    // Range that touches both merges (corners only) should expand to fully cover.
    const expanded = expandRangeWithMerges(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 1,
      r1: 2,
      c1: 3,
    });
    expect(expanded).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 4 });
  });
});

describe('stepWithMerge', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('snaps to a merge anchor when stepping into the body', () => {
    applyMerge(store, wb, null, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 });
    // Stepping right from (1,0) lands inside the merge — snap to anchor (1,1).
    expect(stepWithMerge(store.getState(), { sheet: 0, row: 1, col: 0 }, 0, 1, 1000, 1000)).toEqual(
      { sheet: 0, row: 1, col: 1 },
    );
  });

  it('exits a merge from its right edge when stepping right', () => {
    applyMerge(store, wb, null, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 });
    // From inside the merge body, ArrowRight should land just outside (col 4).
    expect(stepWithMerge(store.getState(), { sheet: 0, row: 1, col: 1 }, 0, 1, 1000, 1000)).toEqual(
      { sheet: 0, row: 1, col: 4 },
    );
  });

  it('exits a merge from its bottom edge when stepping down', () => {
    applyMerge(store, wb, null, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 });
    expect(stepWithMerge(store.getState(), { sheet: 0, row: 1, col: 2 }, 1, 0, 1000, 1000)).toEqual(
      { sheet: 0, row: 3, col: 2 },
    );
  });

  it('exits a merge from its top edge when stepping up', () => {
    applyMerge(store, wb, null, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 });
    expect(
      stepWithMerge(store.getState(), { sheet: 0, row: 2, col: 2 }, -1, 0, 1000, 1000),
    ).toEqual({ sheet: 0, row: 0, col: 2 });
  });

  it('clamps at sheet edges', () => {
    expect(stepWithMerge(store.getState(), { sheet: 0, row: 0, col: 0 }, -1, 0, 100, 100)).toEqual({
      sheet: 0,
      row: 0,
      col: 0,
    });
  });
});
