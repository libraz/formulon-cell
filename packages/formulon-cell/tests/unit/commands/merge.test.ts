import { beforeEach, describe, expect, it } from 'vitest';
import { defaultTableOverlay } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import {
  applyMerge,
  applyMergeAcross,
  applyUnmerge,
  expandRangeWithMerges,
  mergeAcrossWillLoseData,
  mergeAnchorOf,
  mergeAt,
  mergeWillLoseData,
  stepWithMerge,
} from '../../../src/commands/merge.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import type { Range } from '../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

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

const seedBoolAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: boolean,
): void => {
  const addr = { sheet: 0, row, col };
  wb.setBool(addr, value);
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey(addr), { value: { kind: 'bool', value }, formula: null });
    return { ...s, data: { ...s.data, cells } };
  });
  wb.recalc();
};

const seedFormulaAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  formula: string,
): void => {
  const addr = { sheet: 0, row, col };
  wb.setFormula(addr, formula);
  wb.recalc();
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey(addr), { value: wb.getValue(addr), formula });
    return { ...s, data: { ...s.data, cells } };
  });
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

  it('refuses huge merge ranges instead of materializing every covered cell', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'keep' },
      { row: 200_000, col: 0, value: 'drop' },
    ]);
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 200_000, c1: 0 };

    expect(applyMerge(store, wb, history, range)).toBe(false);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(wb.getValue({ sheet: 0, row: 200_000, col: 0 })).toEqual({
      kind: 'text',
      value: 'drop',
    });
  });

  it('rejects invalid coordinates before expanding existing merges', () => {
    const invalid: Range[] = [
      { sheet: 0, r0: -1, c0: 0, r1: 1, c1: 1 },
      { sheet: 0, r0: 0.5, c0: 0, r1: 1, c1: 1 },
      { sheet: 0, r0: 0, c0: 0, r1: Number.POSITIVE_INFINITY, c1: 1 },
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 16_384 },
    ];
    for (const range of invalid) {
      expect(applyMerge(store, wb, history, range)).toBe(false);
      expect(applyMergeAcross(store, wb, history, range)).toBe(false);
    }
    expect(store.getState().merges.byAnchor.size).toBe(0);
  });

  it('clears non-anchor cell values (spreadsheets keep only top-left)', () => {
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

  it('promotes the sole non-anchor typed value to the top-left cell', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 1, value: 42 }]);
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    expect(applyMerge(store, wb, history, r)).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 42 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
  });

  it('promotes a sole non-anchor boolean and keeps false as content', () => {
    seedBoolAndMirror(store, wb, 0, 1, false);
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    expect(applyMerge(store, wb, history, r)).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'bool', value: false });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
  });

  it('promotes a sole non-anchor formula without rewriting its text', () => {
    seedFormulaAndMirror(store, wb, 1, 9, '=I2+10');
    const r: Range = { sheet: 0, r0: 0, c0: 7, r1: 1, c1: 9 };
    expect(applyMerge(store, wb, history, r)).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 7 })).toBe('=I2+10');
    expect(wb.getValue({ sheet: 0, row: 1, col: 9 })).toEqual({ kind: 'blank' });
  });

  it('copies anchor visual formatting to the merged cells and keeps it after unmerge', () => {
    const anchor: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        align: 'right',
        fontSize: 20,
        numFmt: { kind: 'fixed', decimals: 2 },
        color: 'red',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        align: 'left',
        fontSize: 9,
        numFmt: { kind: 'percent', decimals: 0 },
        color: 'green',
      },
    );

    expect(applyMerge(store, wb, history, anchor)).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 1 })),
    ).toMatchObject({
      align: 'right',
      fontSize: 20,
      numFmt: { kind: 'fixed', decimals: 2 },
      color: 'red',
    });

    expect(applyUnmerge(store, wb, history, anchor)).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 1 })),
    ).toMatchObject({
      align: 'right',
      fontSize: 20,
      numFmt: { kind: 'fixed', decimals: 2 },
      color: 'red',
    });
  });

  it('keeps a uniform perimeter border and removes internal borders', () => {
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    const edge = { style: 'thin' as const, color: 'red' };
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        borders: { top: edge, left: edge },
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        borders: { top: edge, right: edge },
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 0 },
      {
        borders: { bottom: edge, left: edge },
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      {
        borders: { bottom: edge, right: edge },
      },
    );

    expect(applyMerge(store, wb, history, range)).toBe(true);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      borders: { top: edge, left: edge },
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 1 }))).toEqual({
      borders: { top: edge, right: edge },
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))).toEqual({
      borders: { bottom: edge, left: edge },
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))).toEqual({
      borders: { bottom: edge, right: edge },
    });
  });

  it('removes partial perimeter and internal borders, then restores them on undo', () => {
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    const thick = { style: 'thick' as const };
    const thin = { style: 'thin' as const };
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        borders: { top: thin, right: thick },
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      {
        borders: { right: thin },
      },
    );
    const before = new Map(store.getState().format.formats);

    expect(applyMerge(store, wb, history, range)).toBe(true);
    for (const format of store.getState().format.formats.values()) {
      expect(format.borders?.top).toBeUndefined();
      expect(format.borders?.right).toBeUndefined();
    }
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats).toEqual(before);
  });

  it('removes partial diagonals but keeps uniform diagonals through unmerge', () => {
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    const diagonalDown = { style: 'medium' as const };
    const diagonalUp = { style: 'dotted' as const };
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        borders: { diagonalDown },
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      {
        borders: { diagonalUp },
      },
    );
    const partialBefore = new Map(store.getState().format.formats);

    expect(applyMerge(store, wb, history, range)).toBe(true);
    for (const format of store.getState().format.formats.values()) {
      expect(format.borders?.diagonalDown).toBeUndefined();
      expect(format.borders?.diagonalUp).toBeUndefined();
    }
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats).toEqual(partialBefore);

    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        mutators.setCellFormat(
          store,
          { sheet: 0, row, col },
          {
            borders: { diagonalDown, diagonalUp },
          },
        );
      }
    }
    expect(applyMerge(store, wb, history, range)).toBe(true);
    for (const format of store.getState().format.formats.values()) {
      expect(format.borders?.diagonalDown).toEqual(diagonalDown);
      expect(format.borders?.diagonalUp).toEqual(diagonalUp);
    }
    expect(applyUnmerge(store, wb, history, range)).toBe(true);
    for (const format of store.getState().format.formats.values()) {
      expect(format.borders?.diagonalDown).toEqual(diagonalDown);
      expect(format.borders?.diagonalUp).toEqual(diagonalUp);
    }
  });

  it('refuses a merge that intersects a structured table before writing cells', () => {
    const tableRange: Range = { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 2 };
    mutators.upsertTableOverlay(store, defaultTableOverlay('table', tableRange));
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'keep' }]);

    expect(applyMerge(store, wb, history, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 })).toBe(false);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'keep' });
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
    // Excel expands the second operation through the first merge, leaving one
    // chained merge from A1 through C3.
    expect(m.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 2,
      c1: 2,
    });
    expect(m.byAnchor.has(addrKey({ sheet: 0, row: 1, col: 1 }))).toBe(false);
  });

  it('expands a chained overlap through every intersecting merge', () => {
    applyMerge(store, wb, history, { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 });
    applyMerge(store, wb, history, { sheet: 0, r0: 1, c0: 4, r1: 2, c1: 5 });
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 3 }))).toEqual({
      sheet: 0,
      r0: 0,
      c0: 3,
      r1: 2,
      c1: 5,
    });
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

  it('removes every merge intersecting a partial selection but leaves unrelated merges', () => {
    const first: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    const second: Range = { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 };
    const unrelated: Range = { sheet: 0, r0: 4, c0: 0, r1: 5, c1: 1 };
    applyMerge(store, wb, history, first);
    applyMerge(store, wb, history, second);
    applyMerge(store, wb, history, unrelated);

    expect(applyUnmerge(store, wb, history, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 3 })).toBe(true);
    expect(store.getState().merges.byAnchor.has(addrKey({ sheet: 0, row: 0, col: 0 }))).toBe(false);
    expect(store.getState().merges.byAnchor.has(addrKey({ sheet: 0, row: 0, col: 3 }))).toBe(false);
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 4, col: 0 }))).toEqual(
      unrelated,
    );
  });

  it('checks the full intersecting merge for protection, not only the selection', () => {
    const merged: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    applyMerge(store, wb, history, merged);
    setCellLocked(store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, true);
    setProtectedSheet(store, 0, true);

    expect(applyUnmerge(store, wb, history, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 })).toBe(false);
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual(
      merged,
    );
  });
});

describe('applyMergeAcross', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    history = new History();
    wb.attachHistory(history);
  });

  it('retains one value per row and is undoable as one transaction', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 1, value: 'top' },
      { row: 1, col: 0, value: 'bottom' },
    ]);
    seedBoolAndMirror(store, wb, 1, 2, false);
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 2 };

    expect(applyMergeAcross(store, wb, history, range)).toBe(true);
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 2,
    });
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 1, col: 0 }))).toEqual({
      sheet: 0,
      r0: 1,
      c0: 0,
      r1: 1,
      c1: 2,
    });
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'top' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'bottom' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({ kind: 'blank' });

    expect(history.undo()).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'top' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({ kind: 'bool', value: false });

    expect(history.redo()).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(2);
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({ kind: 'blank' });
  });

  it('pre-unmerges intersecting vertical merges before making row merges', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'top' }]);
    applyMerge(store, wb, history, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 };

    expect(applyMergeAcross(store, wb, history, range)).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(3);
    expect([...store.getState().merges.byAnchor.values()]).toEqual(
      expect.arrayContaining([
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
        { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 2 },
        { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 2 },
      ]),
    );
  });
});

describe('mergeWillLoseData', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('is false when only the anchor holds content', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'keep' }]);
    expect(mergeWillLoseData(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 })).toBe(
      false,
    );
  });

  it('is true when a non-anchor cell holds content', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'keep' },
      { row: 1, col: 1, value: 'drop' },
    ]);
    expect(mergeWillLoseData(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 })).toBe(
      true,
    );
  });

  it('checks materialized content only for huge ranges', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'keep' },
      { row: 900_000, col: 0, value: 'drop' },
    ]);

    expect(
      mergeWillLoseData(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 }),
    ).toBe(true);
  });

  it('does not warn when a sole non-anchor formula will be promoted', () => {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set('0:0:1', { value: { kind: 'blank' }, formula: '=1+1' });
      return { ...s, data: { ...s.data, cells } };
    });
    expect(mergeWillLoseData(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 })).toBe(
      false,
    );
  });

  it('warns only when more than one value would be retained or discarded', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 1, value: 0 },
      { row: 1, col: 0, value: 'text' },
    ]);
    expect(mergeWillLoseData(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 })).toBe(
      true,
    );
  });

  it('counts false and zero as row contents for Merge Across warnings', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 1, value: 0 }]);
    seedBoolAndMirror(store, wb, 0, 2, false);
    expect(
      mergeAcrossWillLoseData(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 2 }),
    ).toBe(true);

    expect(
      mergeAcrossWillLoseData(store.getState(), { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 2 }),
    ).toBe(false);
  });
});

describe('merge protection gate (H-32)', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    history = new History();
    wb.attachHistory(history);
  });

  it('refuses to merge when a covered cell is locked on a protected sheet', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'a' },
      { row: 0, col: 1, value: 'b' },
    ]);
    setProtectedSheet(store, 0, true);
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    expect(applyMerge(store, wb, history, r)).toBe(false);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    // The locked non-anchor value must survive (no silent clear).
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'b' });
  });

  it('merges when the covered cells are explicitly unlocked on a protected sheet', () => {
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    setCellLocked(store, r, false);
    setProtectedSheet(store, 0, true);
    expect(applyMerge(store, wb, history, r)).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(1);
  });

  it('refuses to unmerge a locked region on a protected sheet', () => {
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    applyMerge(store, wb, history, r);
    setProtectedSheet(store, 0, true);
    expect(applyUnmerge(store, wb, history, r)).toBe(false);
    expect(store.getState().merges.byAnchor.size).toBe(1);
  });

  it('checks existing merges before protection for huge unmerge no-ops', () => {
    setProtectedSheet(store, 0, true);

    expect(applyUnmerge(store, wb, history, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 })).toBe(
      false,
    );
    expect(store.getState().merges.byAnchor.size).toBe(0);
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
