import { describe, expect, it, vi } from 'vitest';
import { autoSum } from '../../../src/commands/auto-sum.js';
import { insertCopiedCellsFromTSV } from '../../../src/commands/clipboard/insert-copied-cells.js';
import { captureSnapshot } from '../../../src/commands/clipboard/snapshot.js';
import { applyMerge, applyMergeAcross } from '../../../src/commands/merge.js';
import { executeQuickAnalysisAction } from '../../../src/commands/quick-analysis.js';
import type { Range } from '../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

/** Count engine-level recalc passes through a narrow cast on the private engine. */
const spyEngineRecalc = (wb: WorkbookHandle): ReturnType<typeof vi.fn> => {
  const internals = wb as unknown as { wb: { recalc: () => { ok: boolean } } };
  const original = internals.wb.recalc.bind(internals.wb);
  const spy = vi.fn(original);
  internals.wb.recalc = spy;
  return spy;
};

const seedNumber = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  v: number,
) => {
  wb.setNumber({ sheet: 0, row, col }, v);
  store.setState((s) => {
    const map = new Map(s.data.cells);
    map.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value: v },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells: map } };
  });
};

const select = (store: SpreadsheetStore, r0: number, c0: number, r1: number, c1: number) =>
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));

const seedGrid = (store: SpreadsheetStore, wb: WorkbookHandle, rows: number, cols: number) => {
  for (let r = 0; r < rows; r += 1)
    for (let c = 0; c < cols; c += 1) seedNumber(store, wb, r, c, r + c + 1);
};

describe('multi-cell writers recalc once', () => {
  it('autoSum over a 2-D selection writes every total in one recalc', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    seedGrid(store, wb, 2, 2);
    select(store, 0, 0, 2, 2);
    const recalc = spyEngineRecalc(wb);
    expect(autoSum(store.getState(), wb)).not.toBeNull();
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=SUM(A1:A2)');
    expect(recalc).toHaveBeenCalledTimes(1);
  });

  it.each(['totals-sum-row', 'totals-sum-col'] as const)('%s recalcs once', async (actionId) => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    seedGrid(store, wb, 3, 3);
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 };
    const recalc = spyEngineRecalc(wb);
    const res = executeQuickAnalysisAction({
      store,
      wb,
      actionId,
      range,
      stats: { cells: 9, numericCount: 9, nonBlankCount: 9, sum: 0, avg: 0, min: 0, max: 0 },
    });
    expect(res.ok).toBe(true);
    expect(recalc).toHaveBeenCalledTimes(1);
  });

  it('applyMerge recalcs once for anchor write plus clears', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    seedNumber(store, wb, 0, 1, 5);
    seedNumber(store, wb, 1, 0, 6);
    seedNumber(store, wb, 1, 1, 7);
    const recalc = spyEngineRecalc(wb);
    expect(applyMerge(store, wb, null, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 })).toBe(true);
    expect(recalc).toHaveBeenCalledTimes(1);
  });

  it('applyMergeAcross recalcs once across rows', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    seedNumber(store, wb, 0, 1, 5);
    seedNumber(store, wb, 1, 1, 7);
    seedNumber(store, wb, 2, 1, 9);
    const recalc = spyEngineRecalc(wb);
    expect(applyMergeAcross(store, wb, null, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 })).toBe(true);
    expect(recalc).toHaveBeenCalledTimes(1);
  });

  it('insert copied cells recalcs no more for a 4-cell snapshot than for a 1-cell one', async () => {
    const passes = async (r1: number, c1: number): Promise<number> => {
      const store = createSpreadsheetStore();
      const wb = await newWb();
      seedGrid(store, wb, 2, 2);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1, c1 });
      if (snap === null) throw new Error('expected snapshot');
      select(store, 5, 5, 5, 5);
      const recalc = spyEngineRecalc(wb);
      insertCopiedCellsFromTSV(store, wb, null, '1', 'down', snap);
      return recalc.mock.calls.length;
    };
    expect(await passes(1, 1)).toBe(await passes(0, 0));
  });
});
