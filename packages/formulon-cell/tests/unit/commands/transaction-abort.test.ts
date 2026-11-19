import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { applyMerge } from '../../../src/commands/merge.js';
import { executeQuickAnalysisAction } from '../../../src/commands/quick-analysis.js';
import { addSheetAfterActive } from '../../../src/commands/sheet-mutate.js';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const at = (row: number, col: number): Addr => ({ sheet: 0, row, col });

/** Lets the first `passes` calls of `method` through and throws once on the next; the rollback's
 *  own calls go through again. */
const failAfter = <K extends 'setFormula' | 'setBlank' | 'moveSheet'>(
  wb: WorkbookHandle,
  method: K,
  passes: number,
): Error => {
  const original = (wb[method] as (...args: unknown[]) => unknown).bind(wb);
  const injected = new Error(`injected ${method} failure`);
  let calls = 0;
  vi.spyOn(wb, method).mockImplementation(((...args: unknown[]) => {
    calls += 1;
    if (calls === passes + 1) throw injected;
    return original(...args);
  }) as never);
  return injected;
};

describe('a command that throws partway leaves the sheet and history unchanged', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    history = new History();
    wb = await WorkbookHandle.createDefault();
    wb.attachHistory(history);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    wb.dispose();
  });

  it('Quick Analysis totals roll back the totals already written', () => {
    for (let row = 1; row <= 3; row += 1) {
      wb.setNumber(at(row, 0), row);
      wb.setNumber(at(row, 1), row * 10);
    }
    history.clear();
    const injected = failAfter(wb, 'setFormula', 1);

    expect(() =>
      executeQuickAnalysisAction({
        store,
        wb,
        history,
        actionId: 'totals-sum-row',
        range: { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 },
        stats: { cells: 6, numericCount: 6, nonBlankCount: 6, sum: 66, avg: 11, min: 1, max: 30 },
      }),
    ).toThrow(injected);

    expect(wb.cellFormula(at(4, 0))).toBeNull();
    expect(wb.cellFormula(at(4, 1))).toBeNull();
    expect(wb.getValue(at(4, 0)).kind).toBe('blank');
    expect(history.canUndo()).toBe(false);
    expect(history.isInTransaction()).toBe(false);
  });

  it('Merge rolls back the promoted anchor value', () => {
    wb.setText(at(0, 1), 'moved');
    mutators.replaceCells(store, wb.cells(0));
    history.clear();
    const injected = failAfter(wb, 'setBlank', 0);

    expect(() => applyMerge(store, wb, history, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 })).toThrow(
      injected,
    );

    expect(wb.getValue(at(0, 0)).kind).toBe('blank');
    expect(wb.getValue(at(0, 1))).toEqual({ kind: 'text', value: 'moved' });
    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(history.canUndo()).toBe(false);
    expect(history.isInTransaction()).toBe(false);
  });

  it('Insert Sheet removes the added sheet when placing it fails', () => {
    wb.addSheet('Sheet2');
    history.clear();
    const before = wb.sheetCount;
    const injected = failAfter(wb, 'moveSheet', 0);

    expect(() => addSheetAfterActive(store, wb, history)).toThrow(injected);

    expect(wb.sheetCount).toBe(before);
    expect(history.canUndo()).toBe(false);
    expect(history.isInTransaction()).toBe(false);
  });
});
