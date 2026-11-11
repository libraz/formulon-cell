import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { moveSelectionTo } from '../../../src/commands/move-cells.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

describe('moveSelectionTo', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let history: History;

  beforeEach(async () => {
    // Excel fixture: A1 1, B1 2, A2 3, B2 "x", E5 "occupied", C10 =A1+B1.
    wb = await WorkbookHandle.createDefault();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 3);
    wb.setText({ sheet: 0, row: 1, col: 1 }, 'x');
    wb.setText({ sheet: 0, row: 4, col: 4 }, 'occupied');
    wb.setFormula({ sheet: 0, row: 9, col: 2 }, '=A1+B1');
    wb.recalc();
    store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    history = new History();
    wb.attachHistory(history);
  });

  afterEach(() => wb.dispose());

  it('moves the block, clears the source and rewrites formulas that point at it', async () => {
    // Excel: dropped on G2, A1:B2 lands at G2:H3 and C10 =A1+B1 becomes =G2+H2.
    const confirm = vi.fn(async () => true);
    expect(
      await moveSelectionTo(
        { store, wb, history, commandId: 'moveCells', confirmOverwrite: confirm },
        { row: 1, col: 6 },
      ),
    ).toBe(true);
    expect(confirm).not.toHaveBeenCalled();
    expect(wb.getValue({ sheet: 0, row: 1, col: 6 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 7 })).toEqual({ kind: 'text', value: 'x' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.cellFormula({ sheet: 0, row: 9, col: 2 })).toBe('=G2+H2');
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 1, c0: 6, r1: 2, c1: 7 });
    expect(history.undo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
  });

  it('asks before overwriting data and leaves both ranges alone on Cancel', async () => {
    // Excel: dropping onto D5:E6 over E5 asks 「既にデータがありますが、置き換えますか?」.
    const cancel = vi.fn(async () => false);
    expect(
      await moveSelectionTo(
        { store, wb, history, commandId: 'moveCells', confirmOverwrite: cancel },
        { row: 4, col: 3 },
      ),
    ).toBe(false);
    expect(cancel).toHaveBeenCalledOnce();
    expect(wb.getValue({ sheet: 0, row: 4, col: 4 })).toEqual({ kind: 'text', value: 'occupied' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });

    const ok = vi.fn(async () => true);
    expect(
      await moveSelectionTo(
        { store, wb, history, commandId: 'moveCells', confirmOverwrite: ok },
        { row: 4, col: 3 },
      ),
    ).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 4, col: 4 })).toEqual({ kind: 'number', value: 2 });
    expect(wb.getValue({ sheet: 0, row: 5, col: 4 })).toEqual({ kind: 'text', value: 'x' });
  });

  it('does nothing when dropped where it already is', async () => {
    const confirm = vi.fn(async () => true);
    expect(
      await moveSelectionTo(
        { store, wb, history, commandId: 'moveCells', confirmOverwrite: confirm },
        { row: 0, col: 0 },
      ),
    ).toBe(false);
    expect(history.canUndo()).toBe(false);
  });
});
