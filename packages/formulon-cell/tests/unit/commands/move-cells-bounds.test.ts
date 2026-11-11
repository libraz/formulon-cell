import { describe, expect, it, vi } from 'vitest';
import {
  moveSelectionTo,
  moveTargetFits,
  moveTargetRange,
} from '../../../src/commands/move-cells.js';
import { MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('move target bounds', () => {
  it('reports whether the shifted block stays inside the grid', () => {
    const src = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    expect(moveTargetFits(moveTargetRange(src, { row: MAX_ROW - 1, col: MAX_COL - 1 }))).toBe(true);
    expect(moveTargetFits(moveTargetRange(src, { row: MAX_ROW, col: 0 }))).toBe(false);
    expect(moveTargetFits(moveTargetRange(src, { row: 0, col: MAX_COL }))).toBe(false);
  });

  it('refuses a drop that would run past the last row without touching the source', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    const confirm = vi.fn(async () => true);
    const moved = await moveSelectionTo(
      { store, wb, history: null, commandId: 'moveCells', confirmOverwrite: confirm },
      { row: MAX_ROW, col: 0 },
    );
    expect(moved).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
  });
});
