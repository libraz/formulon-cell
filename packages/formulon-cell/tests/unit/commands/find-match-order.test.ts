import { describe, expect, it } from 'vitest';
import { findNext } from '../../../src/commands/find.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const LAST_ROW = 1_048_575;

describe('findNext ordering', () => {
  it('orders a last-row match on an earlier sheet before any match on the next sheet', () => {
    const early = { sheet: 0, row: LAST_ROW, col: 0 };
    const late = { sheet: 1, row: 5, col: 0 };
    const text = { kind: 'text', value: 'x' };
    const wb = {
      sheetCount: 2,
      cells: (sheet: number) => [{ addr: sheet === 0 ? early : late, value: text, formula: null }],
      getDisplayText: () => 'x',
      getValue: () => text,
      cellFormula: () => null,
      capabilities: {},
    } as unknown as WorkbookHandle;
    const state = createSpreadsheetStore().getState();
    const opts = { query: 'x', within: 'workbook' as const, workbook: wb };

    expect(findNext(state, opts, { sheet: 1, row: 0, col: 0 }, 'prev')?.addr).toEqual(early);
    expect(findNext(state, opts, early, 'next')?.addr).toEqual(late);
  });
});
