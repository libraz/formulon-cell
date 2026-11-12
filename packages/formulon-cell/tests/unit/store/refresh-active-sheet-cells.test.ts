import { describe, expect, it } from 'vitest';
import type { Addr, CellValue } from '../../../src/engine/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

type Entry = { addr: Addr; value: CellValue; formula: string | null };

describe('refreshActiveSheetCells', () => {
  it("replaces the store's cells with the active sheet's cells from the source", () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 9, col: 9 }, { kind: 'number', value: 7 });
    store.setState((s) => ({ ...s, data: { ...s.data, sheetIndex: 2 } }));
    const asked: number[] = [];
    const source = {
      cells(sheet: number): Entry[] {
        asked.push(sheet);
        return [
          { addr: { sheet, row: 0, col: 1 }, value: { kind: 'number', value: 3 }, formula: '=1+2' },
        ];
      },
    };
    mutators.refreshActiveSheetCells(store, source);
    expect(asked).toEqual([2]);
    expect([...store.getState().data.cells.entries()]).toEqual([
      ['2:0:1', { value: { kind: 'number', value: 3 }, formula: '=1+2' }],
    ]);
  });
});
