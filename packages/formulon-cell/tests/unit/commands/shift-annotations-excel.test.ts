import { beforeEach, describe, expect, it } from 'vitest';
import { deleteCells, insertCells } from '../../../src/commands/cell-shift.js';
import { deleteCols, deleteRows, insertCols, insertRows } from '../../../src/commands/structure.js';
import { MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
import type { Range } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { validationAt } from '../../../src/store/range-validations.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { newWb } from './structure/fixtures.js';

// Measured against desktop Excel: a grid B2:G9 with merges B3:D3, F2:F5, B6:C7,
// cell-value rules on C3:E6 (> 50) and G2:G9 (< 75), and validations on D4:F7
// (whole numbers) and B8:C9 (decimals).

const cellKey = (row: number, col: number): string => `${String.fromCharCode(65 + col)}${row + 1}`;
const cellAt = (a1: string): { row: number; col: number } => ({
  col: a1.charCodeAt(0) - 65,
  row: Number(a1.slice(1)) - 1,
});
/** "D5" or "C3:E6" as a range on sheet 0. */
const rect = (a1: string): Range => {
  const [from = '', to = from] = a1.split(':');
  const a = cellAt(from);
  const b = cellAt(to);
  return { sheet: 0, r0: a.row, c0: a.col, r1: b.row, c1: b.col };
};
const rows = (r0: number, r1: number): Range => ({ sheet: 0, r0, c0: 0, r1, c1: MAX_COL });
const cols = (c0: number, c1: number): Range => ({ sheet: 0, r0: 0, c0, r1: MAX_ROW, c1 });

const cellsOf = (ranges: readonly Range[]): string[] => {
  const out = new Set<string>();
  for (const r of ranges)
    for (let row = r.r0; row <= r.r1; row += 1)
      for (let col = r.c0; col <= r.c1; col += 1) out.add(cellKey(row, col));
  return [...out].sort();
};
const expandCells = (a1s: string[]): string[] => cellsOf(a1s.map(rect));

const mergesOf = (store: SpreadsheetStore): string[] =>
  Array.from(store.getState().merges.byAnchor.values())
    .map((m) => `${cellKey(m.r0, m.c0)}:${cellKey(m.r1, m.c1)}`)
    .sort();

/** Cells covered by the rules whose threshold is `a` (the rules split into areas keep it). */
const cfCells = (store: SpreadsheetStore, a: number): string[] =>
  cellsOf(
    store
      .getState()
      .conditional.rules.filter((rule) => rule.kind === 'cell-value' && rule.a === a)
      .map((rule) => rule.range),
  );

const dvCells = (store: SpreadsheetStore, kind: 'whole' | 'decimal'): string[] => {
  const out: string[] = [];
  for (let row = 0; row < 16; row += 1)
    for (let col = 0; col < 14; col += 1) {
      const v = validationAt(store.getState().format, { sheet: 0, row, col });
      if (v?.kind === kind) out.push(cellKey(row, col));
    }
  return out.sort();
};

describe('cell and axis shifts keep merges, conditional formats and validations like desktop Excel', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  const select = (range: Range): void => {
    mutators.setActive(store, { sheet: 0, row: range.r0, col: range.c0 });
    mutators.extendRangeTo(store, { sheet: 0, row: range.r1, col: range.c1 });
  };

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    for (let row = 1; row <= 8; row += 1)
      for (let col = 1; col <= 6; col += 1)
        wb.setNumber({ sheet: 0, row, col }, (row + 1) * 10 + col + 1);
    for (const m of ['B3:D3', 'F2:F5', 'B6:C7']) mutators.mergeRange(store, rect(m));
    mutators.addConditionalRule(store, {
      kind: 'cell-value',
      range: rect('C3:E6'),
      op: '>',
      a: 50,
      apply: { bold: true },
    });
    mutators.addConditionalRule(store, {
      kind: 'cell-value',
      range: rect('G2:G9'),
      op: '<',
      a: 75,
      apply: { bold: true },
    });
    const blocks = [
      { range: rect('D4:F7'), validation: { kind: 'whole', op: 'between', a: 1, b: 100 } },
      { range: rect('B8:C9'), validation: { kind: 'decimal', op: '>', a: 0 } },
    ] as const;
    for (const { range, validation } of blocks)
      for (let row = range.r0; row <= range.r1; row += 1)
        for (let col = range.c0; col <= range.c1; col += 1)
          mutators.setCellFormat(store, { sheet: 0, row, col }, { validation });
  });

  describe('conditional-format ranges across a cell shift', () => {
    it('delete shift left on one cell leaves the cells outside the band in place', () => {
      select(rect('D5'));
      expect(deleteCells(store, wb, null, rect('D5'), 'left')).toBe(true);
      expect(cfCells(store, 50)).toEqual(expandCells(['C3:E4', 'C6:E6', 'C5:D5']));
      expect(cfCells(store, 75)).toEqual(expandCells(['G2:G4', 'G6:G9', 'F5']));
    });

    it('delete shift up of a block that ends flush with the rule keeps the rule whole', () => {
      select(rect('D5:E6'));
      expect(deleteCells(store, wb, null, rect('D5:E6'), 'up')).toBe(true);
      expect(cfCells(store, 50)).toEqual(expandCells(['C3:E6']));
      expect(cfCells(store, 75)).toEqual(expandCells(['G2:G9']));
    });

    it('insert shift down grows the rule across the inserted cells inside it', () => {
      select(rect('D5:E6'));
      expect(insertCells(store, wb, null, rect('D5:E6'), 'down')).toBe(true);
      expect(cfCells(store, 50)).toEqual(expandCells(['C3:E4', 'C5:C6', 'D5:E8']));
      expect(cfCells(store, 75)).toEqual(expandCells(['G2:G9']));
    });

    it('insert shift right over whole columns inserts columns', () => {
      select(cols(3, 4));
      expect(insertCells(store, wb, null, cols(3, 4), 'right')).toBe(true);
      expect(cfCells(store, 50)).toEqual(expandCells(['C3:G6']));
      expect(cfCells(store, 75)).toEqual(expandCells(['I2:I9']));
    });
  });

  describe('data validation across inserts', () => {
    it('inserting a column inside a validated block validates the new column', () => {
      select(rect('E6'));
      expect(insertCols(store, wb, null, 4, 1)).toBe(true);
      expect(dvCells(store, 'whole')).toEqual(expandCells(['D4:G7']));
      expect(dvCells(store, 'decimal')).toEqual(expandCells(['B8:C9']));
    });

    it('inserting rows inside a validated block validates the new rows', () => {
      select(rows(4, 5));
      expect(insertRows(store, wb, null, 4, 2)).toBe(true);
      expect(dvCells(store, 'whole')).toEqual(expandCells(['D4:F9']));
      expect(dvCells(store, 'decimal')).toEqual(expandCells(['B10:C11']));
    });

    it('insert shift down inside a validated block validates the inserted cells', () => {
      select(rect('D5:E6'));
      expect(insertCells(store, wb, null, rect('D5:E6'), 'down')).toBe(true);
      expect(dvCells(store, 'whole')).toEqual(expandCells(['D4:F7', 'D8:E9']));
      expect(dvCells(store, 'decimal')).toEqual(expandCells(['B8:C9']));
    });

    it('insert shift right over whole columns carries the rule to the new columns', () => {
      select(cols(3, 4));
      expect(insertCells(store, wb, null, cols(3, 4), 'right')).toBe(true);
      expect(dvCells(store, 'whole')).toEqual(expandCells(['F4:H7']));
      expect(dvCells(store, 'decimal')).toEqual(expandCells(['B8:E9']));
    });
  });

  describe('cell shifts that split a merge', () => {
    it('delete shift left drops the merges the band cuts through', () => {
      select(rect('D3'));
      expect(deleteCells(store, wb, null, rect('D3'), 'left')).toBe(true);
      expect(mergesOf(store)).toEqual(['B6:C7']);
      expect(wb.getValue({ sheet: 0, row: 2, col: 3 })).toEqual({ kind: 'number', value: 35 });
    });

    it('delete shift left removes a merge whose cells were deleted', () => {
      select(rect('B3:D3'));
      expect(deleteCells(store, wb, null, rect('B3:D3'), 'left')).toBe(true);
      expect(mergesOf(store)).toEqual(['B6:C7']);
    });

    it('insert shift down drops the merge the band cuts through and keeps the rest', () => {
      select(rect('C4:D5'));
      expect(insertCells(store, wb, null, rect('C4:D5'), 'down')).toBe(true);
      expect(mergesOf(store)).toEqual(['B3:D3', 'F2:F5']);
    });

    it('insert shift right drops the merge that crosses the band row', () => {
      select(rect('C4'));
      expect(insertCells(store, wb, null, rect('C4'), 'right')).toBe(true);
      expect(mergesOf(store)).toEqual(['B3:D3', 'B6:C7']);
    });

    it('delete shift up over whole rows deletes the rows and resizes the merges', () => {
      select(rows(3, 4));
      expect(deleteCells(store, wb, null, rows(3, 4), 'up')).toBe(true);
      expect(mergesOf(store)).toEqual(['B3:D3', 'B4:C5', 'F2:F3']);
    });
  });

  describe('selection after whole-row edits that cross a merge', () => {
    it('inserting rows keeps the selection grown to the merge it crossed', () => {
      select(rows(3, 4));
      expect(insertRows(store, wb, null, 3, 2)).toBe(true);
      const { range } = store.getState().selection;
      expect([range.r0, range.r1]).toEqual([1, 4]);
    });

    it('deleting rows keeps the selection grown to the merge it crossed', () => {
      select(rows(3, 4));
      expect(deleteRows(store, wb, null, 3, 2)).toBe(true);
      const { range } = store.getState().selection;
      expect([range.r0, range.r1]).toEqual([1, 4]);
    });

    it('deleting columns away from every merge leaves the selection alone', () => {
      select(rect('C4:D4'));
      expect(deleteCols(store, wb, null, 2, 2)).toBe(true);
      const { range } = store.getState().selection;
      expect([range.r0, range.c0, range.r1, range.c1]).toEqual([3, 2, 3, 3]);
    });
  });
});
