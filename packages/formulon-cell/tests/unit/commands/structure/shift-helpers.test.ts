import { describe, expect, it } from 'vitest';
import {
  shiftFormatsByCol,
  shiftFormatsByRow,
  shiftIndexedMap,
  shiftIndexedSet,
} from '../../../../src/commands/axis-shift.js';
import { insertRows } from '../../../../src/commands/structure.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { cellText, newWb } from './fixtures.js';

describe('shift helpers', () => {
  describe('shiftIndexedMap', () => {
    it('shifts keys >= split forward by delta>0', () => {
      const m = new Map<number, number>([
        [0, 100],
        [2, 200],
        [5, 500],
      ]);
      const out = shiftIndexedMap(m, 2, 1);
      expect(Array.from(out.entries()).sort()).toEqual([
        [0, 100],
        [3, 200],
        [6, 500],
      ]);
    });

    it('drops keys in [split, split-delta) for delta<0', () => {
      const m = new Map<number, number>([
        [0, 100],
        [2, 200],
        [3, 300],
        [5, 500],
      ]);
      const out = shiftIndexedMap(m, 2, -2);
      // keys 2, 3 are in the deleted band; 5 shifts to 3.
      expect(Array.from(out.entries()).sort()).toEqual([
        [0, 100],
        [3, 500],
      ]);
    });

    it('returns an empty map when input is empty', () => {
      const out = shiftIndexedMap(new Map(), 0, 5);
      expect(out.size).toBe(0);
    });
  });

  describe('shiftIndexedSet', () => {
    it('shifts values >= split forward', () => {
      const s = new Set([0, 2, 5]);
      const out = shiftIndexedSet(s, 2, 1);
      expect(Array.from(out).sort()).toEqual([0, 3, 6]);
    });

    it('drops values in deleted band', () => {
      const s = new Set([1, 2, 3, 5]);
      const out = shiftIndexedSet(s, 2, -2);
      expect(Array.from(out).sort()).toEqual([1, 3]);
    });
  });

  describe('shiftFormatsByRow', () => {
    it('shifts only the targeted sheet', () => {
      const m = new Map([
        ['0:0:0', { bold: true }],
        ['0:5:1', { italic: true }],
        ['1:5:1', { underline: true }], // different sheet — untouched
      ]);
      const out = shiftFormatsByRow(m, 0, 3, 2);
      // sheet 0, row 5 → row 7. Sheet 1 untouched.
      expect(out.get('0:0:0')?.bold).toBe(true);
      expect(out.get('0:7:1')?.italic).toBe(true);
      expect(out.get('0:5:1')).toBeUndefined();
      expect(out.get('1:5:1')?.underline).toBe(true);
    });

    it('drops formats in deleted band on negative delta', () => {
      const m = new Map([
        ['0:2:0', { bold: true }],
        ['0:3:0', { italic: true }],
        ['0:5:0', { underline: true }],
      ]);
      const out = shiftFormatsByRow(m, 0, 2, -2);
      // rows 2 and 3 are in deleted band; row 5 → 3.
      expect(out.get('0:2:0')).toBeUndefined();
      expect(out.get('0:3:0')?.underline).toBe(true);
    });
  });

  describe('shiftFormatsByCol', () => {
    it('shifts cols on targeted sheet only', () => {
      const m = new Map([
        ['0:0:5', { bold: true }],
        ['1:0:5', { italic: true }],
      ]);
      const out = shiftFormatsByCol(m, 0, 3, 2);
      expect(out.get('0:0:7')?.bold).toBe(true);
      expect(out.get('1:0:5')?.italic).toBe(true);
    });
  });
});

describe('integration: format + cell shift consistency', () => {
  it('insertRows preserves the format on the shifted cell', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setText({ sheet: 0, row: 5, col: 2 }, 'hello');
    mutators.setCellFormat(store, { sheet: 0, row: 5, col: 2 }, { bold: true });
    insertRows(store, wb, null, 0, 2);
    expect(cellText(wb, 0, 7, 2)).toBe('hello');
    expect(store.getState().format.formats.get('0:7:2')?.bold).toBe(true);
  });
});
