import { describe, expect, it } from 'vitest';

import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  colName,
  formatSelectionRef,
  lookupDefinedName,
  parseCellRef,
  parseRangeRef,
} from '../../../src/mount/ref-utils.js';

describe('mount/ref-utils', () => {
  describe('colName', () => {
    it('maps single-letter columns', () => {
      expect(colName(0)).toBe('A');
      expect(colName(1)).toBe('B');
      expect(colName(25)).toBe('Z');
    });

    it('maps two-letter columns at the boundary', () => {
      expect(colName(26)).toBe('AA');
      expect(colName(27)).toBe('AB');
      expect(colName(51)).toBe('AZ');
      expect(colName(52)).toBe('BA');
    });

    it('maps three-letter columns up to XFD (16383)', () => {
      expect(colName(701)).toBe('ZZ');
      expect(colName(702)).toBe('AAA');
      expect(colName(16383)).toBe('XFD');
    });
  });

  describe('formatSelectionRef', () => {
    const active = { row: 0, col: 0 };

    it('shows a single cell when the range collapses', () => {
      expect(formatSelectionRef({ r0: 4, c0: 1, r1: 4, c1: 1 }, { row: 4, col: 1 }, false)).toBe(
        'B5',
      );
      expect(formatSelectionRef({ r0: 4, c0: 1, r1: 4, c1: 1 }, { row: 4, col: 1 }, true)).toBe(
        'R5C2',
      );
    });

    it('shows column-only refs for full-column selection in A1 mode', () => {
      expect(
        formatSelectionRef({ r0: 0, c0: 2, r1: 1048575, c1: 2 }, { row: 0, col: 2 }, false),
      ).toBe('C');
      expect(
        formatSelectionRef({ r0: 0, c0: 0, r1: 1048575, c1: 2 }, { row: 0, col: 0 }, false),
      ).toBe('A:C');
    });

    it('shows row-only refs for full-row selection in A1 mode', () => {
      expect(
        formatSelectionRef({ r0: 4, c0: 0, r1: 4, c1: 16383 }, { row: 4, col: 0 }, false),
      ).toBe('5');
      expect(
        formatSelectionRef({ r0: 4, c0: 0, r1: 9, c1: 16383 }, { row: 4, col: 0 }, false),
      ).toBe('5:10');
    });

    it('falls back to A1:B2 form for arbitrary ranges', () => {
      expect(formatSelectionRef({ r0: 0, c0: 0, r1: 1, c1: 1 }, active, false)).toBe('A1:B2');
      expect(formatSelectionRef({ r0: 0, c0: 0, r1: 1, c1: 1 }, active, true)).toBe('R1C1:R2C2');
    });
  });

  describe('parseCellRef', () => {
    it('parses A1 form', () => {
      expect(parseCellRef('A1')).toEqual({ row: 0, col: 0 });
      expect(parseCellRef('B10')).toEqual({ row: 9, col: 1 });
      expect(parseCellRef('XFD1048576')).toEqual({ row: 1048575, col: 16383 });
    });

    it('parses R1C1 form', () => {
      expect(parseCellRef('R1C1')).toEqual({ row: 0, col: 0 });
      expect(parseCellRef('R10C2')).toEqual({ row: 9, col: 1 });
    });

    it('strips dollar anchors', () => {
      expect(parseCellRef('$A$1')).toEqual({ row: 0, col: 0 });
      expect(parseCellRef('$B10')).toEqual({ row: 9, col: 1 });
    });

    it('is case-insensitive and trims whitespace', () => {
      expect(parseCellRef('  a1 ')).toEqual({ row: 0, col: 0 });
      expect(parseCellRef('z9')).toEqual({ row: 8, col: 25 });
    });

    it('rejects out-of-bounds and malformed references', () => {
      expect(parseCellRef('A0')).toBeNull();
      expect(parseCellRef('XFE1')).toBeNull();
      expect(parseCellRef('A1048577')).toBeNull();
      expect(parseCellRef('R0C1')).toBeNull();
      expect(parseCellRef('NOTAREF')).toBeNull();
      expect(parseCellRef('')).toBeNull();
    });
  });

  describe('parseRangeRef', () => {
    it('parses A1:B2', () => {
      expect(parseRangeRef('A1:B2')).toEqual({ r0: 0, c0: 0, r1: 1, c1: 1 });
    });

    it('normalises reversed ranges', () => {
      expect(parseRangeRef('B2:A1')).toEqual({ r0: 0, c0: 0, r1: 1, c1: 1 });
    });

    it('parses whole-column ranges', () => {
      expect(parseRangeRef('A:A')).toEqual({ r0: 0, c0: 0, r1: 1048575, c1: 0 });
      expect(parseRangeRef('A:C')).toEqual({ r0: 0, c0: 0, r1: 1048575, c1: 2 });
      expect(parseRangeRef('$B')).toEqual({ r0: 0, c0: 1, r1: 1048575, c1: 1 });
    });

    it('parses whole-row ranges', () => {
      expect(parseRangeRef('5:10')).toEqual({ r0: 4, c0: 0, r1: 9, c1: 16383 });
      expect(parseRangeRef('5')).toEqual({ r0: 4, c0: 0, r1: 4, c1: 16383 });
    });

    it('returns null for unparsable input', () => {
      expect(parseRangeRef('A1:B2:C3')).toBeNull();
      expect(parseRangeRef('A1:foo')).toBeNull();
      expect(parseRangeRef('1A:2B')).toBeNull();
    });
  });

  describe('lookupDefinedName', () => {
    const fakeWb = (rows: { name: string; formula: string }[]): WorkbookHandle =>
      ({
        definedNames: () => rows,
      }) as unknown as WorkbookHandle;

    it('returns null when the query is empty', () => {
      expect(lookupDefinedName(fakeWb([]), '')).toBeNull();
    });

    it('matches case-insensitively', () => {
      const wb = fakeWb([{ name: 'Sales', formula: '=Sheet1!$A$1:$B$3' }]);
      expect(lookupDefinedName(wb, 'sales')).toBe('A1:B3');
      expect(lookupDefinedName(wb, 'SALES')).toBe('A1:B3');
    });

    it('strips leading equals, sheet qualifier, and $ anchors', () => {
      const wb = fakeWb([{ name: 'Range1', formula: '=$A$1:$B$5' }]);
      expect(lookupDefinedName(wb, 'Range1')).toBe('A1:B5');
    });

    it('returns null when the name is not registered', () => {
      const wb = fakeWb([{ name: 'Foo', formula: '=A1' }]);
      expect(lookupDefinedName(wb, 'bar')).toBeNull();
    });
  });
});
