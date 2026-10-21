import { describe, expect, it } from 'vitest';

import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  formatSelectionRef,
  lookupDefinedName,
  nameForRange,
  parseCellRef,
  parseRangeRef,
  rangeDefinedNames,
  resolveAreaText,
  splitAreas,
  splitSheetQualifier,
} from '../../../src/mount/ref-utils.js';

describe('mount/ref-utils', () => {
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

    it('treats a band running past the sheet edge as a whole row/column ref', () => {
      expect(
        formatSelectionRef({ r0: 0, c0: 2, r1: 1048576, c1: 2 }, { row: 0, col: 2 }, false),
      ).toBe('C');
      expect(
        formatSelectionRef({ r0: 4, c0: 0, r1: 4, c1: 16384 }, { row: 4, col: 0 }, false),
      ).toBe('5');
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

    it('reads a leading-zero row as its numeric value', () => {
      expect(parseCellRef('A01')).toEqual({ row: 0, col: 0 });
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
    const fakeWb = (
      rows: { name: string; formula: string; localSheetId?: number }[],
    ): WorkbookHandle =>
      ({
        definedNames: () => rows.map((row) => ({ ...row, localSheetId: row.localSheetId ?? -1 })),
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

    it('prefers the active sheet scoped name over a workbook scoped name', () => {
      const wb = fakeWb([
        { name: 'Sales', formula: '=Sheet1!$A$1:$A$3', localSheetId: -1 },
        { name: 'Sales', formula: '=Sheet1!$B$1:$B$3', localSheetId: 0 },
        { name: 'Sales', formula: '=Sheet2!$C$1:$C$3', localSheetId: 1 },
      ]);

      expect(lookupDefinedName(wb, 'Sales', 0)).toBe('B1:B3');
      expect(lookupDefinedName(wb, 'Sales', 1)).toBe('C1:C3');
      expect(lookupDefinedName(wb, 'Sales', 2)).toBe('A1:A3');
    });
  });

  describe('sheet-aware names and references', () => {
    const sheets = ['Sheet1', 'My Sheet'];
    const wbWith = (
      rows: { name: string; formula: string; localSheetId?: number }[],
    ): WorkbookHandle =>
      ({
        sheetCount: sheets.length,
        sheetName: (i: number) => sheets[i],
        definedNames: () => rows.map((row) => ({ ...row, localSheetId: row.localSheetId ?? -1 })),
      }) as unknown as WorkbookHandle;

    it('splits quoted and plain sheet qualifiers and comma areas', () => {
      expect(splitSheetQualifier("'My ''Sheet'!A1")).toEqual({ sheet: "My 'Sheet", ref: 'A1' });
      expect(splitSheetQualifier('Sheet2!B4')).toEqual({ sheet: 'Sheet2', ref: 'B4' });
      expect(splitSheetQualifier('A1')).toEqual({ sheet: null, ref: 'A1' });
      expect(splitAreas("A1,'a,b'!C3:D4")).toEqual(['A1', "'a,b'!C3:D4"]);
    });

    it('resolves sheet-qualified references and names to a sheet index', () => {
      const wb = wbWith([{ name: 'Remote', formula: "='My Sheet'!$A$1:$B$3" }]);
      expect(resolveAreaText(wb, "'my sheet'!B4", 0)).toEqual({
        sheet: 1,
        r0: 3,
        c0: 1,
        r1: 3,
        c1: 1,
      });
      expect(resolveAreaText(wb, 'REMOTE', 0)).toEqual({ sheet: 1, r0: 0, c0: 0, r1: 2, c1: 1 });
      expect(resolveAreaText(wb, 'Nope!A1', 0)).toBeNull();
    });

    it('lets a name win over column-like text', () => {
      const wb = wbWith([{ name: 'TAX', formula: 'Sheet1!$B$2' }]);
      expect(resolveAreaText(wb, 'tax', 0)).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    });

    it('lists range names only, sorted, skipping other sheets scopes', () => {
      const wb = wbWith([
        { name: 'b', formula: 'Sheet1!$A$1' },
        { name: 'Rate', formula: '0.08' },
        { name: 'A', formula: 'Sheet1!$A$2', localSheetId: 0 },
        { name: 'Hidden', formula: 'Sheet1!$A$3', localSheetId: 1 },
      ]);
      expect(rangeDefinedNames(wb, 0).map((n) => n.name)).toEqual(['A', 'b']);
    });

    it('picks a sheet-scoped name over workbook names for an exact range', () => {
      const wb = wbWith([
        { name: 'zed', formula: 'Sheet1!$A$1' },
        { name: 'abc', formula: 'Sheet1!$A$1' },
        { name: 'local', formula: 'Sheet1!$A$1', localSheetId: 0 },
      ]);
      const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
      expect(nameForRange(wb, 0, range)).toBe('local');
      expect(nameForRange(wb, 1, range)).toBe('abc');
      expect(nameForRange(wb, 0, { ...range, r1: 1 })).toBeNull();
    });
  });
});
