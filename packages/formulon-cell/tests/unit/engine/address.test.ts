import { describe, expect, it } from 'vitest';

import {
  addrKey,
  colFromLetters,
  colLetter,
  formatA1Cell,
  formatA1Range,
  MAX_COL,
  MAX_ROW,
  parseA1Atom,
  parseAddrKey,
} from '../../../src/engine/address.js';

describe('engine/address', () => {
  it('emits a sheet:row:col string key', () => {
    expect(addrKey({ sheet: 0, row: 0, col: 0 })).toBe('0:0:0');
    expect(addrKey({ sheet: 2, row: 5, col: 8 })).toBe('2:5:8');
  });

  it('distinguishes addresses that differ only in one component', () => {
    expect(addrKey({ sheet: 0, row: 0, col: 0 })).not.toBe(addrKey({ sheet: 1, row: 0, col: 0 }));
    expect(addrKey({ sheet: 0, row: 0, col: 0 })).not.toBe(addrKey({ sheet: 0, row: 1, col: 0 }));
    expect(addrKey({ sheet: 0, row: 0, col: 0 })).not.toBe(addrKey({ sheet: 0, row: 0, col: 1 }));
  });

  it('handles the worksheet upper bound', () => {
    expect(addrKey({ sheet: 0, row: MAX_ROW, col: MAX_COL })).toBe('0:1048575:16383');
  });

  it('pins the zero-based worksheet bounds', () => {
    expect(MAX_ROW).toBe(1_048_575);
    expect(MAX_COL).toBe(16_383);
  });

  describe('colLetter', () => {
    it('maps single-letter columns', () => {
      expect(colLetter(0)).toBe('A');
      expect(colLetter(1)).toBe('B');
      expect(colLetter(25)).toBe('Z');
    });

    it('maps two-letter columns at the boundary', () => {
      expect(colLetter(26)).toBe('AA');
      expect(colLetter(27)).toBe('AB');
      expect(colLetter(51)).toBe('AZ');
      expect(colLetter(52)).toBe('BA');
      expect(colLetter(701)).toBe('ZZ');
    });

    it('maps three-letter columns up to the last column', () => {
      expect(colLetter(702)).toBe('AAA');
      expect(colLetter(MAX_COL)).toBe('XFD');
    });
  });

  describe('colFromLetters', () => {
    it('inverts colLetter across the sheet width', () => {
      for (const col of [0, 25, 26, 51, 52, 701, 702, MAX_COL]) {
        expect(colFromLetters(colLetter(col))).toBe(col);
      }
    });

    it('is case-insensitive', () => {
      expect(colFromLetters('xfd')).toBe(MAX_COL);
      expect(colFromLetters('aB')).toBe(27);
    });

    it('rejects empty input and non-letters', () => {
      expect(colFromLetters('')).toBe(-1);
      expect(colFromLetters('A1')).toBe(-1);
      expect(colFromLetters('$A')).toBe(-1);
    });

    it('does not clamp past the last column', () => {
      expect(colFromLetters('XFE')).toBe(MAX_COL + 1);
    });
  });

  describe('parseA1Atom', () => {
    it('parses relative, absolute, and mixed atoms case-insensitively', () => {
      expect(parseA1Atom('B2')).toEqual({ row: 1, col: 1 });
      expect(parseA1Atom('$A$1')).toEqual({ row: 0, col: 0 });
      expect(parseA1Atom('c$3')).toEqual({ row: 2, col: 2 });
      expect(parseA1Atom('  $d4 ')).toEqual({ row: 3, col: 3 });
    });

    it('accepts the last cell and rejects anything beyond it', () => {
      expect(parseA1Atom('XFD1048576')).toEqual({ row: MAX_ROW, col: MAX_COL });
      expect(parseA1Atom('XFE1')).toBeNull();
      expect(parseA1Atom('A1048577')).toBeNull();
    });

    it('rejects row 0, malformed input, and sheet prefixes', () => {
      expect(parseA1Atom('A0')).toBeNull();
      expect(parseA1Atom('1A')).toBeNull();
      expect(parseA1Atom('A$$1')).toBeNull();
      expect(parseA1Atom('Sheet1!A1')).toBeNull();
      expect(parseA1Atom('')).toBeNull();
    });

    it('reads leading-zero rows as their numeric value', () => {
      expect(parseA1Atom('A01')).toEqual({ row: 0, col: 0 });
    });
  });

  describe('formatA1Cell', () => {
    it('formats relative and absolute cells', () => {
      expect(formatA1Cell(2, 1)).toBe('B3');
      expect(formatA1Cell(2, 1, true)).toBe('$B$3');
    });

    it('uses multi-letter columns', () => {
      expect(formatA1Cell(0, 27)).toBe('AB1');
    });
  });

  describe('formatA1Range', () => {
    const r = { r0: 0, c0: 0, r1: 2, c1: 27 };

    it('formats relative and absolute ranges', () => {
      expect(formatA1Range(r)).toBe('A1:AB3');
      expect(formatA1Range(r, { absolute: true })).toBe('$A$1:$AB$3');
    });

    it('collapses a single cell by default', () => {
      const one = { r0: 4, c0: 2, r1: 4, c1: 2 };
      expect(formatA1Range(one)).toBe('C5');
      expect(formatA1Range(one, { absolute: true })).toBe('$C$5');
    });

    it('keeps the colon when collapse is false', () => {
      const one = { r0: 4, c0: 2, r1: 4, c1: 2 };
      expect(formatA1Range(one, { collapse: false })).toBe('C5:C5');
      expect(formatA1Range(one, { absolute: true, collapse: false })).toBe('$C$5:$C$5');
    });
  });

  describe('parseAddrKey', () => {
    it('inverts addrKey', () => {
      const addr = { sheet: 2, row: 14, col: 7 };
      expect(parseAddrKey(addrKey(addr))).toEqual(addr);
    });

    it('rejects malformed keys', () => {
      for (const key of ['', '1:2', '1:2:3:4', 'a:1:2', '0:1.5:2', '0::2', '0:1:', '0:NaN:1']) {
        expect(parseAddrKey(key), key).toBeNull();
      }
    });
  });
});
