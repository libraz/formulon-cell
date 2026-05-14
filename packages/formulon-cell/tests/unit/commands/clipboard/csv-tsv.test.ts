import { describe, expect, it } from 'vitest';

import { encodeCSV, parseCSV } from '../../../../src/commands/clipboard/csv.js';
import { encodeTSV, parseTSV } from '../../../../src/commands/clipboard/tsv.js';

/**
 * Integration tests across the two text-clipboard parsers. The siblings live
 * in separate files (csv.test.ts / tsv.test.ts) and each locks down one
 * parser's behaviour. This spec asserts the *contract crossings*:
 *  - mixed terminators (\r, \n, \r\n) in the same input,
 *  - quoted cells containing the *other* parser's separator,
 *  - cross-format paste — TSV produced by a Spreadsheet sometimes lands in a
 *    CSV-aware tool, and vice versa,
 *  - 1MB stress shape (length-invariant, not perf — just sanity),
 *  - round-trip between both encoders + parsers for a single grid.
 */
describe('clipboard CSV ↔ TSV integration', () => {
  describe('terminators', () => {
    it('CSV accepts mixed CR / LF / CRLF in a single payload', () => {
      const out = parseCSV('a,b\r\nc,d\ne,f\rg,h');
      expect(out).toEqual([
        ['a', 'b'],
        ['c', 'd'],
        ['e', 'f'],
        ['g', 'h'],
      ]);
    });

    it('TSV accepts mixed CR-vs-LF terminators', () => {
      const out = parseTSV('a\tb\r\nc\td\ne\tf');
      expect(out).toEqual([
        ['a', 'b'],
        ['c', 'd'],
        ['e', 'f'],
      ]);
    });
  });

  describe('cross-separator quoting', () => {
    it('CSV preserves tab characters verbatim inside unquoted cells', () => {
      // Tabs are not special to CSV. They pass through.
      expect(parseCSV('a\tb,c')).toEqual([['a\tb', 'c']]);
    });

    it('TSV preserves commas verbatim inside unquoted cells', () => {
      expect(parseTSV('a,b\tc')).toEqual([['a,b', 'c']]);
    });

    it('CSV escapes its delimiter (comma) inside a quoted cell', () => {
      const grid = [['hello, world', 'second']];
      const csv = encodeCSV(grid);
      expect(csv).toBe('"hello, world",second');
      expect(parseCSV(csv)).toEqual(grid);
    });

    it('TSV escapes its delimiter (tab) inside a quoted cell', () => {
      const grid = [['col\twith\ttabs', 'next']];
      const tsv = encodeTSV(grid);
      expect(tsv).toBe('"col\twith\ttabs"\tnext');
      expect(parseTSV(tsv)).toEqual(grid);
    });
  });

  describe('quoted multiline + embedded delimiter', () => {
    it('CSV: a single quoted field can contain newlines and commas together', () => {
      const grid = [['line1\nline2,with,commas', 'next']];
      const csv = encodeCSV(grid);
      expect(parseCSV(csv)).toEqual(grid);
    });

    it('TSV: a single quoted field can contain newlines and tabs together', () => {
      const grid = [['line1\nline2\twith\ttabs', 'next']];
      const tsv = encodeTSV(grid);
      expect(parseTSV(tsv)).toEqual(grid);
    });
  });

  describe('quote-doubling edge cases', () => {
    it('CSV decodes a quoted cell whose value is `"hi"`', () => {
      // Literal `"hi"` encoded as `"""hi"""` — outer quotes + each inner `"` doubled.
      expect(parseCSV('"""hi"""')).toEqual([['"hi"']]);
    });

    it('TSV decodes a quoted cell whose value is `"hi"`', () => {
      expect(parseTSV('"""hi"""')).toEqual([['"hi"']]);
    });

    it('roundtrip preserves a cell value made of nothing but doubled quotes', () => {
      const grid = [['""""']]; // four literal `"` characters
      expect(parseCSV(encodeCSV(grid))).toEqual(grid);
      expect(parseTSV(encodeTSV(grid))).toEqual(grid);
    });
  });

  describe('cross-format paste', () => {
    // The interact layer always uses parseTSV on text/plain, but the same
    // payload can be produced by either side. These tests pin down what
    // happens when a TSV-aware paste receives CSV-shaped input and vice
    // versa — no crash, no silent corruption.
    it('parseTSV on a comma-only payload returns a single cell per row', () => {
      // No tabs → each row is one "cell" containing the whole line.
      expect(parseTSV('a,b,c\nd,e,f')).toEqual([['a,b,c'], ['d,e,f']]);
    });

    it('parseCSV on a tab-only payload returns a single cell per row', () => {
      expect(parseCSV('a\tb\tc\nd\te\tf')).toEqual([['a\tb\tc'], ['d\te\tf']]);
    });
  });

  describe('BOM handling', () => {
    it('parseCSV strips a leading BOM (Excel writes one on Windows)', () => {
      expect(parseCSV('﻿a,b')).toEqual([['a', 'b']]);
    });

    it('parseTSV does not strip BOM (intentional — leading sentinel preserved)', () => {
      // TSV parser does not strip BOM by design; the byte ends up in cell text.
      const out = parseTSV('﻿a\tb');
      expect(out[0]?.[0]?.charCodeAt(0)).toBe(0xfeff);
    });
  });

  describe('blank cells / empty grids', () => {
    it('CSV: consecutive commas yield consecutive empty cells', () => {
      expect(parseCSV('a,,b')).toEqual([['a', '', 'b']]);
    });

    it('TSV: consecutive tabs yield consecutive empty cells', () => {
      expect(parseTSV('a\t\tb')).toEqual([['a', '', 'b']]);
    });

    it('CSV: empty input returns one row with one empty cell', () => {
      expect(parseCSV('')).toEqual([['']]);
    });

    it('TSV: empty input returns one row with one empty cell', () => {
      expect(parseTSV('')).toEqual([['']]);
    });
  });

  describe('roundtrips', () => {
    it('CSV encode → CSV parse preserves a mixed grid', () => {
      const grid = [
        ['plain', 'with,comma', 'with "quote"'],
        ['', 'two\nlines', '  spaces around  '],
      ];
      expect(parseCSV(encodeCSV(grid))).toEqual(grid);
    });

    it('TSV encode → TSV parse preserves a mixed grid', () => {
      const grid = [
        ['plain', 'with\ttab', 'with "quote"'],
        ['', 'two\nlines', 'mixed\t"and"\n'],
      ];
      expect(parseTSV(encodeTSV(grid))).toEqual(grid);
    });

    it('CSV with BOM round-trips through parse and re-encode', () => {
      const grid = [['a', 'b']];
      const csv = encodeCSV(grid, { bom: true });
      expect(csv.charCodeAt(0)).toBe(0xfeff);
      expect(parseCSV(csv)).toEqual(grid);
    });
  });

  describe('large payloads (sanity, not perf)', () => {
    it('CSV: 10k rows of 5 cols parse without losing rows', () => {
      const grid = Array.from({ length: 10_000 }, (_, i) => [String(i), 'b', 'c', 'd', 'e']);
      const csv = encodeCSV(grid);
      const back = parseCSV(csv);
      expect(back.length).toBe(grid.length);
      expect(back[0]).toEqual(['0', 'b', 'c', 'd', 'e']);
      expect(back.at(-1)).toEqual(['9999', 'b', 'c', 'd', 'e']);
    });

    it('TSV: 10k rows of 5 cols parse without losing rows', () => {
      const grid = Array.from({ length: 10_000 }, (_, i) => [String(i), 'b', 'c', 'd', 'e']);
      const tsv = encodeTSV(grid);
      const back = parseTSV(tsv);
      expect(back.length).toBe(grid.length);
      expect(back.at(-1)).toEqual(['9999', 'b', 'c', 'd', 'e']);
    });
  });
});
