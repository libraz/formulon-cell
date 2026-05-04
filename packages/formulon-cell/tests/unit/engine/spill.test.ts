import { describe, expect, it } from 'vitest';
import {
  detectSpillRange,
  findSpillRanges,
  looksLikeArrayFormula,
} from '../../../src/engine/spill.js';
import type { CellValue } from '../../../src/engine/types.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';

const num = (n: number): CellValue => ({ kind: 'number', value: n });
const blank = (): CellValue => ({ kind: 'blank' });

interface FakeCell {
  value: CellValue;
  formula: string | null;
}

const fill = (entries: [number, number, FakeCell][], sheet = 0): Map<string, FakeCell> => {
  const m = new Map<string, FakeCell>();
  for (const [row, col, cell] of entries) {
    m.set(addrKey({ sheet, row, col }), cell);
  }
  return m;
};

describe('looksLikeArrayFormula', () => {
  it('flags bare ranges', () => {
    expect(looksLikeArrayFormula('=A1:A10')).toBe(true);
    expect(looksLikeArrayFormula('=Sheet1!A1:B5')).toBe(true);
  });

  it('flags known dynamic-array functions', () => {
    expect(looksLikeArrayFormula('=FILTER(A:A, B:B>0)')).toBe(true);
    expect(looksLikeArrayFormula('=UNIQUE(A1:A10)')).toBe(true);
    expect(looksLikeArrayFormula('=SEQUENCE(5)')).toBe(true);
    expect(looksLikeArrayFormula('=filter(a:a)')).toBe(true);
  });

  it('rejects non-array formulas', () => {
    expect(looksLikeArrayFormula('=SUM(A1:A10)')).toBe(false);
    expect(looksLikeArrayFormula('=A1+B2')).toBe(false);
    expect(looksLikeArrayFormula('=IF(A1, B1, C1)')).toBe(false);
  });

  it('rejects non-formula text', () => {
    expect(looksLikeArrayFormula('FILTER(A1)')).toBe(false);
    expect(looksLikeArrayFormula('')).toBe(false);
  });
});

describe('detectSpillRange', () => {
  it('walks down a column from anchor', () => {
    const cells = fill([
      [0, 0, { value: num(1), formula: '=SEQUENCE(3)' }],
      [1, 0, { value: num(2), formula: null }],
      [2, 0, { value: num(3), formula: null }],
    ]);
    expect(detectSpillRange(cells, 0, 0, 0)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
  });

  it('walks right across a row', () => {
    const cells = fill([
      [0, 0, { value: num(1), formula: '=A1:C1' }],
      [0, 1, { value: num(2), formula: null }],
      [0, 2, { value: num(3), formula: null }],
    ]);
    expect(detectSpillRange(cells, 0, 0, 0)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
  });

  it('stops at a blank or formula-bearing cell', () => {
    const cells = fill([
      [0, 0, { value: num(1), formula: '=SEQUENCE(5)' }],
      [1, 0, { value: num(2), formula: null }],
      [2, 0, { value: blank(), formula: null }],
      [3, 0, { value: num(4), formula: null }],
    ]);
    expect(detectSpillRange(cells, 0, 0, 0)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
  });
});

describe('findSpillRanges', () => {
  it('returns multi-cell spills only', () => {
    const cells = fill([
      [0, 0, { value: num(1), formula: '=SEQUENCE(3)' }],
      [1, 0, { value: num(2), formula: null }],
      [2, 0, { value: num(3), formula: null }],
      [5, 5, { value: num(7), formula: '=SUM(A1:A3)' }], // not array
    ]);
    const ranges = findSpillRanges(cells, 0);
    expect(ranges).toHaveLength(1);
    expect(ranges[0]).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
  });

  it('skips anchors that resolve to a 1×1 (no spill)', () => {
    const cells = fill([[0, 0, { value: num(1), formula: '=UNIQUE(A1)' }]]);
    expect(findSpillRanges(cells, 0)).toEqual([]);
  });

  it('isolates per-sheet anchors', () => {
    const cells = fill([
      [0, 0, { value: num(1), formula: '=SEQUENCE(2)' }],
      [1, 0, { value: num(2), formula: null }],
    ]);
    expect(findSpillRanges(cells, 1)).toEqual([]);
  });
});
