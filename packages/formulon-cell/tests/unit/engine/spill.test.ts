import { describe, expect, it } from 'vitest';
import {
  computeEngineSpillRanges,
  detectSpillRange,
  findSpillBlockers,
  findSpillRanges,
  looksLikeArrayFormula,
  type SpillEngineView,
} from '../../../src/engine/spill.js';
import type { Addr, CellValue } from '../../../src/engine/types.js';
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

describe('computeEngineSpillRanges', () => {
  type FormulaCell = { addr: Addr; formula: string | null };
  type SpillEntry = {
    anchorRow: number;
    anchorCol: number;
    rows: number;
    cols: number;
  };

  const makeView = (
    formulaCells: FormulaCell[],
    spillBy: Map<string, SpillEntry | null>,
  ): SpillEngineView => ({
    *cells(_sheet: number) {
      for (const c of formulaCells) yield c;
    },
    spillInfo(sheet, row, col) {
      return spillBy.get(`${sheet}:${row}:${col}`) ?? null;
    },
  });

  it('emits one rect per anchor when phantom cells share the spill region', () => {
    const cells: FormulaCell[] = [
      { addr: { sheet: 0, row: 0, col: 0 }, formula: '=SEQUENCE(3)' },
      // Phantom cells in formulon expose engaged spillInfo too, but the
      // engine surfaces them via `cells()` only when they carry a formula
      // of their own. We model that here: only the anchor is iterated.
    ];
    const spill = new Map<string, SpillEntry>([
      ['0:0:0', { anchorRow: 0, anchorCol: 0, rows: 3, cols: 1 }],
    ]);
    const ranges = computeEngineSpillRanges(makeView(cells, spill), 0);
    expect(ranges).toEqual([{ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 }]);
  });

  it('deduplicates when two formula cells report the same anchor', () => {
    const cells: FormulaCell[] = [
      { addr: { sheet: 0, row: 0, col: 0 }, formula: '=SEQUENCE(2,2)' },
      { addr: { sheet: 0, row: 0, col: 1 }, formula: '=SEQUENCE(2,2)' },
    ];
    const spill = new Map<string, SpillEntry>([
      ['0:0:0', { anchorRow: 0, anchorCol: 0, rows: 2, cols: 2 }],
      ['0:0:1', { anchorRow: 0, anchorCol: 0, rows: 2, cols: 2 }],
    ]);
    const ranges = computeEngineSpillRanges(makeView(cells, spill), 0);
    expect(ranges).toEqual([{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }]);
  });

  it('drops 1×1 spills (no outline to draw)', () => {
    const cells: FormulaCell[] = [{ addr: { sheet: 0, row: 0, col: 0 }, formula: '=UNIQUE(A1)' }];
    const spill = new Map<string, SpillEntry>([
      ['0:0:0', { anchorRow: 0, anchorCol: 0, rows: 1, cols: 1 }],
    ]);
    expect(computeEngineSpillRanges(makeView(cells, spill), 0)).toEqual([]);
  });

  it('skips formula cells the engine reports as not spilled', () => {
    const cells: FormulaCell[] = [
      { addr: { sheet: 0, row: 0, col: 0 }, formula: '=SUM(A1:A3)' },
      { addr: { sheet: 0, row: 1, col: 0 }, formula: '=SEQUENCE(2)' },
    ];
    const spill = new Map<string, SpillEntry | null>([
      ['0:0:0', null],
      ['0:1:0', { anchorRow: 1, anchorCol: 0, rows: 2, cols: 1 }],
    ]);
    const ranges = computeEngineSpillRanges(makeView(cells, spill), 0);
    expect(ranges).toEqual([{ sheet: 0, r0: 1, c0: 0, r1: 2, c1: 0 }]);
  });

  it('ignores cells without a formula', () => {
    const cells: FormulaCell[] = [{ addr: { sheet: 0, row: 0, col: 0 }, formula: null }];
    const spill = new Map<string, SpillEntry>();
    expect(computeEngineSpillRanges(makeView(cells, spill), 0)).toEqual([]);
  });
});

describe('findSpillBlockers', () => {
  const target = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 } as const;

  it('flags non-blank value cells inside the rect (anchor excluded)', () => {
    const cells = fill([
      [0, 0, { value: num(1), formula: '=SEQUENCE(3)' }],
      [1, 0, { value: num(99), formula: null }],
      [2, 0, { value: blank(), formula: null }],
    ]);
    expect(findSpillBlockers(cells, 0, target)).toEqual([{ sheet: 0, row: 1, col: 0 }]);
  });

  it('flags formula-bearing cells inside the rect', () => {
    const cells = fill([
      [0, 0, { value: num(1), formula: '=SEQUENCE(3)' }],
      [2, 0, { value: blank(), formula: '=10' }],
    ]);
    expect(findSpillBlockers(cells, 0, target)).toEqual([{ sheet: 0, row: 2, col: 0 }]);
  });

  it('returns [] when the rect is fully clear apart from the anchor', () => {
    const cells = fill([[0, 0, { value: num(1), formula: '=SEQUENCE(3)' }]]);
    expect(findSpillBlockers(cells, 0, target)).toEqual([]);
  });

  it('skips cells on a different sheet', () => {
    const cells = fill(
      [
        [0, 0, { value: num(1), formula: '=SEQUENCE(3)' }],
        [1, 0, { value: num(99), formula: null }],
      ],
      1,
    );
    // Querying sheet 0 — populated entries are on sheet 1.
    expect(findSpillBlockers(cells, 0, target)).toEqual([]);
  });

  it('emits blockers in row-major order', () => {
    const cells = fill([
      [0, 0, { value: num(1), formula: '=SEQUENCE(3,2)' }],
      [0, 1, { value: num(2), formula: null }],
      [1, 0, { value: num(3), formula: null }],
      [1, 1, { value: num(4), formula: null }],
    ]);
    const wide = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 } as const;
    const blockers = findSpillBlockers(cells, 0, wide);
    // Anchor (0,0) excluded; remaining three reported left-to-right, top-to-bottom.
    expect(blockers).toEqual([
      { sheet: 0, row: 0, col: 1 },
      { sheet: 0, row: 1, col: 0 },
      { sheet: 0, row: 1, col: 1 },
    ]);
  });
});
