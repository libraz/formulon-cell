import { describe, expect, it } from 'vitest';
import { findDependents, findPrecedents } from '../../../src/engine/refs-graph.js';
import type { Addr } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

interface Cell {
  addr: { sheet: number; row: number; col: number };
  formula?: string;
}

/** Fake WorkbookHandle with the engine `precedents` / `dependents` methods
 *  stubbed to `null` — so callers fall through to the same-sheet regex path
 *  exercised by these tests. */
const wb = (cells: readonly Cell[]): WorkbookHandle =>
  ({
    precedents: () => null,
    dependents: () => null,
    cellFormula: (a: { sheet: number; row: number; col: number }) =>
      cells.find((c) => c.addr.sheet === a.sheet && c.addr.row === a.row && c.addr.col === a.col)
        ?.formula ?? null,
    cells: (sheet: number) => cells.filter((c) => c.addr.sheet === sheet),
  }) as unknown as WorkbookHandle;

/** Fake handle that pretends the engine has `traceArrows` capability and
 *  returns canned `Addr[]` slices — used to assert the engine path takes
 *  precedence over the regex fallback when available. */
const wbWithEngine = (
  precedents: ReadonlyMap<string, Addr[]>,
  dependents: ReadonlyMap<string, Addr[]>,
): WorkbookHandle =>
  ({
    precedents: (a: Addr): Addr[] | null => precedents.get(`${a.sheet}:${a.row}:${a.col}`) ?? [],
    dependents: (a: Addr): Addr[] | null => dependents.get(`${a.sheet}:${a.row}:${a.col}`) ?? [],
    // These should not be called when the engine path is engaged.
    cellFormula: () => {
      throw new Error('cellFormula called — engine path should bypass fallback');
    },
    cells: () => {
      throw new Error('cells called — engine path should bypass fallback');
    },
  }) as unknown as WorkbookHandle;

describe('findPrecedents', () => {
  it('returns empty for non-formula cells', () => {
    const handle = wb([{ addr: { sheet: 0, row: 0, col: 0 } }]);
    expect(findPrecedents(handle, { sheet: 0, row: 0, col: 0 })).toEqual([]);
  });

  it('expands single-cell refs', () => {
    const handle = wb([{ addr: { sheet: 0, row: 0, col: 2 }, formula: '=A1+B1' }]);
    expect(findPrecedents(handle, { sheet: 0, row: 0, col: 2 })).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 1 },
    ]);
  });

  it('expands range refs into individual addrs', () => {
    const handle = wb([{ addr: { sheet: 0, row: 5, col: 0 }, formula: '=SUM(A1:A3)' }]);
    expect(findPrecedents(handle, { sheet: 0, row: 5, col: 0 })).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 1, col: 0 },
      { sheet: 0, row: 2, col: 0 },
    ]);
  });

  it('skips cross-sheet refs', () => {
    const handle = wb([{ addr: { sheet: 0, row: 0, col: 0 }, formula: '=Sheet2!A1+B1' }]);
    expect(findPrecedents(handle, { sheet: 0, row: 0, col: 0 })).toEqual([
      { sheet: 0, row: 0, col: 1 },
    ]);
  });

  it('skips self-references', () => {
    const handle = wb([{ addr: { sheet: 0, row: 0, col: 0 }, formula: '=A1+B1' }]);
    expect(findPrecedents(handle, { sheet: 0, row: 0, col: 0 })).toEqual([
      { sheet: 0, row: 0, col: 1 },
    ]);
  });
});

describe('findDependents', () => {
  it('returns cells whose formulas reference the target', () => {
    const handle = wb([
      { addr: { sheet: 0, row: 0, col: 0 } },
      { addr: { sheet: 0, row: 0, col: 1 }, formula: '=A1*2' },
      { addr: { sheet: 0, row: 1, col: 1 }, formula: '=A1+1' },
    ]);
    expect(findDependents(handle, { sheet: 0, row: 0, col: 0 })).toEqual([
      { sheet: 0, row: 0, col: 1 },
      { sheet: 0, row: 1, col: 1 },
    ]);
  });

  it('catches cells whose ranges include the target', () => {
    const handle = wb([
      { addr: { sheet: 0, row: 0, col: 0 } },
      { addr: { sheet: 0, row: 5, col: 0 }, formula: '=SUM(A1:A3)' },
    ]);
    expect(findDependents(handle, { sheet: 0, row: 1, col: 0 })).toEqual([
      { sheet: 0, row: 5, col: 0 },
    ]);
  });

  it('skips cross-sheet ref dependents', () => {
    const handle = wb([
      { addr: { sheet: 0, row: 1, col: 0 }, formula: '=Sheet2!A1' },
      { addr: { sheet: 0, row: 2, col: 0 }, formula: '=A1' },
    ]);
    expect(findDependents(handle, { sheet: 0, row: 0, col: 0 })).toEqual([
      { sheet: 0, row: 2, col: 0 },
    ]);
  });
});

describe('findPrecedents (engine path)', () => {
  it('returns the engine result without invoking the regex fallback', () => {
    const handle = wbWithEngine(
      new Map([
        [
          '0:0:0',
          [
            { sheet: 0, row: 1, col: 1 },
            // Cross-sheet ref — engine surfaces it, regex path would have skipped it.
            { sheet: 1, row: 0, col: 0 },
          ],
        ],
      ]),
      new Map(),
    );
    expect(findPrecedents(handle, { sheet: 0, row: 0, col: 0 })).toEqual([
      { sheet: 0, row: 1, col: 1 },
      { sheet: 1, row: 0, col: 0 },
    ]);
  });

  it('returns an empty list when the engine reports no precedents', () => {
    const handle = wbWithEngine(new Map(), new Map());
    expect(findPrecedents(handle, { sheet: 0, row: 7, col: 7 })).toEqual([]);
  });
});

describe('findDependents (engine path)', () => {
  it('returns the engine result and includes cross-sheet dependents', () => {
    const handle = wbWithEngine(
      new Map(),
      new Map([
        [
          '0:0:0',
          [
            { sheet: 0, row: 5, col: 5 },
            { sheet: 2, row: 0, col: 0 },
          ],
        ],
      ]),
    );
    expect(findDependents(handle, { sheet: 0, row: 0, col: 0 })).toEqual([
      { sheet: 0, row: 5, col: 5 },
      { sheet: 2, row: 0, col: 0 },
    ]);
  });
});
