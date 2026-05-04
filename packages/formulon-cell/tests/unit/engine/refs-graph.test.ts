import { describe, expect, it } from 'vitest';
import { findDependents, findPrecedents } from '../../../src/engine/refs-graph.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

interface Cell {
  addr: { sheet: number; row: number; col: number };
  formula?: string;
}

const wb = (cells: readonly Cell[]): WorkbookHandle =>
  ({
    cellFormula: (a: { sheet: number; row: number; col: number }) =>
      cells.find((c) => c.addr.sheet === a.sheet && c.addr.row === a.row && c.addr.col === a.col)
        ?.formula ?? null,
    cells: (sheet: number) => cells.filter((c) => c.addr.sheet === sheet),
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
