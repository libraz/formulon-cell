import { describe, expect, it } from 'vitest';
import {
  addTraceArrow,
  clearTraceArrows,
  traceDependents,
  tracePrecedents,
} from '../../../src/commands/traces.js';
import type { Addr } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const key = (addr: Addr) => `${addr.sheet}:${addr.row}:${addr.col}`;

const wbWithEngine = (
  precedents: ReadonlyMap<string, Addr[]>,
  dependents: ReadonlyMap<string, Addr[]>,
): WorkbookHandle =>
  ({
    precedents: (addr: Addr): Addr[] | null => precedents.get(key(addr)) ?? [],
    dependents: (addr: Addr): Addr[] | null => dependents.get(key(addr)) ?? [],
  }) as unknown as WorkbookHandle;

describe('trace commands', () => {
  it('adds trace arrows without duplicating identical endpoints', () => {
    const store = createSpreadsheetStore();
    const arrow = {
      kind: 'precedent' as const,
      from: { sheet: 0, row: 0, col: 0 },
      to: { sheet: 0, row: 1, col: 1 },
    };

    addTraceArrow(store, arrow);
    addTraceArrow(store, arrow);

    expect(store.getState().traces.items).toEqual([arrow]);
  });

  it('traces precedents and dependents for an explicit address', () => {
    const store = createSpreadsheetStore();
    const a1 = { sheet: 0, row: 0, col: 0 };
    const b1 = { sheet: 0, row: 0, col: 1 };
    const c1 = { sheet: 0, row: 0, col: 2 };
    const workbook = wbWithEngine(new Map([[key(c1), [a1, b1]]]), new Map([[key(a1), [c1]]]));

    expect(tracePrecedents(store, workbook, c1)).toBe(2);
    expect(traceDependents(store, workbook, a1)).toBe(1);

    expect(store.getState().traces.items).toEqual([
      { kind: 'precedent', from: a1, to: c1 },
      { kind: 'precedent', from: b1, to: c1 },
      { kind: 'dependent', from: a1, to: c1 },
    ]);
  });

  it('defaults to the active cell and clears all trace arrows', () => {
    const store = createSpreadsheetStore();
    const a1 = { sheet: 0, row: 0, col: 0 };
    const b1 = { sheet: 0, row: 0, col: 1 };
    const workbook = wbWithEngine(new Map([[key(b1), [a1]]]), new Map());
    mutators.setActive(store, b1);

    expect(tracePrecedents(store, workbook)).toBe(1);
    clearTraceArrows(store);

    expect(store.getState().traces.items).toEqual([]);
  });
});
