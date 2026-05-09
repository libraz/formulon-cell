import { describe, expect, it } from 'vitest';
import {
  clearSparkline,
  clearSparklinesInRange,
  listSparklines,
  setSparkline,
  sparklineAt,
} from '../../../src/commands/sparkline.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('sparkline commands', () => {
  it('sets, reads, and lists sparklines in row-major order', () => {
    const store = createSpreadsheetStore();
    setSparkline(store, { sheet: 0, row: 2, col: 1 }, { kind: 'line', source: 'A1:A4' });
    setSparkline(store, { sheet: 0, row: 1, col: 1 }, { kind: 'column', source: 'B1:B4' });

    expect(sparklineAt(store.getState(), { sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'line',
      source: 'A1:A4',
    });
    expect(listSparklines(store.getState())).toEqual([
      { addr: { sheet: 0, row: 1, col: 1 }, spec: { kind: 'column', source: 'B1:B4' } },
      { addr: { sheet: 0, row: 2, col: 1 }, spec: { kind: 'line', source: 'A1:A4' } },
    ]);
  });

  it('clears a single sparkline or all sparklines in a range', () => {
    const store = createSpreadsheetStore();
    setSparkline(store, { sheet: 0, row: 0, col: 0 }, { kind: 'line', source: 'A1:A4' });
    setSparkline(store, { sheet: 0, row: 2, col: 0 }, { kind: 'column', source: 'B1:B4' });

    clearSparkline(store, { sheet: 0, row: 0, col: 0 });
    expect(sparklineAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();

    clearSparklinesInRange(store, { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 });
    expect(listSparklines(store.getState())).toEqual([]);
  });
});
