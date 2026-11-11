import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
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
    expect(
      setSparkline(store, { sheet: 0, row: 2, col: 1 }, { kind: 'line', source: 'A1:A4' }),
    ).toBe(true);
    expect(
      setSparkline(store, { sheet: 0, row: 1, col: 1 }, { kind: 'column', source: 'B1:B4' }),
    ).toBe(true);

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

    expect(clearSparkline(store, { sheet: 0, row: 0, col: 0 })).toBe(true);
    expect(sparklineAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();

    expect(clearSparklinesInRange(store, { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 })).toBe(1);
    expect(listSparklines(store.getState())).toEqual([]);
  });

  it('refuses inserting a sparkline under any protection, Edit objects included', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 0, col: 0 };
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    for (const permissions of [{}, { objects: true }]) {
      setProtectedSheet(store, 0, true, { permissions });
      expect(setSparkline(store, addr, { kind: 'line', source: 'A1:A4' })).toBe(false);
      expect(sparklineAt(store.getState(), addr)).toBeNull();
    }
  });

  it('clears sparklines by the lock of their cells, whatever Edit objects says', () => {
    const store = createSpreadsheetStore();
    setSparkline(store, { sheet: 0, row: 0, col: 0 }, { kind: 'line', source: 'A1:A4' });
    setSparkline(store, { sheet: 0, row: 0, col: 1 }, { kind: 'line', source: 'B1:B4' });
    setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
    setProtectedSheet(store, 0, true);

    expect(clearSparkline(store, { sheet: 0, row: 0, col: 0 })).toBe(false);
    expect(clearSparklinesInRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 })).toBe(1);
    expect(sparklineAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'line',
      source: 'A1:A4',
    });
    expect(sparklineAt(store.getState(), { sheet: 0, row: 0, col: 1 })).toBeNull();
  });

  it('records sparkline set and clear operations in history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const addr = { sheet: 0, row: 2, col: 1 };

    setSparkline(store, addr, { kind: 'column', source: 'A1:A4' }, history);
    expect(sparklineAt(store.getState(), addr)).toEqual({ kind: 'column', source: 'A1:A4' });
    expect(history.undo()).toBe(true);
    expect(sparklineAt(store.getState(), addr)).toBeNull();
    expect(history.redo()).toBe(true);
    expect(sparklineAt(store.getState(), addr)).toEqual({ kind: 'column', source: 'A1:A4' });

    clearSparkline(store, addr, history);
    expect(sparklineAt(store.getState(), addr)).toBeNull();
    expect(history.undo()).toBe(true);
    expect(sparklineAt(store.getState(), addr)).toEqual({ kind: 'column', source: 'A1:A4' });
  });
});
