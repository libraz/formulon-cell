import { SparklineType } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import {
  clearSparkline,
  clearSparklinesInRange,
  defaultSparklineGroupSettings,
  groupOfCell,
  listSparklines,
  setSparkline,
  sparklineAt,
} from '../../../src/commands/sparkline.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const LINE = defaultSparklineGroupSettings(SparklineType.Line);
const COLUMN = defaultSparklineGroupSettings(SparklineType.Column);

describe('sparkline commands', () => {
  it('sets, reads, and lists sparklines in row-major order', () => {
    const store = createSpreadsheetStore();
    expect(setSparkline(store, { sheet: 0, row: 2, col: 1 }, 'A1:A4', LINE)).toBe(true);
    expect(setSparkline(store, { sheet: 0, row: 1, col: 1 }, 'B1:B4', COLUMN)).toBe(true);

    expect(sparklineAt(store.getState(), { sheet: 0, row: 2, col: 1 })).toMatchObject({
      source: 'A1:A4',
    });
    expect(groupOfCell(store.getState(), { sheet: 0, row: 2, col: 1 })?.settings.type).toBe(
      SparklineType.Line,
    );
    expect(
      listSparklines(store.getState()).map(({ addr, spec }) => [addr.row, spec.source]),
    ).toEqual([
      [1, 'B1:B4'],
      [2, 'A1:A4'],
    ]);
  });

  it('clears a single sparkline or all sparklines in a range', () => {
    const store = createSpreadsheetStore();
    setSparkline(store, { sheet: 0, row: 0, col: 0 }, 'A1:A4', LINE);
    setSparkline(store, { sheet: 0, row: 2, col: 0 }, 'B1:B4', COLUMN);

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
      expect(setSparkline(store, addr, 'A1:A4', LINE)).toBe(false);
      expect(sparklineAt(store.getState(), addr)).toBeNull();
    }
  });

  it('clears sparklines by the lock of their cells, whatever Edit objects says', () => {
    const store = createSpreadsheetStore();
    setSparkline(store, { sheet: 0, row: 0, col: 0 }, 'A1:A4', LINE);
    setSparkline(store, { sheet: 0, row: 0, col: 1 }, 'B1:B4', LINE);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
    setProtectedSheet(store, 0, true);

    expect(clearSparkline(store, { sheet: 0, row: 0, col: 0 })).toBe(false);
    expect(clearSparklinesInRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 })).toBe(1);
    expect(sparklineAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toMatchObject({
      source: 'A1:A4',
    });
    expect(sparklineAt(store.getState(), { sheet: 0, row: 0, col: 1 })).toBeNull();
  });

  it('records sparkline set and clear operations in history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const addr = { sheet: 0, row: 2, col: 1 };

    setSparkline(store, addr, 'A1:A4', COLUMN, history);
    expect(sparklineAt(store.getState(), addr)).toMatchObject({ source: 'A1:A4' });
    expect(history.undo()).toBe(true);
    expect(sparklineAt(store.getState(), addr)).toBeNull();
    expect(history.redo()).toBe(true);
    expect(sparklineAt(store.getState(), addr)).toMatchObject({ source: 'A1:A4' });

    clearSparkline(store, addr, history);
    expect(sparklineAt(store.getState(), addr)).toBeNull();
    expect(history.undo()).toBe(true);
    expect(sparklineAt(store.getState(), addr)).toMatchObject({ source: 'A1:A4' });
  });
});
