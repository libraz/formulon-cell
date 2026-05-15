import { beforeEach, describe, expect, it } from 'vitest';
import { sortRange } from '../../../src/commands/sort.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedNumber = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: number,
): void => {
  wb.setNumber({ sheet: 0, row, col }, value);
  store.setState((s) => {
    const map = new Map(s.data.cells);
    map.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells: map } };
  });
};

const seedText = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: string,
): void => {
  wb.setText({ sheet: 0, row, col }, value);
  store.setState((s) => {
    const map = new Map(s.data.cells);
    map.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'text', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells: map } };
  });
};

describe('sortRange', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('ascending sort by a single numeric column reorders rows in place', () => {
    seedNumber(store, wb, 0, 0, 30);
    seedNumber(store, wb, 1, 0, 10);
    seedNumber(store, wb, 2, 0, 20);

    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      { byCol: 0, direction: 'asc' },
    );
    expect(ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 10 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 20 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'number', value: 30 });
  });

  it('descending sort by a numeric column reverses the order', () => {
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 1, 0, 2);
    seedNumber(store, wb, 2, 0, 3);

    sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      { byCol: 0, direction: 'desc' },
    );
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 2 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'number', value: 1 });
  });

  it('hasHeader excludes row 0 from the move and keeps the header in place', () => {
    seedText(store, wb, 0, 0, 'header');
    seedText(store, wb, 0, 1, 'qty');
    seedText(store, wb, 1, 0, 'banana');
    seedNumber(store, wb, 1, 1, 7);
    seedText(store, wb, 2, 0, 'apple');
    seedNumber(store, wb, 2, 1, 12);

    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      { byCol: 0, direction: 'asc', hasHeader: true },
    );
    expect(ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'header' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'apple' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'banana' });
    // The companion column moves with its row.
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 12 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'number', value: 7 });
  });

  it('mixed numeric + text rows place numbers before text under ascending sort', () => {
    seedText(store, wb, 0, 0, 'banana');
    seedNumber(store, wb, 1, 0, 99);
    seedText(store, wb, 2, 0, 'apple');
    seedNumber(store, wb, 3, 0, 1);

    sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
      { byCol: 0, direction: 'asc' },
    );
    // Numbers grouped at the top in ascending order, text after.
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 99 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'apple' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'text', value: 'banana' });
  });

  it('header-only range (single row, hasHeader: true) is a no-op (returns false)', () => {
    seedText(store, wb, 0, 0, 'just-header');
    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      { byCol: 0, direction: 'asc', hasHeader: true },
    );
    expect(ok).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'just-header',
    });
  });

  it('refuses to sort when the byCol is outside the range', () => {
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 1, 0, 2);
    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      { byCol: 5, direction: 'asc' },
    );
    expect(ok).toBe(false);
  });

  it('refuses to sort when the range overlaps a merged cell', () => {
    seedNumber(store, wb, 0, 0, 30);
    seedNumber(store, wb, 1, 0, 10);
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      { byCol: 0, direction: 'asc' },
    );
    expect(ok).toBe(false);
    // Original ordering preserved.
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 30 });
  });

  it('integrates with the workbook engine: recalc preserves sorted ordering', () => {
    seedNumber(store, wb, 0, 0, 5);
    seedNumber(store, wb, 1, 0, 3);
    seedNumber(store, wb, 2, 0, 8);

    sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      { byCol: 0, direction: 'asc' },
    );
    // Recalc again — sort already calls recalc internally; double-checking
    // that the engine's stored values still match the sorted layout.
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 5 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'number', value: 8 });
  });
});
