import { beforeEach, describe, expect, it, vi } from 'vitest';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { inferSortHasHeader, removeDuplicates, sortRange } from '../../../src/commands/sort.js';
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

  it('sorts by multiple keys in order', () => {
    seedText(store, wb, 0, 0, 'Region');
    seedText(store, wb, 0, 1, 'Item');
    seedText(store, wb, 1, 0, 'West');
    seedText(store, wb, 1, 1, 'Paper');
    seedText(store, wb, 2, 0, 'East');
    seedText(store, wb, 2, 1, 'Ink');
    seedText(store, wb, 3, 0, 'East');
    seedText(store, wb, 3, 1, 'Paper');
    seedText(store, wb, 4, 0, 'West');
    seedText(store, wb, 4, 1, 'Ink');

    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 },
      {
        byCol: 0,
        direction: 'asc',
        hasHeader: true,
        keys: [
          { byCol: 0, direction: 'asc' },
          { byCol: 1, direction: 'desc' },
        ],
      },
    );

    expect(ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'East' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'Paper' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'East' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'text', value: 'Ink' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'text', value: 'West' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 1 })).toEqual({ kind: 'text', value: 'Paper' });
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toEqual({ kind: 'text', value: 'West' });
    expect(wb.getValue({ sheet: 0, row: 4, col: 1 })).toEqual({ kind: 'text', value: 'Ink' });
  });

  it('infers headers for label rows but not for plain numeric ranges', () => {
    seedText(store, wb, 0, 0, 'item');
    seedText(store, wb, 0, 1, 'qty');
    seedText(store, wb, 1, 0, 'paper');
    seedNumber(store, wb, 1, 1, 30);
    seedText(store, wb, 2, 0, 'ink');
    seedNumber(store, wb, 2, 1, 10);

    expect(inferSortHasHeader(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 })).toBe(
      true,
    );

    const numericStore = createSpreadsheetStore();
    const numericWb = wb;
    seedNumber(numericStore, numericWb, 0, 0, 30);
    seedNumber(numericStore, numericWb, 1, 0, 10);
    expect(
      inferSortHasHeader(numericStore.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }),
    ).toBe(false);
  });

  it('infers a formatted text header when the data below is also text', () => {
    seedText(store, wb, 0, 0, 'name');
    seedText(store, wb, 1, 0, 'paper');
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });

    expect(inferSortHasHeader(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 })).toBe(
      true,
    );
  });

  it('infers a header for an all-text table even without format contrast (H-18)', () => {
    // Region / Product columns, no bold header, no numeric column — the old
    // heuristic dropped the label row into the sort.
    seedText(store, wb, 0, 0, 'Region');
    seedText(store, wb, 0, 1, 'Product');
    seedText(store, wb, 1, 0, 'West');
    seedText(store, wb, 1, 1, 'Paper');
    seedText(store, wb, 2, 0, 'East');
    seedText(store, wb, 2, 1, 'Ink');

    expect(inferSortHasHeader(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 })).toBe(
      true,
    );
  });

  it('moves cell formatting with sorted rows while leaving the header format in place', () => {
    seedText(store, wb, 0, 0, 'item');
    seedText(store, wb, 1, 0, 'banana');
    seedText(store, wb, 2, 0, 'apple');
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fill: '#fff2cc' });
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 0 }, { fill: '#c6efce' });

    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      { byCol: 0, direction: 'asc', hasHeader: true },
    );

    expect(ok).toBe(true);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      bold: true,
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))).toEqual({
      fill: '#c6efce',
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 0 }))).toEqual({
      fill: '#fff2cc',
    });
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

  it('refuses to sort a locked range on a protected sheet', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    seedNumber(store, wb, 0, 0, 30);
    seedNumber(store, wb, 1, 0, 10);
    setProtectedSheet(store, 0, true);

    try {
      const ok = sortRange(
        store.getState(),
        store,
        wb,
        { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
        { byCol: 0, direction: 'asc' },
      );

      expect(ok).toBe(false);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 30 });
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 10 });
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });

  it('sorts protected sheets when every affected cell is explicitly unlocked', () => {
    seedNumber(store, wb, 0, 0, 30);
    seedNumber(store, wb, 1, 0, 10);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }, false);
    setProtectedSheet(store, 0, true);

    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      { byCol: 0, direction: 'asc' },
    );

    expect(ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 10 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 30 });
  });
});

describe('removeDuplicates', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('removes duplicate rows and blanks empty cells moved over old content', () => {
    seedText(store, wb, 0, 0, 'alpha');
    seedNumber(store, wb, 0, 1, 1);
    seedText(store, wb, 1, 0, 'alpha');
    seedNumber(store, wb, 1, 1, 1);
    seedText(store, wb, 2, 0, 'beta');

    const removed = removeDuplicates(store.getState(), store, wb, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 2,
      c1: 1,
    });

    expect(removed).toBe(1);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'alpha' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'beta' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'blank' });
  });

  it('moves formatting with kept rows and clears the duplicate tail formats', () => {
    seedText(store, wb, 0, 0, 'alpha');
    seedText(store, wb, 1, 0, 'alpha');
    seedText(store, wb, 2, 0, 'beta');
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fill: '#fff2cc' });
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fill: '#f4cccc' });
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 0 }, { fill: '#c6efce' });

    const removed = removeDuplicates(store.getState(), store, wb, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 2,
      c1: 0,
    });

    expect(removed).toBe(1);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      fill: '#fff2cc',
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))).toEqual({
      fill: '#c6efce',
    });
    expect(store.getState().format.formats.has(addrKey({ sheet: 0, row: 2, col: 0 }))).toBe(false);
  });

  it('compares only the selected columns when removing duplicates', () => {
    seedText(store, wb, 0, 0, 'alpha');
    seedNumber(store, wb, 0, 1, 1);
    seedText(store, wb, 1, 0, 'alpha');
    seedNumber(store, wb, 1, 1, 2);
    seedText(store, wb, 2, 0, 'beta');
    seedNumber(store, wb, 2, 1, 1);

    const removed = removeDuplicates(
      store.getState(),
      store,
      wb,
      {
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 2,
        c1: 1,
      },
      { columns: [0] },
    );

    expect(removed).toBe(1);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'alpha' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'beta' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'blank' });
  });

  it('preserves the header row when requested', () => {
    seedText(store, wb, 0, 0, 'item');
    seedText(store, wb, 0, 1, 'qty');
    seedText(store, wb, 1, 0, 'paper');
    seedNumber(store, wb, 1, 1, 1);
    seedText(store, wb, 2, 0, 'paper');
    seedNumber(store, wb, 2, 1, 1);

    const removed = removeDuplicates(
      store.getState(),
      store,
      wb,
      {
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 2,
        c1: 1,
      },
      { hasHeader: true },
    );

    expect(removed).toBe(1);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'item' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'qty' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'paper' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'blank' });
  });

  it('refuses to remove duplicates from a locked range on a protected sheet', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    seedText(store, wb, 0, 0, 'alpha');
    seedText(store, wb, 1, 0, 'alpha');
    setProtectedSheet(store, 0, true);

    try {
      const removed = removeDuplicates(store.getState(), store, wb, {
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 1,
        c1: 0,
      });

      expect(removed).toBe(0);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'alpha' });
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'alpha' });
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });
});
