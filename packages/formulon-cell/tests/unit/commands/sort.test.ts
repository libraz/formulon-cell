import { beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { inferSortHasHeader, removeDuplicates, sortRange } from '../../../src/commands/sort.js';
import type { CellValue } from '../../../src/engine/types.js';
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

const seedCell = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: CellValue,
  formula: string | null = null,
): void => {
  const addr = { sheet: 0, row, col };
  if (formula) {
    wb.setFormula(addr, formula);
  } else {
    switch (value.kind) {
      case 'number':
        wb.setNumber(addr, value.value);
        break;
      case 'text':
        wb.setText(addr, value.value);
        break;
      case 'bool':
        wb.setBool(addr, value.value);
        break;
      case 'error':
        wb.setError(addr, value.code);
        break;
      default:
        wb.setBlank(addr);
    }
  }
  store.setState((s) => {
    const map = new Map(s.data.cells);
    if (value.kind === 'blank' && !formula) map.delete(addrKey(addr));
    else map.set(addrKey(addr), { value, formula });
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

  it('sorts by cell fill color while moving whole rows', () => {
    seedText(store, wb, 0, 0, 'Item');
    seedText(store, wb, 0, 1, 'Status');
    seedText(store, wb, 1, 0, 'blue row');
    seedText(store, wb, 1, 1, 'later');
    seedText(store, wb, 2, 0, 'red row');
    seedText(store, wb, 2, 1, 'first');
    seedText(store, wb, 3, 0, 'plain row');
    seedText(store, wb, 3, 1, 'last');
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fill: '#0000ff' });
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 0 }, { fill: '#ff0000' });

    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
      {
        byCol: 0,
        direction: 'asc',
        hasHeader: true,
        keys: [{ byCol: 0, direction: 'asc', sortOn: 'cellColor', color: '#FF0000' }],
      },
    );

    expect(ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'red row',
    });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'first',
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))).toEqual({
      fill: '#ff0000',
    });
  });

  it('sorts by built-in custom list order', () => {
    seedText(store, wb, 0, 0, 'Month');
    seedText(store, wb, 1, 0, 'Mar');
    seedText(store, wb, 2, 0, 'Jan');
    seedText(store, wb, 3, 0, 'Feb');

    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
      {
        byCol: 0,
        direction: 'asc',
        hasHeader: true,
        keys: [{ byCol: 0, direction: 'asc', sortOn: 'customList' }],
      },
    );

    expect(ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'Jan' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'Feb' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'text', value: 'Mar' });
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

  it('orders typed values as number, text, bool, error, then blank', () => {
    seedCell(store, wb, 0, 0, { kind: 'bool', value: true });
    seedText(store, wb, 1, 0, 'Item2');
    seedNumber(store, wb, 2, 0, 9);
    seedCell(store, wb, 3, 0, { kind: 'bool', value: false });
    wb.setFormula({ sheet: 0, row: 4, col: 0 }, '=NA()');
    wb.recalc();
    const errorValue = wb.getValue({ sheet: 0, row: 4, col: 0 });
    expect(errorValue.kind).toBe('error');
    seedCell(store, wb, 4, 0, errorValue, '=NA()');
    seedText(store, wb, 5, 0, 'Item10');

    sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 6, c1: 0 },
      { byCol: 0, direction: 'asc' },
    );

    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 9 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'Item10' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'Item2' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'bool', value: false });
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toEqual({ kind: 'bool', value: true });
    expect(wb.getValue({ sheet: 0, row: 5, col: 0 })).toMatchObject({ kind: 'error' });
    expect(wb.getValue({ sheet: 0, row: 6, col: 0 })).toEqual({ kind: 'blank' });
  });

  it('moves formulas by row and keeps outside references unchanged in one undo step', () => {
    const gValues = [3, 1, 2];
    const jValues = [100, 200, 300];
    for (let row = 0; row < gValues.length; row += 1) {
      seedNumber(store, wb, row, 6, gValues[row] ?? 0);
      seedNumber(store, wb, row, 9, jValues[row] ?? 0);
      wb.setFormula({ sheet: 0, row, col: 7 }, `=G${row + 1}+$J$1+J${row + 1}`);
      wb.recalc();
      seedCell(
        store,
        wb,
        row,
        7,
        wb.getValue({ sheet: 0, row, col: 7 }),
        `=G${row + 1}+$J$1+J${row + 1}`,
      );
    }
    wb.setFormula({ sheet: 0, row: 0, col: 10 }, '=G1');
    wb.recalc();
    seedCell(store, wb, 0, 10, wb.getValue({ sheet: 0, row: 0, col: 10 }), '=G1');
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 7 },
      {
        bold: true,
        comment: 'moved note',
        commentAuthor: 'Alice',
      },
    );

    const history = new History();
    wb.attachHistory(history);
    history.begin();
    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 6, r1: 2, c1: 7 },
      { byCol: 6, direction: 'asc' },
      history,
    );
    history.end();

    expect(ok).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 7 })).toBe('=G1+$J$1+J1');
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 7 })).toBe('=G2+$J$1+J2');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 7 })).toBe('=G3+$J$1+J3');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 10 })).toBe('=G1');
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 7 })),
    ).toMatchObject({
      bold: true,
      comment: 'moved note',
      commentAuthor: 'Alice',
    });

    expect(history.undo()).toBe(true);
    expect(history.canUndo()).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 0, col: 6 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 7 })).toBe('=G1+$J$1+J1');
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 7 })),
    ).toMatchObject({
      bold: true,
      comment: 'moved note',
      commentAuthor: 'Alice',
    });

    expect(history.redo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 6 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 7 })).toBe('=G1+$J$1+J1');
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 7 })),
    ).toMatchObject({
      bold: true,
      comment: 'moved note',
      commentAuthor: 'Alice',
    });
  });

  it('sorts text case-insensitively without numeric collation', () => {
    seedText(store, wb, 0, 0, 'item10');
    seedText(store, wb, 1, 0, 'Item2');
    seedText(store, wb, 2, 0, 'item1');

    sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      { byCol: 0, direction: 'asc' },
    );

    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'item1' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'item10' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'Item2' });
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

  it('refuses huge sort ranges before scanning every cell', () => {
    seedNumber(store, wb, 0, 0, 30);
    seedNumber(store, wb, 1, 0, 10);

    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 },
      { byCol: 0, direction: 'asc' },
    );

    expect(ok).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 30 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 10 });
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

  it('preserves static errors when kept rows are compacted', () => {
    seedText(store, wb, 0, 0, 'alpha');
    seedCell(store, wb, 0, 1, { kind: 'error', code: 7, text: '#ERR!' });
    seedText(store, wb, 1, 0, 'alpha');
    seedNumber(store, wb, 1, 1, 1);
    seedText(store, wb, 2, 0, 'beta');
    seedCell(store, wb, 2, 1, { kind: 'error', code: 15, text: '#VALUE!' });
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 1 }, { bold: true });

    const removed = removeDuplicates(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      { columns: [0] },
    );

    expect(removed).toBe(1);
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toMatchObject({
      kind: 'error',
      code: 15,
    });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'blank' });
    expect(store.getState().format.formats.has(addrKey({ sheet: 0, row: 2, col: 1 }))).toBe(false);
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

  it('refuses huge duplicate-removal ranges before rewriting the tail', () => {
    seedText(store, wb, 0, 0, 'alpha');
    seedText(store, wb, 1, 0, 'alpha');

    const removed = removeDuplicates(store.getState(), store, wb, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1048575,
      c1: 0,
    });

    expect(removed).toBe(0);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'alpha' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'alpha' });
  });
});
