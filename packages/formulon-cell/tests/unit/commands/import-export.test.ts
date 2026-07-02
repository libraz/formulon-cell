import { beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { exportCSV, importCSV } from '../../../src/commands/import-export.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string | boolean }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else if (typeof c.value === 'boolean') {
        wb.setBool(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'bool', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

const mirrorEngine = (store: SpreadsheetStore, wb: WorkbookHandle): void => {
  mutators.replaceCells(store, wb.cells(0));
};

describe('importCSV', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('writes a 2x3 CSV grid starting at the active cell', () => {
    mutators.setActive(store, { sheet: 0, row: 2, col: 1 });
    const result = importCSV(store.getState(), wb, 'a,b,c\n1,2,3');
    expect(result).not.toBeNull();
    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 2, c0: 1, r1: 3, c1: 3 });
    expect(result?.cellsWritten).toBe(6);
    expect(result?.rows).toBe(2);

    mirrorEngine(store, wb);
    const cells = store.getState().data.cells;
    expect(cells.get(addrKey({ sheet: 0, row: 2, col: 1 }))?.value).toEqual({
      kind: 'text',
      value: 'a',
    });
    expect(cells.get(addrKey({ sheet: 0, row: 3, col: 3 }))?.value).toEqual({
      kind: 'number',
      value: 3,
    });
  });

  it('coerces numeric strings to numbers and = strings to formulas', () => {
    importCSV(store.getState(), wb, '=1+2,42,hello');
    mirrorEngine(store, wb);
    const cells = store.getState().data.cells;
    // 0:0 is a formula cell whose evaluated value is 3.
    expect(cells.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.formula).toBe('=1+2');
    expect(cells.get(addrKey({ sheet: 0, row: 0, col: 1 }))?.value).toEqual({
      kind: 'number',
      value: 42,
    });
    expect(cells.get(addrKey({ sheet: 0, row: 0, col: 2 }))?.value).toEqual({
      kind: 'text',
      value: 'hello',
    });
  });

  it('respects Text-formatted destination cells when importing CSV', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { numFmt: { kind: 'text' } });
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { numFmt: { kind: 'text' } });

    const result = importCSV(store.getState(), wb, '00123,=A1');

    expect(result?.cellsWritten).toBe(2);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: '00123' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: '=A1' });
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBeNull();
  });

  it('uses the explicit anchor when provided', () => {
    const result = importCSV(store.getState(), wb, 'x', { sheet: 0, row: 9, col: 5 });
    expect(result?.writtenRange.r0).toBe(9);
    expect(result?.writtenRange.c0).toBe(5);
  });

  it('returns null on empty input', () => {
    expect(importCSV(store.getState(), wb, '')).toBeNull();
  });

  it('handles ragged rows (cellsWritten reflects per-row width)', () => {
    importCSV(store.getState(), wb, 'a,b,c\n1\n9,8,7,6');
    mirrorEngine(store, wb);
    const cells = store.getState().data.cells;
    // (1, 1) was never written by the CSV (row 2 is "1" only) — stays blank.
    expect(cells.get(addrKey({ sheet: 0, row: 1, col: 1 }))).toBeUndefined();
    // (2, 3) was written by the third row.
    expect(cells.get(addrKey({ sheet: 0, row: 2, col: 3 }))?.value).toEqual({
      kind: 'number',
      value: 6,
    });
  });

  it('skips locked protected destinations while importing into unlocked cells', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    seedAndMirror(store, wb, [{ row: 0, col: 1, value: 'locked' }]);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    setCellLocked(store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 1 }, false);
    setProtectedSheet(store, 0, true);

    try {
      const result = importCSV(store.getState(), wb, 'a,b\n1,2');

      expect(result?.cellsWritten).toBe(3);
      expect(result?.writtenRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'a' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
        kind: 'text',
        value: 'locked',
      });
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 1 });
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 2 });
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });

  it('groups CSV imports into one undo step when history is supplied', () => {
    const history = new History();
    wb.attachHistory(history);
    history.clear();

    const result = importCSV(store.getState(), wb, 'a,b\n1,2', undefined, history);

    expect(result?.cellsWritten).toBe(4);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'a' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 2 });

    expect(history.undo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'blank' });
  });
});

describe('exportCSV', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('serialises the current selection range', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'a' },
      { row: 0, col: 1, value: 'b' },
      { row: 1, col: 0, value: 1 },
      { row: 1, col: 1, value: 2 },
    ]);
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      },
    }));
    expect(exportCSV(store.getState())).toBe('a,b\r\n1,2');
  });

  it('falls back to the used range when selection is a single cell', () => {
    seedAndMirror(store, wb, [
      { row: 5, col: 3, value: 'tl' },
      { row: 7, col: 5, value: 'br' },
    ]);
    expect(exportCSV(store.getState())).toBe('tl,,\r\n,,\r\n,,br');
  });

  it('exports hyperlink display text from hyperlink-only cells', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 4, col: 2 },
      {
        hyperlink: 'https://example.test',
        hyperlinkDisplay: 'Example',
      },
    );

    expect(exportCSV(store.getState())).toBe('Example');
    expect(exportCSV(store.getState(), { range: { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 } })).toBe(
      'Example',
    );
  });

  it('returns empty string when sheet is empty and no range is selected', () => {
    expect(exportCSV(store.getState())).toBe('');
  });

  it('trims huge selections to the materialized used span before exporting', () => {
    seedAndMirror(store, wb, [
      { row: 8, col: 2, value: 'top' },
      { row: 10, col: 2, value: 'bottom' },
      { row: 9, col: 3, value: 'outside' },
    ]);
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: 0, row: 0, col: 2 },
        anchor: { sheet: 0, row: 0, col: 2 },
        range: { sheet: 0, r0: 0, c0: 2, r1: 1_048_575, c1: 2 },
      },
    }));

    expect(exportCSV(store.getState())).toBe('top\r\n\r\nbottom');
  });

  it('refuses sparse huge used spans after trimming oversized exports', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 2, value: 'top' },
      { row: 200_000, col: 2, value: 'bottom' },
    ]);
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: 0, row: 0, col: 2 },
        anchor: { sheet: 0, row: 0, col: 2 },
        range: { sheet: 0, r0: 0, c0: 2, r1: 1_048_575, c1: 2 },
      },
    }));

    expect(exportCSV(store.getState())).toBe('');
  });

  it('respects an explicit range option (overrides selection)', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'a' },
      { row: 0, col: 1, value: 'b' },
    ]);
    expect(exportCSV(store.getState(), { range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 } })).toBe(
      'a',
    );
  });

  it('emits BOM and \\n EOL when requested', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'a' }]);
    const out = exportCSV(store.getState(), {
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      bom: true,
      eol: '\n',
    });
    expect(out.charCodeAt(0)).toBe(0xfeff);
  });

  it('quotes cells that contain commas', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'a, b' }]);
    expect(exportCSV(store.getState(), { range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 } })).toBe(
      '"a, b"',
    );
  });
});
