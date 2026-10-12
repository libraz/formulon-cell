import { describe, expect, it } from 'vitest';
import { setFont } from '../../../src/commands/format.js';
import { MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
import {
  activeSheetFontBaseline,
  formatWithSheetFontPatch,
  isWholeSheetSelection,
  syncWholeSheetFontOverride,
} from '../../../src/engine/whole-sheet-font-sync.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const whole = { sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: MAX_COL } as const;

describe('whole-sheet font planning', () => {
  it('recognises the exact worksheet union without materialising cells', () => {
    expect(isWholeSheetSelection([whole])).toBe(true);
    expect(
      isWholeSheetSelection([
        { sheet: 0, r0: 0, c0: 0, r1: 10, c1: MAX_COL },
        { sheet: 0, r0: 11, c0: 0, r1: MAX_ROW, c1: MAX_COL },
      ]),
    ).toBe(true);
    expect(isWholeSheetSelection([{ sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: MAX_COL - 1 }])).toBe(
      false,
    );
    expect(
      isWholeSheetSelection([
        { sheet: 0, r0: 0, c0: 0, r1: 10, c1: MAX_COL },
        { sheet: 0, r0: 12, c0: 0, r1: MAX_ROW, c1: MAX_COL },
      ]),
    ).toBe(false);
  });

  it('keeps an explicit cell font axis above the worksheet default', () => {
    expect(
      formatWithSheetFontPatch(
        { bold: true, fontSize: 24 },
        { fontFamily: '游ゴシック', fontSize: 18 },
      ),
    ).toEqual({ bold: true, fontSize: 24, fontFamily: '游ゴシック' });
    expect(formatWithSheetFontPatch({ bold: true }, { fontSize: 18 })).toEqual({
      bold: true,
      fontSize: 18,
    });
  });

  it('updates sparse physical entries while keeping blank cells implicit', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active: { sheet: 0, row: 0, col: 0 },
        range: whole,
      },
    }));
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'A' });
    mutators.setCell(store, { sheet: 0, row: 4, col: 4 }, { kind: 'text', value: 'B' });
    mutators.setCellFormat(store, { sheet: 0, row: 8, col: 8 }, { bold: true });

    expect(setFont(store.getState(), store, { fontSize: 18 })).toBe(true);
    const formats = store.getState().format.formats;
    expect(formats.size).toBe(3);
    expect(formats.get('0:0:0')?.fontSize).toBe(18);
    expect(formats.get('0:4:4')?.fontSize).toBe(18);
    expect(formats.get('0:8:8')).toEqual({ bold: true, fontSize: 18 });
    expect(store.getState().format.sheetFontOverrides?.get(0)?.patch).toEqual({ fontSize: 18 });
  });

  it('rejects invalid whole-sheet font axes atomically', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      selection: { ...state.selection, range: whole },
    }));
    expect(setFont(store.getState(), store, { fontSize: Number.NaN })).toBe(false);
    expect(setFont(store.getState(), store, { fontSize: Number.POSITIVE_INFINITY })).toBe(false);
    expect(setFont(store.getState(), store, { fontSize: 0 })).toBe(false);
    expect(setFont(store.getState(), store, { fontFamily: '   ' })).toBe(false);
    expect(store.getState().format.sheetFontOverrides?.size ?? 0).toBe(0);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('remaps the workbook-local baseline when sheet indexes move', () => {
    const store = createSpreadsheetStore();
    const baseline = {
      columns: [{ first: 0, last: 3, styleXf: 7 }],
      rows: [{ row: 4, styleXf: 8 }],
    };
    store.setState((state) => ({
      ...state,
      format: {
        ...state.format,
        sheetFontOverrides: new Map([[0, { patch: { fontSize: 18 } }]]),
      },
    }));
    const calls: string[] = [];
    let notifySheetMove:
      | ((event: { kind: 'sheet-move'; from: number; to: number }) => void)
      | undefined;
    const wb = {
      capabilities: { sheetFontStyles: true },
      workbookDefaultFont: { name: 'Aptos', size: 11 },
      subscribe: (listener: (event: { kind: 'sheet-move'; from: number; to: number }) => void) => {
        notifySheetMove = listener;
        return () => undefined;
      },
      getSheetFontAxisSnapshot: () => baseline,
      getCellXf: () => ({ fontIndex: 0 }),
      getFontRecord: () => ({ name: 'Aptos', size: 11 }),
      addFontRecord: () => 9,
      addXfRecord: () => 10,
      clearColumnStyle: () => {
        calls.push('clear-columns');
        return true;
      },
      setColumnStyle: () => {
        calls.push('set-columns');
        return true;
      },
      setRowStyle: () => {
        calls.push('set-row');
        return true;
      },
      clearRowStyle: () => true,
    } as unknown as WorkbookHandle;

    expect(syncWholeSheetFontOverride(wb, store, 0)).toBe(true);
    expect(activeSheetFontBaseline(wb, 0)).toEqual(baseline);
    notifySheetMove?.({ kind: 'sheet-move', from: 0, to: 2 });
    expect(activeSheetFontBaseline(wb, 0)).toBeUndefined();
    expect(activeSheetFontBaseline(wb, 2)).toEqual(baseline);
    expect(calls).toContain('set-columns');
  });

  it('clears axis visuals from XF0 while retaining raw protection metadata', () => {
    const store = createSpreadsheetStore();
    const baseline = {
      columns: [{ first: 0, last: MAX_COL, styleXf: 7 }],
      rows: [{ row: 4, styleXf: 8 }],
    };
    store.setState((state) => ({
      ...state,
      format: {
        ...state.format,
        sheetFontOverrides: new Map([[0, { patch: {}, clearAxes: true, baseline }]]),
      },
    }));
    const added: Record<string, unknown>[] = [];
    const setColumns: number[] = [];
    const setRows: number[] = [];
    const records = new Map<number, Record<string, unknown>>();
    records.set(0, {
      fontIndex: 0,
      fillIndex: 0,
      borderIndex: 0,
      numFmtId: 0,
      horizontalAlign: 0,
      verticalAlign: 0,
      wrapText: false,
    });
    records.set(7, {
      fontIndex: 1,
      fillIndex: 4,
      borderIndex: 5,
      numFmtId: 14,
      horizontalAlign: 2,
      verticalAlign: 1,
      wrapText: true,
      applyProtection: true,
      hasProtection: true,
      locked: false,
      hidden: true,
      quotePrefix: true,
    });
    records.set(8, { ...records.get(7), locked: true, hidden: false });
    const wb = {
      capabilities: { sheetFontStyles: true },
      subscribe: () => () => undefined,
      getCellXf: (index: number) => records.get(index) ?? null,
      addXfRecord: (record: Record<string, unknown>) => {
        added.push(record);
        return 20 + added.length;
      },
      clearColumnStyle: () => true,
      clearRowStyle: () => true,
      setColumnStyle: (_sheet: number, _first: number, _last: number, xf: number) => {
        setColumns.push(xf);
        return true;
      },
      setRowStyle: (_sheet: number, _row: number, xf: number) => {
        setRows.push(xf);
        return true;
      },
    } as unknown as WorkbookHandle;

    expect(syncWholeSheetFontOverride(wb, store, 0)).toBe(true);
    expect(added).toHaveLength(2);
    expect(added[0]).toMatchObject({
      fillIndex: 0,
      borderIndex: 0,
      numFmtId: 0,
      applyProtection: true,
      hasProtection: true,
      locked: false,
      hidden: true,
      quotePrefix: true,
    });
    expect(added[1]).toMatchObject({ locked: true, hidden: false });
    expect(setColumns).toHaveLength(1);
    expect(setRows).toHaveLength(1);
  });
});
