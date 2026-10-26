import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  applyPrecisionAsDisplayed,
  displayedPrecisionValue,
} from '../../../src/commands/precision-as-displayed.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

describe('displayedPrecisionValue', () => {
  it('rounds to the decimals a fixed, currency or accounting format shows', () => {
    expect(displayedPrecisionValue(1.23456, { kind: 'fixed', decimals: 2 })).toBe(1.23);
    expect(displayedPrecisionValue(1.005, { kind: 'fixed', decimals: 2 })).toBe(1.01);
    expect(displayedPrecisionValue(-1.005, { kind: 'fixed', decimals: 2 })).toBe(-1.01);
    expect(displayedPrecisionValue(2.5, { kind: 'currency', decimals: 0 })).toBe(3);
    expect(displayedPrecisionValue(9.876, { kind: 'accounting', decimals: 1 })).toBe(9.9);
  });

  it('stays finite for values that print in exponent form', () => {
    const two = { kind: 'fixed', decimals: 2 } as const;
    expect(displayedPrecisionValue(1e-7, two)).toBe(0);
    expect(displayedPrecisionValue(-1e-7, two)).toBe(0);
    expect(displayedPrecisionValue(1e22, two)).toBe(1e22);
    expect(displayedPrecisionValue(-1e22, two)).toBe(-1e22);
    expect(displayedPrecisionValue(1e19, two)).toBe(1e19);
    expect(displayedPrecisionValue(Number.MAX_VALUE, two)).toBe(Number.MAX_VALUE);
    expect(displayedPrecisionValue(-2.345, two)).toBe(-2.35);
    expect(displayedPrecisionValue(1.5e-7, { kind: 'percent', decimals: 5 })).toBe(0.0000002);
    expect(displayedPrecisionValue(1e-300, { kind: 'scientific', decimals: 1 })).toBe(1e-300);
  });

  it('rounds percentages on the displayed scale', () => {
    expect(displayedPrecisionValue(0.12345, { kind: 'percent', decimals: 0 })).toBe(0.12);
    expect(displayedPrecisionValue(0.12345, { kind: 'percent', decimals: 1 })).toBe(0.123);
  });

  it('keeps the significant digits of a scientific format', () => {
    expect(displayedPrecisionValue(123456, { kind: 'scientific', decimals: 1 })).toBe(120000);
  });

  it('leaves General, text, date, time and custom formats at full precision', () => {
    expect(displayedPrecisionValue(1.23456, undefined)).toBe(1.23456);
    expect(displayedPrecisionValue(1.23456, { kind: 'general' })).toBe(1.23456);
    expect(displayedPrecisionValue(1.23456, { kind: 'text' })).toBe(1.23456);
    expect(displayedPrecisionValue(45000.75, { kind: 'date', pattern: 'yyyy-mm-dd' })).toBe(
      45000.75,
    );
    expect(displayedPrecisionValue(1.23456, { kind: 'custom', pattern: '0.0' })).toBe(1.23456);
  });
});

describe.skipIf(!canLoadWasm())('applyPrecisionAsDisplayed', () => {
  it('rounds formatted constants on every sheet, leaves formulas, and undoes in one step', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.addSheet('Two');
      const store = createSpreadsheetStore();
      const history = new History();
      wb.attachHistory(history);
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1.23456);
      wb.setNumber({ sheet: 0, row: 1, col: 0 }, 1.23456);
      wb.setFormula({ sheet: 0, row: 2, col: 0 }, '=A1*1');
      wb.setNumber({ sheet: 1, row: 0, col: 0 }, 0.12345);
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 0 },
        {
          numFmt: { kind: 'fixed', decimals: 2 },
        },
      );
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 2, col: 0 },
        {
          numFmt: { kind: 'fixed', decimals: 1 },
        },
      );
      mutators.setCellFormat(
        store,
        { sheet: 1, row: 0, col: 0 },
        {
          numFmt: { kind: 'percent', decimals: 0 },
        },
      );
      history.clear();

      expect(applyPrecisionAsDisplayed(wb, store, history)).toBe(2);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1.23 });
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
        kind: 'number',
        value: 1.23456,
      });
      expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'number', value: 1.23 });
      expect(wb.getValue({ sheet: 1, row: 0, col: 0 })).toEqual({ kind: 'number', value: 0.12 });

      history.undo();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'number',
        value: 1.23456,
      });
      expect(wb.getValue({ sheet: 1, row: 0, col: 0 })).toEqual({
        kind: 'number',
        value: 0.12345,
      });
    } finally {
      wb.dispose();
    }
  });
});
