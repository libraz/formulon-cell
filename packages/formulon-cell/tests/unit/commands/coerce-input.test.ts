import { describe, expect, it, vi } from 'vitest';
import {
  coerceInput,
  writeCoerced,
  writeInput,
  writeInputValidated,
} from '../../../src/commands/coerce-input.js';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('coerceInput', () => {
  it('returns blank for empty / whitespace strings', () => {
    expect(coerceInput('')).toEqual({ kind: 'blank' });
    expect(coerceInput('   ')).toEqual({ kind: 'blank' });
    expect(coerceInput('\t\n')).toEqual({ kind: 'blank' });
  });

  it('returns formula when leading "=" is present', () => {
    expect(coerceInput('=A1+1')).toEqual({ kind: 'formula', text: '=A1+1' });
  });

  it('trims surrounding whitespace before classifying as formula', () => {
    expect(coerceInput('  =SUM(A1:A3)  ')).toEqual({ kind: 'formula', text: '=SUM(A1:A3)' });
  });

  it('returns booleans case-insensitively', () => {
    expect(coerceInput('TRUE')).toEqual({ kind: 'bool', value: true });
    expect(coerceInput('FALSE')).toEqual({ kind: 'bool', value: false });
    expect(coerceInput('true')).toEqual({ kind: 'bool', value: true });
    expect(coerceInput('ｆａｌｓｅ')).toEqual({ kind: 'bool', value: false });
  });

  it('parses integers and decimals as numbers', () => {
    expect(coerceInput('42')).toEqual({ kind: 'number', value: 42 });
    expect(coerceInput('-3.5')).toEqual({ kind: 'number', value: -3.5 });
    expect(coerceInput('1e3')).toEqual({ kind: 'number', value: 1000 });
  });

  it('normalizes full-width numeric input like desktop spreadsheets', () => {
    expect(coerceInput('１２３')).toEqual({ kind: 'number', value: 123 });
    expect(coerceInput('－３．５')).toEqual({ kind: 'number', value: -3.5 });
    expect(coerceInput('１，２３４')).toEqual({ kind: 'number', value: 1234 });
    expect(coerceInput('①')).toEqual({ kind: 'text', value: '①' });
  });

  it('parses percent input as a numeric fraction', () => {
    expect(coerceInput('12%')).toEqual({ kind: 'number', value: 0.12 });
    expect(coerceInput('１２．５％')).toEqual({ kind: 'number', value: 0.125 });
  });

  it('parses common currency-prefixed numeric input', () => {
    expect(coerceInput('$1,234.50')).toEqual({ kind: 'number', value: 1234.5 });
    expect(coerceInput('￥１，２３４')).toEqual({ kind: 'number', value: 1234 });
    expect(coerceInput('¥ 12%')).toEqual({ kind: 'number', value: 0.12 });
  });

  it('parses accounting-style parenthesized negatives', () => {
    expect(coerceInput('(123)')).toEqual({ kind: 'number', value: -123 });
    expect(coerceInput('(￥１，２３４)')).toEqual({ kind: 'number', value: -1234 });
    expect(coerceInput('(12%)')).toEqual({ kind: 'number', value: -0.12 });
  });

  it('parses time input as an spreadsheet serial-day fraction', () => {
    expect(coerceInput('12:00')).toEqual({ kind: 'number', value: 0.5 });
    expect(coerceInput('1:30:00')).toEqual({ kind: 'number', value: 1.5 / 24 });
    expect(coerceInput('２５：００')).toEqual({ kind: 'number', value: 25 / 24 });
    expect(coerceInput('1:30 PM')).toEqual({ kind: 'number', value: 13.5 / 24 });
    expect(coerceInput('12:00 AM')).toEqual({ kind: 'number', value: 0 });
  });

  it('uses a leading apostrophe to force text input', () => {
    expect(coerceInput("'123")).toEqual({ kind: 'text', value: '123' });
    expect(coerceInput("'=A1")).toEqual({ kind: 'text', value: '=A1' });
  });

  it('can force nonblank input to text for preformatted Text cells', () => {
    expect(coerceInput('123', { forceText: true })).toEqual({ kind: 'text', value: '123' });
    expect(coerceInput('=A1', { forceText: true })).toEqual({ kind: 'text', value: '=A1' });
    expect(coerceInput('', { forceText: true })).toEqual({ kind: 'blank' });
  });

  it('preserves the original (non-trimmed) string for text values', () => {
    // Trim is only applied for classification; the text value keeps its original spacing
    // so users can place leading-space text.
    expect(coerceInput('  hello  ')).toEqual({ kind: 'text', value: '  hello  ' });
  });

  it('does not parse partially-numeric strings as numbers', () => {
    expect(coerceInput('12abc')).toEqual({ kind: 'text', value: '12abc' });
    expect(coerceInput('1.2.3')).toEqual({ kind: 'text', value: '1.2.3' });
  });
});

const stubHandle = () => {
  return {
    setBlank: vi.fn(),
    setFormula: vi.fn(),
    setNumber: vi.fn(),
    setBool: vi.fn(),
    setText: vi.fn(),
  } as unknown as WorkbookHandle & {
    setBlank: ReturnType<typeof vi.fn>;
    setFormula: ReturnType<typeof vi.fn>;
    setNumber: ReturnType<typeof vi.fn>;
    setBool: ReturnType<typeof vi.fn>;
    setText: ReturnType<typeof vi.fn>;
  };
};

const addr: Addr = { sheet: 0, row: 0, col: 0 };

describe('writeCoerced', () => {
  it('dispatches each kind to the matching setter', () => {
    const wb = stubHandle();
    writeCoerced(wb, addr, { kind: 'blank' });
    writeCoerced(wb, addr, { kind: 'formula', text: '=1+1' });
    writeCoerced(wb, addr, { kind: 'number', value: 7 });
    writeCoerced(wb, addr, { kind: 'bool', value: true });
    writeCoerced(wb, addr, { kind: 'text', value: 'hi' });
    expect(wb.setBlank).toHaveBeenCalledWith(addr);
    expect(wb.setFormula).toHaveBeenCalledWith(addr, '=1+1');
    expect(wb.setNumber).toHaveBeenCalledWith(addr, 7);
    expect(wb.setBool).toHaveBeenCalledWith(addr, true);
    expect(wb.setText).toHaveBeenCalledWith(addr, 'hi');
  });
});

describe('writeInput', () => {
  it('coerces and writes through the workbook handle', () => {
    const wb = stubHandle();
    writeInput(wb, addr, '=A1');
    writeInput(wb, addr, '');
    writeInput(wb, addr, '3.14');
    expect(wb.setFormula).toHaveBeenCalledWith(addr, '=A1');
    expect(wb.setBlank).toHaveBeenCalledWith(addr);
    expect(wb.setNumber).toHaveBeenCalledWith(addr, 3.14);
  });

  it('normalizes R1C1 formula input to A1 before writing to the workbook', () => {
    const wb = stubHandle();

    writeInput(wb, { sheet: 0, row: 3, col: 3 }, '=SUM(R[-2]C[-2]:R[-1]C[-1])');

    expect(wb.setFormula).toHaveBeenCalledWith({ sheet: 0, row: 3, col: 3 }, '=SUM(B2:C3)');
  });

  it('normalizes R1C1 formula input before workbook calculation', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 2);
    wb.setNumber({ sheet: 0, row: 1, col: 2 }, 3);
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 4);
    wb.setNumber({ sheet: 0, row: 2, col: 2 }, 5);

    writeInput(wb, { sheet: 0, row: 3, col: 3 }, '=SUM(R[-2]C[-2]:R[-1]C[-1])');
    wb.recalc();

    expect(wb.cellFormula({ sheet: 0, row: 3, col: 3 })).toBe('=SUM(B2:C3)');
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 14 });
  });

  it('writes numeric and formula-looking input as text when the cell format is Text', () => {
    const wb = stubHandle();
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, addr, { numFmt: { kind: 'text' } });

    writeInput(wb, addr, '00123', store);
    writeInput(wb, addr, '=A1', store);

    expect(wb.setText).toHaveBeenCalledWith(addr, '00123');
    expect(wb.setText).toHaveBeenCalledWith(addr, '=A1');
    expect(wb.setNumber).not.toHaveBeenCalled();
    expect(wb.setFormula).not.toHaveBeenCalled();
  });
});

describe('writeInputValidated', () => {
  it('blocks invalid stop-style input when error alerts are enabled', () => {
    const wb = stubHandle();

    const outcome = writeInputValidated(wb, addr, 'Closed', {
      kind: 'list',
      source: ['Open'],
      errorStyle: 'stop',
    });

    expect(outcome.ok).toBe(false);
    expect(wb.setText).not.toHaveBeenCalled();
  });

  it('writes invalid input when the validation error alert is disabled', () => {
    const wb = stubHandle();

    const outcome = writeInputValidated(wb, addr, 'Closed', {
      kind: 'list',
      source: ['Open'],
      errorStyle: 'stop',
      showErrorMessage: false,
    });

    expect(outcome.ok).toBe(false);
    expect(wb.setText).toHaveBeenCalledWith(addr, 'Closed');
  });
});
