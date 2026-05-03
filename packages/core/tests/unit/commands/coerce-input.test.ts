import { describe, expect, it, vi } from 'vitest';
import { coerceInput, writeCoerced, writeInput } from '../../../src/commands/coerce-input.js';
import type { Addr } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

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

  it('returns booleans for canonical TRUE / FALSE only', () => {
    expect(coerceInput('TRUE')).toEqual({ kind: 'bool', value: true });
    expect(coerceInput('FALSE')).toEqual({ kind: 'bool', value: false });
    // Lowercase is not coerced — Excel treats it as text.
    expect(coerceInput('true')).toEqual({ kind: 'text', value: 'true' });
  });

  it('parses integers and decimals as numbers', () => {
    expect(coerceInput('42')).toEqual({ kind: 'number', value: 42 });
    expect(coerceInput('-3.5')).toEqual({ kind: 'number', value: -3.5 });
    expect(coerceInput('1e3')).toEqual({ kind: 'number', value: 1000 });
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
});
