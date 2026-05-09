import { describe, expect, it } from 'vitest';
import { computeF9Preview, renderCellValueForF9 } from '../../../src/commands/f9-preview.js';
import type { CellValue } from '../../../src/engine/types.js';

const numCell = (value: number): { value: CellValue; formula: string | null } => ({
  value: { kind: 'number', value },
  formula: null,
});

const txtCell = (value: string): { value: CellValue; formula: string | null } => ({
  value: { kind: 'text', value },
  formula: null,
});

describe('renderCellValueForF9', () => {
  it('renders number as decimal string', () => {
    expect(renderCellValueForF9({ kind: 'number', value: 3.14 })).toBe('3.14');
  });

  it('quotes text values', () => {
    expect(renderCellValueForF9({ kind: 'text', value: 'hi' })).toBe('"hi"');
  });

  it('uppercases booleans (spreadsheet parity)', () => {
    expect(renderCellValueForF9({ kind: 'bool', value: true })).toBe('TRUE');
    expect(renderCellValueForF9({ kind: 'bool', value: false })).toBe('FALSE');
  });

  it('passes through error text when present', () => {
    expect(renderCellValueForF9({ kind: 'error', code: 7, text: '#DIV/0!' })).toBe('#DIV/0!');
  });

  it('treats blank/undefined as 0 (spreadsheets coerce blank to zero)', () => {
    expect(renderCellValueForF9(undefined)).toBe('0');
    expect(renderCellValueForF9({ kind: 'blank' })).toBe('0');
  });
});

describe('computeF9Preview', () => {
  it('substitutes a numeric literal verbatim', () => {
    const r = computeF9Preview('=42 + A1', '42', 0, new Map());
    expect(r).toEqual({ display: '42', substitutable: true });
  });

  it('substitutes a quoted string literal verbatim', () => {
    const r = computeF9Preview('=A1 & "x"', '"x"', 0, new Map());
    expect(r).toEqual({ display: '"x"', substitutable: true });
  });

  it('substitutes TRUE / false (case-normalized)', () => {
    expect(computeF9Preview('=A1', 'true', 0, new Map())).toEqual({
      display: 'TRUE',
      substitutable: true,
    });
    expect(computeF9Preview('=A1', 'FALSE', 0, new Map())).toEqual({
      display: 'FALSE',
      substitutable: true,
    });
  });

  it('resolves a single-cell ref to its current value', () => {
    const cells = new Map([['0:0:0', numCell(7)]]);
    expect(computeF9Preview('=A1+1', 'A1', 0, cells)).toEqual({
      display: '7',
      substitutable: true,
    });
  });

  it('quotes a text cell when resolving via ref', () => {
    const cells = new Map([['0:1:0', txtCell('hello')]]);
    expect(computeF9Preview('=A2', 'A2', 0, cells)).toEqual({
      display: '"hello"',
      substitutable: true,
    });
  });

  it('returns 0 for an unpopulated cell ref (spreadsheet convention: blank → 0)', () => {
    expect(computeF9Preview('=A1', 'A1', 0, new Map())).toEqual({
      display: '0',
      substitutable: true,
    });
  });

  it('honours sheetByName for sheet-qualified refs', () => {
    const cells = new Map([['1:0:0', numCell(99)]]);
    const sheetByName = (n: string): number => (n === 'Other' ? 1 : -1);
    expect(computeF9Preview('=Other!A1', 'Other!A1', 0, cells, sheetByName)).toEqual({
      display: '99',
      substitutable: true,
    });
  });

  it('returns #REF! when a sheet name does not resolve', () => {
    expect(computeF9Preview('=Z!A1', 'Z!A1', 0, new Map())).toEqual({
      display: '#REF!',
      substitutable: false,
    });
  });

  it('reports unsupported for sub-expressions / function calls', () => {
    expect(computeF9Preview('=SUM(A1:A3)', 'SUM(A1:A3)', 0, new Map())).toEqual({
      display: '',
      substitutable: false,
    });
    expect(computeF9Preview('=A1+B1', 'A1+B1', 0, new Map())).toEqual({
      display: '',
      substitutable: false,
    });
  });

  it('reports unsupported for an empty selection', () => {
    expect(computeF9Preview('=1', '   ', 0, new Map())).toEqual({
      display: '',
      substitutable: false,
    });
  });
});
