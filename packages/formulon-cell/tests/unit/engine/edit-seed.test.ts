import { describe, expect, it } from 'vitest';
import { formatCellForEdit } from '../../../src/engine/edit-seed.js';
import type { Addr, CellValue } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const num = (n: number): CellValue => ({ kind: 'number', value: n });
const text = (s: string): CellValue => ({ kind: 'text', value: s });
const err = (code: number, txt: string): CellValue => ({ kind: 'error', code, text: txt });
const blank = (): CellValue => ({ kind: 'blank' });

const wbWithLambda = (text: string | null): WorkbookHandle =>
  ({
    getLambdaText: (_a: Addr) => text,
  }) as unknown as WorkbookHandle;

describe('formatCellForEdit', () => {
  it('returns the formula verbatim when present', () => {
    expect(formatCellForEdit({ value: num(42), formula: '=SUM(A1:A3)' })).toBe('=SUM(A1:A3)');
  });

  it('renders a number without formatting', () => {
    expect(formatCellForEdit({ value: num(1234.56), formula: null })).toBe('1234.56');
  });

  it('renders booleans as TRUE / FALSE', () => {
    expect(formatCellForEdit({ value: { kind: 'bool', value: true }, formula: null })).toBe('TRUE');
    expect(formatCellForEdit({ value: { kind: 'bool', value: false }, formula: null })).toBe(
      'FALSE',
    );
  });

  it('renders text values verbatim', () => {
    expect(formatCellForEdit({ value: text('hello'), formula: null })).toBe('hello');
  });

  it('renders the error sentinel for error cells', () => {
    expect(formatCellForEdit({ value: err(2, '#VALUE!'), formula: null })).toBe('#VALUE!');
  });

  it('returns empty string for missing cells', () => {
    expect(formatCellForEdit(undefined)).toBe('');
    expect(formatCellForEdit({ value: blank(), formula: null })).toBe('');
  });

  it('falls through to lambda body when value kind is blank but engine reports a lambda', () => {
    const cell = { value: blank(), formula: null };
    const wb = wbWithLambda('LAMBDA(x, x+1)');
    expect(formatCellForEdit(cell, wb, { sheet: 0, row: 0, col: 0 })).toBe('=LAMBDA(x, x+1)');
  });

  it('does not call getLambdaText when a formula is already present', () => {
    let called = 0;
    const wb = {
      getLambdaText: () => {
        called += 1;
        return 'LAMBDA()';
      },
    } as unknown as WorkbookHandle;
    expect(
      formatCellForEdit({ value: blank(), formula: '=A1' }, wb, { sheet: 0, row: 0, col: 0 }),
    ).toBe('=A1');
    expect(called).toBe(0);
  });

  it('falls through to value formatting when getLambdaText returns null', () => {
    const wb = wbWithLambda(null);
    expect(
      formatCellForEdit({ value: num(7), formula: null }, wb, { sheet: 0, row: 0, col: 0 }),
    ).toBe('7');
  });
});
