import { describe, expect, it } from 'vitest';
import {
  computeF9Preview,
  renderCellValueForF9,
  replaceFormulaSelectionWithF9Preview,
} from '../../../src/commands/f9-preview.js';
import {
  type CellValue,
  type EvalArrayResult,
  type EvalResult,
  type Value,
  ValueKind,
} from '../../../src/engine/types.js';

const numCell = (value: number): { value: CellValue; formula: string | null } => ({
  value: { kind: 'number', value },
  formula: null,
});

const txtCell = (value: string): { value: CellValue; formula: string | null } => ({
  value: { kind: 'text', value },
  formula: null,
});

const evalNumber =
  (expectedFormula: string, value: number) =>
  (formula: string): EvalResult => {
    expect(formula).toBe(expectedFormula);
    return {
      status: { status: 0 },
      value: { kind: ValueKind.Number, number: value },
    } as EvalResult;
  };

const numV = (n: number): Value => ({ kind: ValueKind.Number, number: n }) as Value;
const txtV = (t: string): Value => ({ kind: ValueKind.Text, text: t }) as Value;

const evalArray =
  (expectedFormula: string, cells: Value[][]) =>
  (formula: string): EvalArrayResult => {
    expect(formula).toBe(expectedFormula);
    return {
      status: { status: 0 },
      rows: cells.length,
      cols: cells[0]?.length ?? 0,
      cells,
    } as EvalArrayResult;
  };

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

  it('evaluates a single-cell-ref expression when an engine evaluator is supplied', () => {
    const cells = new Map([
      ['0:0:0', numCell(7)],
      ['0:0:1', numCell(5)],
    ]);
    expect(computeF9Preview('=A1+B1', 'A1+B1', 0, cells, undefined, evalNumber('7+5', 12))).toEqual(
      {
        display: '12',
        substitutable: true,
      },
    );
  });

  it('substitutes sheet-qualified refs before engine-backed F9 evaluation', () => {
    const cells = new Map([
      ['0:0:0', numCell(7)],
      ['1:0:0', numCell(5)],
    ]);
    const sheetByName = (n: string): number => (n === 'Other' ? 1 : -1);
    expect(
      computeF9Preview('=A1+Other!A1', 'A1+Other!A1', 0, cells, sheetByName, evalNumber('7+5', 12)),
    ).toEqual({
      display: '12',
      substitutable: true,
    });
  });

  it('does not treat digit-suffixed function names as refs during engine-backed evaluation', () => {
    const cells = new Map([['0:0:0', numCell(100)]]);
    expect(
      computeF9Preview('=LOG10(A1)', 'LOG10(A1)', 0, cells, undefined, evalNumber('LOG10(100)', 2)),
    ).toEqual({
      display: '2',
      substitutable: true,
    });
  });

  it('expands range refs before engine-backed F9 evaluation', () => {
    const cells = new Map([
      ['0:0:0', numCell(7)],
      ['0:1:0', numCell(5)],
      ['0:2:0', numCell(3)],
    ]);
    expect(
      computeF9Preview(
        '=SUM(A1:A3)',
        'SUM(A1:A3)',
        0,
        cells,
        undefined,
        evalNumber('SUM(7,5,3)', 15),
      ),
    ).toEqual({
      display: '15',
      substitutable: true,
    });
  });

  it('expands sheet-qualified range refs before engine-backed F9 evaluation', () => {
    const cells = new Map([
      ['1:0:0', numCell(2)],
      ['1:0:1', numCell(4)],
    ]);
    const sheetByName = (n: string): number => (n === 'Other' ? 1 : -1);
    expect(
      computeF9Preview(
        '=SUM(Other!A1:B1)',
        'SUM(Other!A1:B1)',
        0,
        cells,
        sheetByName,
        evalNumber('SUM(2,4)', 6),
      ),
    ).toEqual({
      display: '6',
      substitutable: true,
    });
  });

  it('keeps oversized ranges unsupported even when an engine evaluator is supplied', () => {
    expect(
      computeF9Preview(
        '=SUM(A1:A10001)',
        'SUM(A1:A10001)',
        0,
        new Map(),
        undefined,
        evalNumber('', 0),
      ),
    ).toEqual({
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

  it('renders a multi-cell array result as a spreadsheet array constant', () => {
    expect(
      computeF9Preview(
        '=SEQUENCE(2,2)',
        'SEQUENCE(2,2)',
        0,
        new Map(),
        undefined,
        undefined,
        true,
        evalArray('=SEQUENCE(2,2)', [
          [numV(1), numV(2)],
          [numV(3), numV(4)],
        ]),
      ),
    ).toEqual({ display: '{1,2;3,4}', substitutable: true });
  });

  it('quotes text cells inside an array constant', () => {
    expect(
      computeF9Preview(
        '={"a","b"}',
        '{"a","b"}',
        0,
        new Map(),
        undefined,
        undefined,
        true,
        evalArray('={"a","b"}', [[txtV('a'), txtV('b')]]),
      ),
    ).toEqual({ display: '{"a","b"}', substitutable: true });
  });

  it('collapses a 1x1 array result to the bare scalar', () => {
    expect(
      computeF9Preview(
        '=A1',
        'A1+0',
        0,
        new Map(),
        undefined,
        undefined,
        true,
        evalArray('=A1+0', [[numV(42)]]),
      ),
    ).toEqual({ display: '42', substitutable: true });
  });

  it('falls back to the scalar evaluator when array evaluation errors', () => {
    const arrErr =
      () =>
      (_formula: string): EvalArrayResult =>
        ({ status: { status: 1 }, rows: 0, cols: 0, cells: [] as Value[][] }) as EvalArrayResult;
    expect(
      computeF9Preview(
        '=A1+B1',
        'A1+B1',
        0,
        new Map([
          ['0:0:0', numCell(7)],
          ['0:0:1', numCell(5)],
        ]),
        undefined,
        evalNumber('=A1+B1', 12),
        true,
        arrErr(),
      ),
    ).toEqual({ display: '12', substitutable: true });
  });
});

describe('replaceFormulaSelectionWithF9Preview', () => {
  it('replaces a selected reference with its F9 value and moves the caret after it', () => {
    const cells = new Map([['0:0:0', numCell(7)]]);
    expect(replaceFormulaSelectionWithF9Preview('=A1+1', 1, 3, 0, cells)).toEqual({
      text: '=7+1',
      start: 2,
      end: 2,
      preview: { display: '7', substitutable: true },
    });
  });

  it('does not replace unsupported expressions but returns the preview outcome', () => {
    expect(replaceFormulaSelectionWithF9Preview('=A1+B1', 1, 6, 0, new Map())).toEqual({
      text: '=A1+B1',
      start: 1,
      end: 6,
      preview: { display: '', substitutable: false },
    });
  });

  it('ignores non-formula text and collapsed selections', () => {
    expect(replaceFormulaSelectionWithF9Preview('A1', 0, 2, 0, new Map())).toBeNull();
    expect(replaceFormulaSelectionWithF9Preview('=A1', 1, 1, 0, new Map())).toBeNull();
  });
});
