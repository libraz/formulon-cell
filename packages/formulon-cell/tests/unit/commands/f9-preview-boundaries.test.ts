import { describe, expect, it } from 'vitest';
import {
  computeF9Preview,
  replaceFormulaSelectionWithF9Preview,
} from '../../../src/commands/f9-preview.js';
import type { CellValue, EvalArrayResult, EvalResult, Value } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const numCell = (value: number): { value: CellValue; formula: string | null } => ({
  value: { kind: 'number', value },
  formula: null,
});
const textCell = (value: string): { value: CellValue; formula: string | null } => ({
  value: { kind: 'text', value },
  formula: null,
});
const numberResult = (value: number): EvalResult =>
  ({ status: { status: 0 }, value: { kind: 1, number: value } }) as EvalResult;
const textResult = (value: string): EvalResult =>
  ({ status: { status: 0 }, value: { kind: 3, text: value } }) as EvalResult;

const quotedText = 'say "hi"\\literal\nnext';
const quotedFormulaLiteral = '"say ""hi""\\literal\nnext"';

describe('F9 formula edit boundaries', () => {
  it('retains the equals sign when the selection includes the formula prefix', () => {
    const formula = '=1+1+2';
    const evaluate = (value: string): EvalResult => {
      expect(value).toBe('=1+1');
      return numberResult(2);
    };
    const replacement = replaceFormulaSelectionWithF9Preview(
      formula,
      0,
      4,
      0,
      new Map(),
      undefined,
      evaluate,
      true,
    );
    expect(replacement?.text).toBe('=2+2');
  });

  it('keeps a reversed selection inside the formula body and ignores a selection of only equals', () => {
    const evaluate = (value: string): EvalResult => {
      expect(value).toBe('=1+1');
      return numberResult(2);
    };
    expect(
      replaceFormulaSelectionWithF9Preview('=1+1+2', 4, 0, 0, new Map(), undefined, evaluate, true)
        ?.text,
    ).toBe('=2+2');
    expect(replaceFormulaSelectionWithF9Preview('=1+1', 0, 1, 0, new Map())).toBeNull();
  });

  it('uses spreadsheet quote doubling for text cell references', () => {
    const cells = new Map([['0:0:0', textCell(quotedText)]]);
    expect(computeF9Preview('=A1', 'A1', 0, cells).display).toBe(quotedFormulaLiteral);
  });

  it('uses spreadsheet quote doubling for scalar evaluator results', () => {
    const evaluate = (): EvalResult => textResult(quotedText);
    expect(computeF9Preview('=1+1', '1+1', 0, new Map(), undefined, evaluate, true)).toEqual({
      display: quotedFormulaLiteral,
      substitutable: true,
    });
  });

  it('uses spreadsheet quote doubling for array evaluator results', () => {
    const array: EvalArrayResult = {
      status: { status: 0 },
      rows: 1,
      cols: 1,
      cells: [[{ kind: 3, number: 0, boolean: 0, text: quotedText, errorCode: 0 } as Value]],
    } as EvalArrayResult;
    expect(
      computeF9Preview('=1+0', '1+0', 0, new Map(), undefined, undefined, true, () => array),
    ).toEqual({ display: quotedFormulaLiteral, substitutable: true });
  });

  it('uses spreadsheet quote doubling in reference-substitution fallback expressions', () => {
    const cells = new Map([['0:0:0', textCell(quotedText)]]);
    const evaluate = (value: string): EvalResult => {
      expect(value).toBe(`${quotedFormulaLiteral}+0`);
      return numberResult(1);
    };
    expect(computeF9Preview('=A1', 'A1+0', 0, cells, undefined, evaluate, false)).toEqual({
      display: '1',
      substitutable: true,
    });
  });

  it('resolves AA1 and AB2 as unqualified multi-letter-column references', () => {
    const cells = new Map([
      ['0:0:26', numCell(42)],
      ['0:1:27', numCell(84)],
    ]);
    expect(computeF9Preview('=AA1', 'AA1', 0, cells)).toEqual({
      display: '42',
      substitutable: true,
    });
    expect(computeF9Preview('=AB2', 'AB2', 0, cells)).toEqual({
      display: '84',
      substitutable: true,
    });
  });

  it('expands a range in multi-letter columns before fallback evaluation', () => {
    const cells = new Map([
      ['0:0:26', numCell(1)],
      ['0:0:27', numCell(2)],
      ['0:1:26', numCell(3)],
      ['0:1:27', numCell(4)],
    ]);
    const evaluate = (value: string): EvalResult => {
      expect(value).toBe('SUM(1,2,3,4)');
      return numberResult(10);
    };
    expect(
      computeF9Preview('=SUM(AA1:AB2)', 'SUM(AA1:AB2)', 0, cells, undefined, evaluate),
    ).toEqual({ display: '10', substitutable: true });
  });

  it('evaluates the expanded multi-letter range with the real workbook engine', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      const cells = new Map([
        ['0:0:26', numCell(1)],
        ['0:0:27', numCell(2)],
        ['0:1:26', numCell(3)],
        ['0:1:27', numCell(4)],
      ]);
      const preview = computeF9Preview(
        '=SUM(AA1:AB2)',
        'SUM(AA1:AB2)',
        0,
        cells,
        undefined,
        (formula) => workbook.evalFormula(formula.startsWith('=') ? formula : `=${formula}`),
      );
      expect(preview).toEqual({ display: '10', substitutable: true });
    } finally {
      workbook.dispose();
    }
  });

  it('resolves an escaped quoted sheet name with a multi-letter column', () => {
    const cells = new Map([['1:0:26', numCell(17)]]);
    const sheetByName = (name: string): number => (name === "O'Brien" ? 1 : -1);
    expect(computeF9Preview("='O''Brien'!AA1", "'O''Brien'!AA1", 0, cells, sheetByName)).toEqual({
      display: '17',
      substitutable: true,
    });
  });
});
