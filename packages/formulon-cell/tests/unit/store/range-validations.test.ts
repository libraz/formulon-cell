import { describe, expect, it } from 'vitest';
import { shiftRangeValidations } from '../../../src/commands/axis-shift.js';
import { clearValidationInRange, validateAgainst } from '../../../src/commands/validate.js';
import { MAX_ROW } from '../../../src/engine/address.js';
import {
  subtractFromRangeValidations,
  validationAt,
} from '../../../src/store/range-validations.js';
import { type CellValidation, createSpreadsheetStore } from '../../../src/store/store.js';

const whole: CellValidation = { kind: 'whole', op: 'between', a: 0, b: 9 };
const column = { sheet: 0, r0: 0, c0: 2, r1: MAX_ROW, c1: 2 };

describe('range-stored validation rules', () => {
  it('a cell rule wins over a range rule', () => {
    const store = createSpreadsheetStore();
    const own: CellValidation = { kind: 'list', source: ['x'] };
    store.setState((s) => ({
      ...s,
      format: {
        ...s.format,
        formats: new Map([['0:4:2', { validation: own }]]),
        rangeValidations: [{ ranges: [column], validation: whole }],
      },
    }));
    const fmt = store.getState().format;
    expect(validationAt(fmt, { sheet: 0, row: 4, col: 2 })).toBe(own);
    expect(validationAt(fmt, { sheet: 0, row: 5, col: 2 })).toBe(whole);
    expect(validationAt(fmt, { sheet: 1, row: 5, col: 2 })).toBeUndefined();
  });

  it('clearing a block splits the rule around it', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: { ...s.format, rangeValidations: [{ ranges: [column], validation: whole }] },
    }));
    clearValidationInRange(store, { sheet: 0, r0: 10, c0: 2, r1: 19, c1: 2 });
    const fmt = store.getState().format;
    expect(validationAt(fmt, { sheet: 0, row: 12, col: 2 })).toBeUndefined();
    expect(validationAt(fmt, { sheet: 0, row: 9, col: 2 })).toBe(whole);
    expect(validationAt(fmt, { sheet: 0, row: 20, col: 2 })).toBe(whole);
    expect(subtractFromRangeValidations(fmt.rangeValidations, column)).toEqual([]);
  });

  it('survives a row insert as a whole-column rule and shifts a bounded one', () => {
    const bounded = { sheet: 0, r0: 5, c0: 0, r1: 9, c1: 0 };
    const rules = [
      { ranges: [column], validation: whole },
      { ranges: [bounded], validation: whole },
    ];
    const shifted = shiftRangeValidations(rules, 0, 'row', 3, 2);
    expect(shifted?.[0]?.ranges).toEqual([column]);
    expect(shifted?.[1]?.ranges).toEqual([{ ...bounded, r0: 7, r1: 11 }]);
  });
});

describe('formula-bound limits', () => {
  const rule: CellValidation = {
    kind: 'decimal',
    op: 'between',
    a: 0,
    b: 0,
    aFormula: 'A1',
    bFormula: 'B1',
  };
  const input = { kind: 'number', value: 5 } as const;
  it('reads single-cell limits through the resolver', () => {
    const cells: Record<string, string[]> = { A1: ['1'], B1: ['10'] };
    expect(validateAgainst(rule, input, (ref) => cells[ref] ?? []).ok).toBe(true);
    expect(validateAgainst(rule, { kind: 'number', value: 50 }, (ref) => cells[ref] ?? []).ok).toBe(
      false,
    );
  });
  it('defers to the engine when a limit cannot be evaluated', () => {
    expect(validateAgainst(rule, { kind: 'number', value: 50 }).ok).toBe(true);
  });
});
