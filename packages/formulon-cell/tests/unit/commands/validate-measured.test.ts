import { describe, expect, it } from 'vitest';
import { cellValueViolatesValidation } from '../../../src/commands/validate.js';
import type { CellValue } from '../../../src/engine/types.js';
import type { CellValidation } from '../../../src/store/store.js';

// Excel Validation.Value per rule and value (dialogs-data-validation-results).
const num = (value: number): CellValue => ({ kind: 'number', value });
const text = (value: string): CellValue => ({ kind: 'text', value });
const blank: CellValue = { kind: 'blank' };
const error: CellValue = { kind: 'error', code: 1, text: '#DIV/0!' };

const whole = (over: Partial<CellValidation> = {}) =>
  ({ kind: 'whole', op: 'between', a: 1, b: 10, ...over }) as CellValidation;
const list = { kind: 'list', source: ['a', 'b', 'c'] } as CellValidation;
const rangeList = { kind: 'list', source: { ref: '$C$1:$C$3' } } as CellValidation;
const resolve = () => ['1', '2', '3'];

const CASES: [string, CellValidation, CellValue, boolean][] = [
  ['whole 1..10, 5', whole(), num(5), true],
  ['whole 1..10, 11', whole(), num(11), false],
  ['whole 1..10, 5.5', whole(), num(5.5), false],
  ['whole 1..10, text', whole(), text('abc'), false],
  ['whole 1..10, blank, ignore blank off', whole({ allowBlank: false }), blank, false],
  ['whole 1..10, blank, ignore blank on', whole({ allowBlank: true }), blank, true],
  ['whole 1..10, an error value', whole(), error, false],
  ['whole notBetween 1..10, 5', whole({ op: 'notBetween' }), num(5), false],
  ['whole <> 5, 5', whole({ op: '<>', a: 5 }), num(5), false],
  ['decimal > 5, 5.5', { kind: 'decimal', op: '>', a: 5 } as CellValidation, num(5.5), true],
  ['list a,b,c, b', list, text('b'), true],
  ['list a,b,c, B (case-sensitive)', list, text('B'), false],
  ['list a,b,c, d', list, text('d'), false],
  [
    'list a,b,c, empty text, ignore blank off',
    { ...list, allowBlank: false } as CellValidation,
    text(''),
    false,
  ],
  ['list =$C$1:$C$3, 2', rangeList, num(2), true],
  ['list =$C$1:$C$3, 4', rangeList, num(4), false],
  [
    'text length <= 3, abcd',
    { kind: 'textLength', op: '<=', a: 3 } as CellValidation,
    text('abcd'),
    false,
  ],
  [
    'text length <= 3, abc',
    { kind: 'textLength', op: '<=', a: 3 } as CellValidation,
    text('abc'),
    true,
  ],
  [
    'text length = 5, 12345',
    { kind: 'textLength', op: '=', a: 5 } as CellValidation,
    num(12345),
    true,
  ],
  [
    'date > 45000, 2024-01-05',
    { kind: 'date', op: '>', a: 45000 } as CellValidation,
    num(45296),
    true,
  ],
  ['time < 0.75, 0.5', { kind: 'time', op: '<', a: 0.75 } as CellValidation, num(0.5), true],
];

describe('cellValueViolatesValidation as Excel measures it', () => {
  it.each(CASES)('%s', (_name, validation, value, valid) => {
    expect(cellValueViolatesValidation(value, validation, resolve)).toBe(!valid);
  });
});
