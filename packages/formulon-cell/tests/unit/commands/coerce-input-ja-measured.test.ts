import { describe, expect, it } from 'vitest';
import { coerceInput } from '../../../src/commands/coerce-input.js';
import { formatCodeToNumFmt } from '../../../src/engine/format-writeback.js';
import { excelSerialFromParts } from '../../../src/format/date-serial.js';

// Excel for Mac (ja UI), Range.Formula set on General cells.
// Serials for year-less dates use the
// capture year; `thisYear` re-bases them on the current one.
type Case = [
  input: string,
  kind: 'number' | 'text' | 'bool',
  value: number | string | boolean,
  format?: string,
];

const thisYear = (month: number, day: number): number => {
  const serial = excelSerialFromParts(new Date().getUTCFullYear(), month, day);
  if (serial === null) throw new Error('bad date');
  return serial;
};

const CASES: Case[] = [
  ['1/2', 'number', thisYear(1, 2), 'm"月"d"日"'],
  ['1-2', 'number', thisYear(1, 2), 'm"月"d"日"'],
  ['12%', 'number', 0.12, '0%'],
  ['12.5%', 'number', 0.125, '0.00%'],
  ['1e3', 'number', 1000, '0.00E+00'],
  ['1E+3', 'number', 1000, '0.00E+00'],
  ['0.1e2', 'number', 10, '0.00E+00'],
  ['001', 'number', 1],
  ['0123', 'number', 123],
  [' 1', 'number', 1],
  ['+5', 'number', 5],
  ['-5', 'number', -5],
  ['(5)', 'number', -5],
  ['.5', 'number', 0.5],
  ['5.', 'number', 5],
  ['１２３', 'number', 123],
  ['1,234', 'number', 1234, '#,##0'],
  ['1 1/2', 'number', 1.5, '# ?/?'],
  // Excel reads back $#,##0_);[Red]($#,##0); the stored code escapes the literals.
  ['$5', 'number', 5, '\\$#,##0_);[Red]\\(\\$#,##0\\)'],
  ['€5', 'number', 5],
  ['2024-1-5', 'number', 45296, 'mm-dd-yy'],
  ['2024/1/5', 'number', 45296, 'mm-dd-yy'],
  ['1/32', 'number', 11689, 'mmm-yy'],
  ['Jan 5', 'number', thisYear(1, 5), 'd-mmm'],
  ['5-Jan', 'number', thisYear(1, 5), 'd-mmm'],
  ['3:45', 'number', 0.15625, 'h:mm'],
  ['00:00', 'number', 0, 'h:mm'],
  ['10:30 PM', 'number', 0.9375, 'h:mm AM/PM'],
  ['25:00', 'number', 25 / 24, '[h]:mm:ss'],
  ['1:2:3', 'number', (3600 + 120 + 3) / 86400, 'h:mm:ss'],
  ['12345678901234567890', 'number', 12345678901234500000],
  ['TRUE', 'bool', true],
  ['true', 'bool', true],
  ['0x1', 'text', '0x1'],
  ['13/13', 'text', '13/13'],
  ['1,2', 'text', '1,2'],
  ['1.2.3', 'text', '1.2.3'],
  ['2024-13-45', 'text', '2024-13-45'],
  ['50%%', 'text', '50%%'],
  ['1e400', 'text', '1e400'],
  ['TRUE ', 'text', 'TRUE '],
  ['abc', 'text', 'abc'],
  ['1 000', 'text', '1 000'],
];

describe('coerceInput (ja-JP) as Excel measures it', () => {
  it.each(CASES)('%j', (input, kind, value, format) => {
    const got = coerceInput(input, { locale: 'ja-JP' });
    if (kind !== 'number') {
      expect(got).toEqual({ kind, value });
      return;
    }
    expect(got.kind).toBe('number');
    if (got.kind !== 'number') return;
    expect(got.value).toBeCloseTo(value as number, 9);
    expect(got.implicitFormat ?? null).toEqual(format ? formatCodeToNumFmt(format) : null);
  });

  it('keeps an error literal and a formula', () => {
    expect(coerceInput('#N/A', { locale: 'ja-JP' })).toMatchObject({ kind: 'error', text: '#N/A' });
    expect(coerceInput('=1/0', { locale: 'ja-JP' })).toEqual({ kind: 'formula', text: '=1/0' });
  });
});
