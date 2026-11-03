import { describe, expect, it } from 'vitest';
import { formatCodeToNumFmt, numFmtToFormatCode } from '../../../src/engine/format-writeback.js';
import type { NegativeStyle, NumFmt } from '../../../src/store/types.js';
import { numberFormatForAction } from '../../../src/toolbar/number-format.js';

// The code desktop Excel (ja) stores in styles.xml for each negative-number
// choice of the Number and Currency (yen) categories.
const NUMBER: [NegativeStyle, string][] = [
  ['red-parens', '0_);[Red]\\(0\\)'],
  ['parens', '0_);\\(0\\)'],
  ['red-plain', '0;[Red]0'],
  ['triangle', '0;"△ "0'],
  ['triangle-filled', '0;"▲ "0'],
];
const CURRENCY: [NegativeStyle, string][] = [
  ['red-parens', '"¥"#,##0_);[Red]\\("¥"#,##0\\)'],
  ['parens', '"¥"#,##0_);\\("¥"#,##0\\)'],
  ['red-plain', '"¥"#,##0;[Red]"¥"#,##0'],
  ['symbol-minus', '"¥"#,##0;"¥"\\-#,##0'],
  ['red-symbol-minus', '"¥"#,##0;[Red]"¥"\\-#,##0'],
];

describe('negative-number format codes match the saved Excel codes', () => {
  it.each(NUMBER)('number %s', (negativeStyle, code) => {
    const fmt: NumFmt = { kind: 'fixed', decimals: 0, negativeStyle };
    expect(numFmtToFormatCode(fmt)).toBe(code);
  });

  it.each(CURRENCY)('currency (yen) %s', (negativeStyle, code) => {
    const fmt: NumFmt = { kind: 'currency', decimals: 0, symbol: '¥', negativeStyle };
    expect(numFmtToFormatCode(fmt)).toBe(code);
    expect(formatCodeToNumFmt(code)).toMatchObject({ kind: 'currency', negativeStyle });
  });
});

describe('padded plain and red-minus number codes', () => {
  const cases: [NumFmt, string][] = [
    [{ kind: 'fixed', decimals: 0, trailingPad: true }, '0_ '],
    [{ kind: 'fixed', decimals: 0, negativeStyle: 'red', trailingPad: true }, '0_ ;[Red]\\-0\\ '],
    [{ kind: 'fixed', decimals: 2, trailingPad: true }, '0.00_ '],
    [
      { kind: 'fixed', decimals: 0, thousands: true, negativeStyle: 'red', trailingPad: true },
      '#,##0_ ;[Red]\\-#,##0\\ ',
    ],
    [{ kind: 'fixed', decimals: 0, thousands: true, trailingPad: true }, '#,##0_ '],
    [{ kind: 'fixed', decimals: 0, negativeStyle: 'parens', trailingPad: true }, '0_);\\(0\\)'],
  ];
  it.each(cases)('writes %j as %s', (fmt, code) => {
    expect(numFmtToFormatCode(fmt)).toBe(code);
  });

  it.each(['ja-JP', 'en-US'])('keeps the ribbon Number action unpadded for %s', (locale) => {
    expect(numFmtToFormatCode(numberFormatForAction('fixed', locale) ?? undefined)).toBe('0');
  });

  it('keeps the unpadded default codes', () => {
    expect(numFmtToFormatCode({ kind: 'fixed', decimals: 0 })).toBe('0');
    expect(numFmtToFormatCode({ kind: 'fixed', decimals: 0, negativeStyle: 'red' })).toBe(
      '0;[Red]-0',
    );
  });

  it.each([
    ['0_ ', { kind: 'fixed', decimals: 0, trailingPad: true }],
    ['0_ ;[Red]\\-0\\ ', { kind: 'fixed', decimals: 0, negativeStyle: 'red', trailingPad: true }],
    ['0_ ;[Red]-0 ', { kind: 'fixed', decimals: 0, negativeStyle: 'red', trailingPad: true }],
    ['#,##0.00_ ', { kind: 'fixed', decimals: 2, thousands: true, trailingPad: true }],
  ])('recognises %s as the Number category', (code, expected) => {
    expect(formatCodeToNumFmt(code)).toEqual(expected);
  });

  it('does not flag unpadded codes', () => {
    expect(formatCodeToNumFmt('0')).toEqual({ kind: 'fixed', decimals: 0 });
    expect(formatCodeToNumFmt('0;[Red]-0')).toEqual({
      kind: 'fixed',
      decimals: 0,
      negativeStyle: 'red',
    });
  });
});
