import { describe, expect, it } from 'vitest';
import { formatCodeToNumFmt, numFmtToFormatCode } from '../../../src/engine/format-writeback.js';
import type { NegativeStyle, NumFmt } from '../../../src/store/types.js';

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
