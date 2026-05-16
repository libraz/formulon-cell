import { describe, expect, it } from 'vitest';
import { numberFormatForAction } from '../../../src/toolbar/number-format.js';

describe('toolbar/number-format', () => {
  it('maps toolbar number-format actions to locale-aware formats', () => {
    expect(numberFormatForAction('currency', 'ja')).toEqual({
      kind: 'currency',
      decimals: 2,
      symbol: '¥',
    });
    expect(numberFormatForAction('accounting', 'en')).toEqual({
      kind: 'accounting',
      decimals: 2,
      symbol: '$',
    });
    expect(numberFormatForAction('shortDate', 'ja')).toEqual({
      kind: 'date',
      pattern: 'yyyy/m/d',
    });
    expect(numberFormatForAction('longDate', 'en')).toEqual({
      kind: 'date',
      pattern: 'mmmm d, yyyy',
    });
    expect(numberFormatForAction('time', 'ja')).toEqual({ kind: 'time', pattern: 'H:MM' });
    expect(numberFormatForAction('more', 'en')).toBeNull();
  });
});
