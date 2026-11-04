import { describe, expect, it } from 'vitest';
import { numberFormatForAction, numberFormatSample } from '../../../src/toolbar/number-format.js';

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

  it('samples each dropdown row the way Excel shows it for 1234.5', () => {
    const sample = (action: Parameters<typeof numberFormatSample>[0]) =>
      numberFormatSample(action, 1234.5, 'en');
    expect(sample('fixed')).toBe('1234.50');
    expect(sample('percent')).toBe('123450.00%');
    expect(sample('scientific')).toBe('1.23E+03');
    expect(sample('fraction')).toBe('1234 1/2');
    expect(sample('text')).toBe('1234.5');
    expect(sample('general')).toBeNull();
    expect(sample('more')).toBeNull();
  });
});
