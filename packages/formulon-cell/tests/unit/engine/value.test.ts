import { describe, expect, it } from 'vitest';
import {
  formatBoolean,
  formatCell,
  formatEditableNumber,
  formatError,
  formatGeneralNumber,
  formatScientificNumber,
} from '../../../src/engine/value.js';

describe('formatGeneralNumber', () => {
  it('uses locale grouping for ordinary General numbers', () => {
    expect(formatGeneralNumber(1234.5, 'en-US')).toBe('1,234.5');
  });

  it('can omit grouping while keeping General scientific notation', () => {
    expect(formatGeneralNumber(1234.5, 'en-US', { useGrouping: false })).toBe('1234.5');
    expect(formatGeneralNumber(123456789012, 'en-US', { useGrouping: false })).toBe('1.23457E+11');
    expect(formatGeneralNumber(-0.0000000001234, 'en-US', { useGrouping: false })).toBe(
      '-1.23400E-10',
    );
  });

  it('uses spreadsheet-like scientific notation for very large or tiny numbers', () => {
    expect(formatGeneralNumber(123456789012)).toBe('1.23457E+11');
    expect(formatGeneralNumber(0.0000000001234)).toBe('1.23400E-10');
  });

  it.each([
    ['en-US', '.'],
    ['ja-JP', '.'],
    ['de-DE', ','],
    ['fr-FR', ','],
    ['zh-CN', '.'],
    ['ko-KR', '.'],
    ['th-TH', '.'],
  ] as const)('uses the profile decimal mark for scientific output in %s', (locale, decimal) => {
    expect(formatScientificNumber(-12345.678, 2, locale)).toBe(`-1${decimal}23E+04`);
    expect(formatScientificNumber(-0.00012, 2, locale)).toBe(`-1${decimal}20E-04`);
    expect(formatScientificNumber(1234567890123, 2, locale)).toBe(`1${decimal}23E+12`);
    expect(formatScientificNumber(1.9996, 4, locale)).toBe(`1${decimal}9996E+00`);
  });

  it('allows custom exponent padding while keeping scientific rounding', () => {
    expect(formatScientificNumber(12345, 2, 'en-US', 1)).toBe('1.23E+4');
    expect(formatScientificNumber(12345, 2, 'de-DE', 3)).toBe('1,23E+004');
  });

  it('is used by formatCell for number values', () => {
    expect(formatCell({ kind: 'number', value: 123456789012 })).toBe('1.23457E+11');
  });

  it('localizes the measured German #N/A error and preserves unknown fallbacks', () => {
    expect(formatError(6, '#N/A', 'de-DE')).toBe('#NV');
    expect(formatCell({ kind: 'error', code: 6, text: '#N/A' }, 'de-DE')).toBe('#NV');
    expect(formatError(6, '#N/A', 'fr-FR')).toBe('#N/A');
    expect(formatError(10, '#ERR!', 'de-DE')).toBe('#FIELD!');
    expect(formatError(99, '#CUSTOM!', 'de-DE')).toBe('#CUSTOM!');
  });

  it('keeps editable decimal precision while using German/French decimal marks', () => {
    expect(formatEditableNumber(1.5, 'de-DE')).toBe('1,5');
    expect(formatEditableNumber(1.5, 'fr-FR')).toBe('1,5');
    expect(formatEditableNumber(1.5, 'en-US')).toBe('1.5');
    expect(formatEditableNumber(1.23e-7, 'de-DE')).toBe('1,23e-7');
  });
});

describe('formatBoolean', () => {
  it('uses localized German/French tokens and canonical defaults elsewhere', () => {
    expect(formatBoolean(true, 'de-DE')).toBe('WAHR');
    expect(formatBoolean(false, 'de-DE')).toBe('FALSCH');
    expect(formatBoolean(true, 'fr-FR')).toBe('VRAI');
    expect(formatBoolean(false, 'fr-FR')).toBe('FAUX');
    expect(formatBoolean(true, 'ja-JP')).toBe('TRUE');
    expect(formatBoolean(false, 'en-US')).toBe('FALSE');
    expect(formatCell({ kind: 'bool', value: true }, 'de-DE')).toBe('WAHR');
  });
});
