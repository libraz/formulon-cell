import { describe, expect, it } from 'vitest';
import { formatCell, formatGeneralNumber } from '../../../src/engine/value.js';

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

  it('is used by formatCell for number values', () => {
    expect(formatCell({ kind: 'number', value: 123456789012 })).toBe('1.23457E+11');
  });
});
