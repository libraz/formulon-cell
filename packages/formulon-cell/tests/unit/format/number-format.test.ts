import { describe, expect, it } from 'vitest';
import { formatNumber } from '../../../src/format/number-format.js';

describe('formatNumber', () => {
  it('returns the input as a string for non-finite numbers', () => {
    expect(formatNumber(Number.NaN, undefined)).toBe('NaN');
    expect(formatNumber(Number.POSITIVE_INFINITY, undefined)).toBe('Infinity');
  });

  it('uses general formatting when no format is supplied', () => {
    expect(formatNumber(1234.5, undefined)).toBe('1,234.5');
  });

  it('applies fixed decimals', () => {
    expect(formatNumber(1.5, { kind: 'fixed', decimals: 3 })).toBe('1.500');
    expect(formatNumber(1.2345, { kind: 'fixed', decimals: 2 })).toBe('1.23');
    expect(formatNumber(1234.5, { kind: 'fixed', decimals: 2, thousands: true })).toBe('1,234.50');
  });

  it('solves fraction formats instead of printing ?/? verbatim (H-3)', () => {
    expect(formatNumber(1.5, { kind: 'custom', pattern: '# ?/?' })).toBe('1 1/2');
    expect(formatNumber(0.5, { kind: 'custom', pattern: '# ?/?' })).toBe('1/2');
    expect(formatNumber(2.25, { kind: 'custom', pattern: '# ??/??' })).toBe('2 1/4');
    // Fixed denominator: 0.625 in eighths → 5/8.
    expect(formatNumber(0.625, { kind: 'custom', pattern: '# ?/8' })).toBe('5/8');
    // Whole number → no fraction shown.
    expect(formatNumber(3, { kind: 'custom', pattern: '# ?/?' })).toBe('3');
    // Negative keeps its sign.
    expect(formatNumber(-1.5, { kind: 'custom', pattern: '# ?/?' })).toBe('-1 1/2');
  });

  it('zero-pads the scientific exponent to two digits (H-4)', () => {
    expect(formatNumber(12345, { kind: 'scientific', decimals: 2 })).toBe('1.23E+04');
    expect(formatNumber(0.00012, { kind: 'scientific', decimals: 2 })).toBe('1.20E-04');
    // Already ≥ 2 exponent digits stays unchanged.
    expect(formatNumber(1.23e15, { kind: 'scientific', decimals: 2 })).toBe('1.23E+15');
  });

  it('renders elapsed-time formats without wrapping at 24h/60 (H-5)', () => {
    // 1.5 days = 36 hours. [h] must show 36, not 12.
    expect(formatNumber(1.5, { kind: 'custom', pattern: '[h]:mm:ss' })).toBe('36:00:00');
    // 90 minutes = 0.0625 days → [m] = 90, [mm]:ss stays within its unit.
    expect(formatNumber(90 / 1440, { kind: 'custom', pattern: '[m]:ss' })).toBe('90:00');
    // 1h30m45s → total seconds for [s].
    const serial = (1 * 3600 + 30 * 60 + 45) / 86400;
    expect(formatNumber(serial, { kind: 'custom', pattern: '[s]' })).toBe('5445');
    // [hh] zero-pads to at least two digits.
    expect(formatNumber(5 / 24, { kind: 'custom', pattern: '[hh]:mm' })).toBe('05:00');
  });

  it('renders five-m month tokens as a single month initial', () => {
    expect(formatNumber(45651, { kind: 'custom', pattern: 'mmmm' }, 'en-US')).toBe('December');
    expect(formatNumber(45651, { kind: 'custom', pattern: 'mmmmm' }, 'en-US')).toBe('D');
  });

  it('matches Excel 1900 date serials before the leap-year bug boundary', () => {
    expect(formatNumber(1, { kind: 'custom', pattern: 'yyyy-mm-dd' })).toBe('1900-01-01');
    expect(formatNumber(59, { kind: 'custom', pattern: 'yyyy-mm-dd' })).toBe('1900-02-28');
    expect(formatNumber(61, { kind: 'custom', pattern: 'yyyy-mm-dd' })).toBe('1900-03-01');
  });

  it('renders scientific and engineering notation instead of E+0 verbatim (H-6)', () => {
    // Engineering: 3 integer placeholders → exponent stepped in multiples of 3.
    expect(formatNumber(12345, { kind: 'custom', pattern: '##0.0E+0' })).toBe('12.3E+3');
    expect(formatNumber(0.00012345, { kind: 'custom', pattern: '##0.0E+0' })).toBe('123.5E-6');
    // Plain scientific custom pattern.
    expect(formatNumber(12345, { kind: 'custom', pattern: '0.00E+00' })).toBe('1.23E+04');
    // Negative keeps its sign; exponent sign spec `E-` omits the plus.
    expect(formatNumber(-12345, { kind: 'custom', pattern: '0.0E-0' })).toBe('-1.2E4');
  });

  it('distributes digits across multi-run integer patterns (M-1)', () => {
    // Phone number: digits spread right-to-left across the placeholder runs.
    expect(formatNumber(1234567890, { kind: 'custom', pattern: '000-000-0000' })).toBe(
      '123-456-7890',
    );
    // SSN-style grouping.
    expect(formatNumber(123456789, { kind: 'custom', pattern: '000-00-0000' })).toBe('123-45-6789');
    // A plain grouped pattern is NOT treated as multi-run.
    expect(formatNumber(1234567, { kind: 'custom', pattern: '#,##0' })).toBe('1,234,567');
  });

  it('renders currency with prefix symbol and respects negatives', () => {
    expect(formatNumber(99, { kind: 'currency', decimals: 2, symbol: '$' })).toBe('$99.00');
    expect(formatNumber(-99, { kind: 'currency', decimals: 0, symbol: '€' })).toBe('-€99');
  });

  it('renders Excel-style negative number variants', () => {
    expect(formatNumber(-99, { kind: 'fixed', decimals: 0, negativeStyle: 'parens' })).toBe('(99)');
    expect(formatNumber(-99, { kind: 'fixed', decimals: 0, negativeStyle: 'red' })).toBe('-99');
    expect(
      formatNumber(-99, {
        kind: 'currency',
        decimals: 0,
        symbol: '$',
        negativeStyle: 'red-parens',
      }),
    ).toBe('($99)');
  });

  it('falls back to "$" when currency symbol is missing', () => {
    expect(formatNumber(7.5, { kind: 'currency', decimals: 1 })).toBe('$7.5');
  });

  it('renders percent values', () => {
    expect(formatNumber(0.25, { kind: 'percent', decimals: 0 })).toBe('25%');
    expect(formatNumber(0.1234, { kind: 'percent', decimals: 2 })).toBe('12.34%');
  });

  it('does not scale for percent signs inside quoted custom literals', () => {
    expect(formatNumber(0.25, { kind: 'custom', pattern: '0%' })).toBe('25%');
    expect(formatNumber(0.25, { kind: 'custom', pattern: '0"%"' })).toBe('0%');
  });

  it('renders spreadsheet locale-tagged currency custom formats', () => {
    const fmt = { kind: 'custom' as const, pattern: '[$¥-411]#,##0;[Red]-[$¥-411]#,##0' };
    expect(formatNumber(1234, fmt, 'ja-JP')).toBe('¥1,234');
    expect(formatNumber(-1234, fmt, 'ja-JP')).toBe('-¥1,234');
  });

  it('renders Special category masks without losing the category semantics', () => {
    expect(formatNumber(12345, { kind: 'special', pattern: '00000' })).toBe('12345');
    expect(formatNumber(123456789, { kind: 'special', pattern: '00000-0000' })).toBe('12345-6789');
    expect(formatNumber(123456789, { kind: 'special', pattern: '000-00-0000' })).toBe(
      '123-45-6789',
    );
  });

  it('hides spreadsheet accounting spacing and fill directives in custom output', () => {
    const fmt = {
      kind: 'custom' as const,
      pattern: '_-"$"* #,##0.00_-;-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-',
    };
    expect(formatNumber(1234, fmt)).toBe('$1,234.00');
    expect(formatNumber(-1234, fmt)).toBe('-$1,234.00');
  });
});
