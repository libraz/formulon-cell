import { describe, expect, it } from 'vitest';
import { formatChartNumber, formatChartValue } from '../../../src/charts/format.js';

describe('chart value formatting', () => {
  it('supports General and common grouped/percent formats', () => {
    expect(formatChartNumber(12345.6, undefined, 'en-US')).toBe('12,345.6');
    expect(formatChartNumber(12345.6, '#,##0.00', 'en-US')).toBe('12,345.60');
    expect(formatChartNumber(0.125, '0.0%', 'en-US')).toBe('12.5%');
    expect(formatChartNumber(12345.6, '#,##0.0', 'de-DE')).toBe('12.345,6');
  });

  it('keeps meaningful extreme General values and falls back for unknown codes', () => {
    expect(formatChartNumber(1e-20, undefined, 'en-US')).not.toBe('0');
    expect(formatChartNumber(1e20, undefined, 'en-US')).toContain('e');
    expect(formatChartNumber(1.2345, 'host-specific-code', 'en-US')).toBe('1.2345');
  });

  it('uses the injected formatter and falls back for nullish callback results', () => {
    expect(formatChartValue(3, 'x', 'ja-JP', () => 'Excel')).toBe('Excel');
    expect(formatChartValue(3, 'x', 'ja-JP', () => null)).toBe('3');
    expect(formatChartValue(3, 'x', 'ja-JP', () => undefined)).toBe('3');
  });
});
