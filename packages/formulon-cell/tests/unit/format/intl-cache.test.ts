import { afterEach, describe, expect, it, vi } from 'vitest';
import { IntlFormatterCache } from '../../../src/format/intl-cache.js';

describe('IntlFormatterCache', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('reuses canonical-locale number formatters when option order differs', () => {
    const cache = new IntlFormatterCache();
    const numberFormatSpy = vi.spyOn(Intl, 'NumberFormat');

    const first = cache.numberFormat('en_US', { useGrouping: false, maximumFractionDigits: 2 });
    const second = cache.numberFormat('en-US', { maximumFractionDigits: 2, useGrouping: false });

    expect(second).toBe(first);
    expect(numberFormatSpy).toHaveBeenCalledTimes(1);
  });

  it('keeps distinct locales, options, and formatter kinds separate', () => {
    const cache = new IntlFormatterCache();
    const numberFormat = vi.spyOn(Intl, 'NumberFormat');
    const dateTimeFormat = vi.spyOn(Intl, 'DateTimeFormat');

    const en = cache.numberFormat('en-US', { useGrouping: false });
    const de = cache.numberFormat('de-DE', { useGrouping: false });
    const grouped = cache.numberFormat('en-US', { useGrouping: true });
    const date = cache.dateTimeFormat('en-US', { month: 'short', timeZone: 'UTC' });

    expect(en.format(1234.5)).toBe('1234.5');
    expect(de.format(1234.5)).toBe('1234,5');
    expect(grouped.format(1234.5)).toBe('1,234.5');
    expect(date.format(new Date(Date.UTC(2024, 0, 1)))).toBe('Jan');
    expect(numberFormat).toHaveBeenCalledTimes(3);
    expect(dateTimeFormat).toHaveBeenCalledTimes(1);
  });

  it('evicts the least recently used entries at the configured bound', () => {
    const cache = new IntlFormatterCache(2);
    const numberFormatSpy = vi.spyOn(Intl, 'NumberFormat');

    const first = cache.numberFormat('en-US', { useGrouping: false });
    cache.numberFormat('de-DE', { useGrouping: false });
    cache.numberFormat('en-US', { useGrouping: false }); // Make the first entry most recent.
    cache.numberFormat('fr-FR', { useGrouping: false }); // Evict the de-DE entry.
    const firstAgain = cache.numberFormat('en-US', { useGrouping: false });
    cache.numberFormat('de-DE', { useGrouping: false });

    expect(firstAgain).toBe(first);
    expect(numberFormatSpy).toHaveBeenCalledTimes(4);
  });

  it('reuses formatters through the General number formatting entry point', async () => {
    vi.resetModules();
    const numberFormatSpy = vi.spyOn(Intl, 'NumberFormat');
    const { formatGeneralNumber } = await import('../../../src/engine/value.js');

    const outputs = Array.from({ length: 64 }, (_, index) =>
      formatGeneralNumber(1000 + index / 10, 'sv-SE'),
    );

    expect(outputs[0]).toBe('1 000');
    expect(outputs.at(-1)).toBe('1 006,3');
    expect(numberFormatSpy).toHaveBeenCalledTimes(1);
  });

  it('reuses fixed, currency, scientific, and date formatters through formatNumber', async () => {
    vi.resetModules();
    const numberFormat = vi.spyOn(Intl, 'NumberFormat');
    const dateTimeFormat = vi.spyOn(Intl, 'DateTimeFormat');
    const { formatNumber } = await import('../../../src/format/number-format.js');

    const fixed = { kind: 'fixed', decimals: 2, thousands: true } as const;
    const currency = { kind: 'currency', decimals: 2, symbol: '€' } as const;
    const scientific = { kind: 'scientific', decimals: 2 } as const;
    const date = { kind: 'custom', pattern: 'dddd, mmmm d, yyyy' } as const;

    expect(formatNumber(1234.5, fixed, 'de-DE')).toBe('1.234,50');
    expect(formatNumber(1234.5, fixed, 'de-DE')).toBe('1.234,50');
    expect(numberFormat).toHaveBeenCalledTimes(1);

    expect(formatNumber(1234.5, currency, 'de-DE')).toBe('1.234,50 €');
    expect(formatNumber(1234.5, currency, 'de-DE')).toBe('1.234,50 €');
    expect(numberFormat).toHaveBeenCalledTimes(2);

    expect(formatNumber(12345, scientific, 'de-DE')).toBe('1,23E+04');
    expect(formatNumber(12345, scientific, 'de-DE')).toBe('1,23E+04');
    expect(numberFormat).toHaveBeenCalledTimes(3);

    expect(formatNumber(45651, date, 'de-DE')).toBe('Mittwoch, Dezember 25, 2024');
    expect(formatNumber(45651, date, 'de-DE')).toBe('Mittwoch, Dezember 25, 2024');
    expect(dateTimeFormat).toHaveBeenCalledTimes(2);
    expect(numberFormat).toHaveBeenCalledTimes(3);
  });
});
