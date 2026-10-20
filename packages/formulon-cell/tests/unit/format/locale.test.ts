import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  currencySymbolForLocale,
  datePatternsForLocale,
  normalizeFormatLocale as fromShared,
} from '../../../src/format/locale.js';
import { formatNumber } from '../../../src/format/number-format.js';
import { normalizeFormatLocale as fromFormatDialog } from '../../../src/interact/format-dialog-model.js';
import { normalizeFormatLocale as fromHitState } from '../../../src/render/grid/hit-state.js';

describe('format/locale.ts — shared normalizer', () => {
  it('canonicalises short locales to BCP-47', () => {
    expect(fromShared('ja')).toBe('ja-JP');
    expect(fromShared('en')).toBe('en-US');
    expect(fromShared('en-GB')).toBe('en-GB');
    expect(fromShared('')).toBe('en-US');
  });

  it.each([
    ['de', 'de-DE', '€'],
    ['fr', 'fr-FR', '€'],
    ['zh', 'zh-CN', '¥'],
    ['ko', 'ko-KR', '₩'],
    ['th', 'th-TH', '฿'],
  ])('uses the format tag and currency for %s', (locale, tag, symbol) => {
    expect(fromShared(locale)).toBe(tag);
    expect(fromShared(tag?.replace('-', '_') ?? '')).toBe(tag);
    expect(currencySymbolForLocale(locale ?? '')).toBe(symbol);
  });

  it('preserves a regional currency and handles malformed locale input', () => {
    expect(currencySymbolForLocale('en-GB')).toBe('£');
    expect(fromShared('not_a_locale_!')).toBe('en-US');
    expect(datePatternsForLocale('de-DE').short).toBe('dd.mm.yyyy');
    expect(datePatternsForLocale('fr-FR').short).toBe('dd/mm/yyyy');
  });

  it.each(['en-US', 'ja-JP', 'de-DE', 'fr-FR', 'zh-CN', 'ko-KR', 'th-TH'])(
    'offers date patterns the engine accepts for %s',
    async (locale) => {
      const wb = await WorkbookHandle.createDefault();
      try {
        const patterns = datePatternsForLocale(locale);
        for (const pattern of [patterns.short, patterns.long]) {
          // An invalid code falls back to General, which prints the bare serial.
          expect(wb.formatValue({ kind: 'number', value: 45296 }, pattern)).toContain('2024');
        }
      } finally {
        wb.dispose();
      }
    },
  );

  it('formats custom and malformed locale ids without throwing', () => {
    for (const locale of ['de_@@', 'custom_brand']) {
      expect(formatNumber(1234.5, { kind: 'fixed', decimals: 2 }, locale)).toBe('1234.50');
      for (const kind of ['date', 'custom'] as const) {
        expect(() => formatNumber(45651, { kind, pattern: 'mmmm d, yyyy' }, locale)).not.toThrow();
      }
    }
  });

  it.each(['de-DE', 'fr-FR'])('places currency after the amount for %s', (locale) => {
    const number = new Intl.NumberFormat(locale, {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(1234.5);
    expect(formatNumber(1234.5, { kind: 'currency', decimals: 2, symbol: '€' }, locale)).toBe(
      `${number}\u00a0€`,
    );
    expect(
      formatNumber(
        -1234.5,
        { kind: 'currency', decimals: 2, symbol: '€', negativeStyle: 'parens' },
        locale,
      ),
    ).toBe(`(${number}\u00a0€)`);
  });

  it('hit-state re-exports the same implementation', () => {
    expect(fromHitState).toBe(fromShared);
  });

  it('format-dialog-model re-exports the same implementation', () => {
    expect(fromFormatDialog).toBe(fromShared);
  });
});
