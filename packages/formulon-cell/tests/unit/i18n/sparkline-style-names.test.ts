import { describe, expect, it } from 'vitest';
import { ribbonDe } from '../../../src/i18n/strings/ribbon/de.js';
import { ribbonEn } from '../../../src/i18n/strings/ribbon/en.js';
import { ribbonFr } from '../../../src/i18n/strings/ribbon/fr.js';
import { ribbonJa } from '../../../src/i18n/strings/ribbon/ja.js';
import { ribbonKo } from '../../../src/i18n/strings/ribbon/ko.js';
import { ribbonTh } from '../../../src/i18n/strings/ribbon/th.js';
import { ribbonZh } from '../../../src/i18n/strings/ribbon/zh.js';
import words from '../../fixtures/excel-color-words.json' with { type: 'json' };

const LOCALES = {
  en: ribbonEn,
  de: ribbonDe,
  fr: ribbonFr,
  ko: ribbonKo,
  th: ribbonTh,
  zh: ribbonZh,
} as const;

// The gallery chevron's accessible name in each locale, as the desktop app bundle ships it.
const MORE_ITEMS = {
  en: 'show more items',
  de: 'Weitere Elemente anzeigen',
  fr: 'afficher plus éléments',
  ko: '추가 항목 표시',
  th: 'แสดงรายการเพิ่มเติม',
  zh: '显示更多项目',
} as const;

// Each locale names a Style preset's colour with the desktop app's own word for the ja one.
describe('Sparkline Style names', () => {
  const ja = ribbonJa.ribbon.sparkline.styleColorNames;

  for (const [locale, strings] of Object.entries(LOCALES)) {
    it(`uses the desktop colour words and chevron name in ${locale}`, () => {
      const expected = ja.map(
        (word) => (words as Record<string, Record<string, string>>)[word]?.[locale],
      );
      expect(strings.ribbon.sparkline.styleColorNames).toEqual(expected);
      expect(strings.ribbon.sparkline.moreStyles).toBe(
        MORE_ITEMS[locale as keyof typeof MORE_ITEMS],
      );
    });
  }
});
