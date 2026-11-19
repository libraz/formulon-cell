import { describe, expect, it } from 'vitest';

import { dictionaries } from '../../../src/i18n/strings.js';

/**
 * The desktop app draws a trailing ellipsis in menu items as three dots with no
 * preceding space. A few items are drawn with the single glyph; each is listed
 * here with the captured control backing it.
 */
const ALLOW: Record<string, string> = {
  'de:fillGroup': 'Home > Fill menu item 0.4 (Across Worksheets), drawn "Across Arbeitsblätter …"',
  'th:fillGroup': 'Home > Fill menu item 0.4 (Across Worksheets), drawn "ข้ามเวิร์กชีต…"',
  'zh:goTo': 'Home > Find & Select menu item 0.2 (Go To), drawn "转到…"',
  'zh:pictureStock': 'Insert > Pictures > Stock Images (both placements), drawn "图像集…"',
};

const LOCALES = ['en', 'ja', 'de', 'fr', 'ko', 'th', 'zh'] as const;

const menuStrings = (locale: (typeof LOCALES)[number]): [string, string][] =>
  Object.entries(dictionaries[locale].ribbonMenu).filter(
    (entry): entry is [string, string] => typeof entry[1] === 'string',
  );

describe('ribbon menu trailing ellipsis', () => {
  for (const locale of LOCALES) {
    it(`${locale}: uses three dots unless the item is drawn with the glyph`, () => {
      const offenders = menuStrings(locale)
        .filter(([key, value]) => /(…| \.\.\.)$/.test(value) && !(`${locale}:${key}` in ALLOW))
        .map(([key, value]) => `${key}: ${value}`);
      expect(offenders).toEqual([]);
    });
  }

  it('every allow entry still ends in the glyph', () => {
    const stale = Object.keys(ALLOW).filter((id) => {
      const [locale, key] = id.split(':') as [(typeof LOCALES)[number], string];
      const value = (dictionaries[locale].ribbonMenu as Record<string, unknown>)[key];
      return typeof value !== 'string' || !value.endsWith('…');
    });
    expect(stale).toEqual([]);
  });
});
