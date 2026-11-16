import { describe, expect, it } from 'vitest';
import { dictionaries, type Locale } from '../../../../src/i18n/strings.js';
import { createQuickStyleStrip } from '../../../../src/toolbar/ribbon/quick-style-strip.js';

const stripFor = (lang: Locale): HTMLElement =>
  createQuickStyleStrip({
    lang,
    kind: 'table',
    mac: false,
    classPrefix: 'fc-test',
    chipKey: 'style',
    entries: () => [],
    current: () => null,
    apply: () => {},
  });

describe('quick style strip arrow labels', () => {
  it.each(Object.keys(dictionaries) as Locale[])('are localized for %s', (locale) => {
    const strip = stripFor(locale);
    const previous = strip.querySelector('.fc-test-prev')?.getAttribute('aria-label');
    const next = strip.querySelector('.fc-test-next')?.getAttribute('aria-label');
    expect(previous).toBe(dictionaries[locale].ribbon.quickStylesPrevious);
    expect(next).toBe(dictionaries[locale].ribbon.quickStylesNext);
    if (locale !== 'en') {
      expect(previous).not.toBe(dictionaries.en.ribbon.quickStylesPrevious);
      expect(next).not.toBe(dictionaries.en.ribbon.quickStylesNext);
    }
  });
});
