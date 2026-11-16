import { describe, expect, it } from 'vitest';
import { dictionaries, type Locale } from '../../../src/i18n/strings.js';
import { buildRibbonSearchIndex } from '../../../src/toolbar/ribbon/search-index.js';

const helpItem = (input: Parameters<typeof buildRibbonSearchIndex>[0]) =>
  buildRibbonSearchIndex(input).find((item) => item.id === 'help:helpAndTraining');

describe('help search entry language', () => {
  it.each(Object.keys(dictionaries) as Locale[])('uses the %s strings', (locale) => {
    expect(helpItem(locale)?.label).toBe(dictionaries[locale].ribbon.helpSearch.label);
  });

  it('does not depend on the translated tab label', () => {
    const strings = structuredClone(dictionaries.en);
    strings.ribbon.tabs.help = 'ヘルプ';
    expect(helpItem(strings)?.label).toBe('Help and training');
  });
});
