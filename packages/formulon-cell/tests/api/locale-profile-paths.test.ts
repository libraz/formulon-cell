import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, expect, it } from 'vitest';
import { workbookHandleOf } from '../../src/api/workbook.js';
import { makeHost } from './helpers.js';

const live: SpreadsheetInstance[] = [];
afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
});

const mountSheet = async (options: Parameters<typeof mount>[1]): Promise<SpreadsheetInstance> => {
  const sheet = await mount(makeHost(), options);
  live.push(sheet);
  return sheet;
};

const profileOf = (sheet: SpreadsheetInstance): string =>
  workbookHandleOf(sheet.workbook).engineWorkbook().getProfileId();

// Bare languages, regions, scripts and private extensions of a built-in language.
const TAGS = [
  'de',
  'de-DE',
  'de-Latn-DE',
  'de-brand',
  'ja-JP',
  'ja-Jpan-JP',
  'fr-CA',
  'zh-TW',
  'zh-Hant',
  'zh-Hans-HK',
  'en-GB',
] as const;
const PLATFORMS = ['default', 'mac'] as const;
const ROWS = PLATFORMS.flatMap((platform) => TAGS.map((tag) => [platform, tag] as const));

it.each(ROWS)(
  'the %s platform gives %s the same profile at mount and at update',
  async (platform, tag) => {
    const mounted = await mountSheet({ locale: tag, ui: { platform } });
    const updated = await mountSheet({ ui: { platform } });
    updated.update({ locale: tag });
    expect(profileOf(updated)).toBe(profileOf(mounted));
  },
);

it.each(PLATFORMS)(
  'the %s platform keeps the profile for a language without one',
  async (platform) => {
    const fresh = await mountSheet({ ui: { platform } });
    const english = profileOf(fresh);
    const mounted = await mountSheet({ locale: 'xx-brand', ui: { platform } });
    expect(profileOf(mounted)).toBe(english);

    const updated = await mountSheet({ locale: 'ja', ui: { platform } });
    const japanese = profileOf(updated);
    updated.i18n.override('xx-brand', {});
    updated.update({ locale: 'xx-brand' });
    expect(profileOf(updated)).toBe(japanese);
  },
);
