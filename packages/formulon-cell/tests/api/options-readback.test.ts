import { mount, type SpreadsheetInstance, type SpreadsheetPatch } from '@libraz/formulon-cell';
import { afterEach, expect, it } from 'vitest';
import { makeHost } from './helpers.js';

const live: SpreadsheetInstance[] = [];
afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
});

// `workbook` is left out: a null workbook is refused, not reset.
const RESETTABLE_KEYS = [
  'ui',
  'policy',
  'viewport',
  'contextMenu',
  'overlays',
  'toolbar',
  'theme',
  'locale',
  'strings',
  'extensions',
  'formulaHints',
  'spelling',
  'autoCorrect',
  'customLists',
  'host',
  'uploadStatus',
  'macroRecording',
  'showAllNotes',
] as const satisfies readonly (keyof SpreadsheetPatch)[];

type Uncovered = Exclude<keyof SpreadsheetPatch, 'workbook' | (typeof RESETTABLE_KEYS)[number]>;
const everyKeyCovered: [Uncovered] extends [never] ? true : Uncovered = true;

it.each(RESETTABLE_KEYS)(
  'options.%s is absent after a bare mount and after a null reset',
  async (key) => {
    expect(everyKeyCovered).toBe(true);
    const sheet = await mount(makeHost(), {});
    live.push(sheet);
    expect(key in sheet.options).toBe(false);
    sheet.update({ [key]: null });
    expect(key in sheet.options).toBe(false);
  },
);

it('reads the effective locale from i18n when options leave it out', async () => {
  const sheet = await mount(makeHost(), {});
  live.push(sheet);
  expect(sheet.i18n.locale).toBe('en');
  sheet.update({ locale: 'ja' });
  sheet.update({ locale: null });
  expect(sheet.i18n.locale).toBe('en');
});
