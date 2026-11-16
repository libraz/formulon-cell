import { describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import { macRibbonLabelForCommand } from '../../../src/toolbar/ribbon/mac/labels.js';
import { MAC_LOCALES, MAC_MENU_TRANSLATIONS } from '../../../src/toolbar/ribbon/mac/labels-data.js';

/** Menu items the app bundle has no string for: id, dictionary field, English wording. */
const DICTIONARY_ITEMS = [
  ['mac.draw.penBlack', 'penBlack', 'Pen: Black, 0.35 mm'],
  ['mac.draw.penRed', 'penRed', 'Pen: Red, 0.35 mm'],
  ['mac.draw.pencil', 'pencil', 'Pencil: Gray, 0.5 mm'],
  ['mac.draw.highlighter', 'highlighter', 'Highlighter: Yellow, 6 mm'],
  ['mac.data.validation.clearRules', 'clearValidationRules', 'Clear Validation'],
] as const;

/** Menu items whose bundle source competes or is unconfirmed; they keep their current text. */
const HELD_ITEMS = [
  'mac.page.margins.normal',
  'mac.page.size.legal',
  'mac.page.size.letter',
  'mac.review.editNote',
  'mac.view.freeze.off',
];

describe('Mac menu item label sources', () => {
  it.each(DICTIONARY_ITEMS)('%s reads the dictionaries in every locale', (id, field, en) => {
    expect(dictionaries.en.mac.ribbonMenuItems[field]).toBe(en);
    for (const locale of MAC_LOCALES) {
      expect(macRibbonLabelForCommand(id, locale), locale).toBe(
        dictionaries[locale].mac.ribbonMenuItems[field],
      );
    }
  });

  it('keeps hand-written translations only for held menu items', () => {
    expect(Object.keys(MAC_MENU_TRANSLATIONS).sort()).toEqual([...HELD_ITEMS].sort());
  });
});
