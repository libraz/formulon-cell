import { describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import { labels } from '../../../src/toolbar/ribbon/mac/labels.js';
import {
  MAC_LABEL_TRANSLATIONS,
  MAC_LOCALES,
} from '../../../src/toolbar/ribbon/mac/labels-data.js';

/** Ribbon labels the app bundle has no string for: label key, dictionary field, English wording. */
const DICTIONARY_LABELS = [['trackpad', 'trackpad', 'Draw with Trackpad']] as const;

/** Labels whose bundle source competes or is unconfirmed; they keep their current text. */
const HELD_LABELS = ['penBlack', 'penRed', 'shapeLine', 'shapeOval', 'shapeDiamond'];

describe('Mac ribbon label sources', () => {
  it.each(DICTIONARY_LABELS)('%s reads the dictionaries in every locale', (key, field, en) => {
    expect(dictionaries.en.mac.ribbonLabels[field]).toBe(en);
    for (const locale of MAC_LOCALES) {
      expect(labels[key]?.[locale], locale).toBe(dictionaries[locale].mac.ribbonLabels[field]);
    }
  });

  it('keeps hand-written translations only for held labels', () => {
    expect(Object.keys(MAC_LABEL_TRANSLATIONS).sort()).toEqual([...HELD_LABELS].sort());
  });
});
