import { describe, expect, it } from 'vitest';
import { de, en, fr, ja, ko, th, zh } from '../../../src/i18n/strings.js';

/** Every string of a dictionary with its dotted path. */
const leaves = (value: unknown, path: string, out: [string, string][] = []): [string, string][] => {
  if (typeof value === 'string') out.push([path, value]);
  else if (value && typeof value === 'object') {
    for (const [key, child] of Object.entries(value)) leaves(child, `${path}.${key}`, out);
  }
  return out;
};

// Wording taken from the desktop app drops the product's name, so the UI never names it.
describe('locale strings', () => {
  it('never name the desktop product', () => {
    for (const [locale, strings] of Object.entries({ de, en, fr, ja, ko, th, zh })) {
      const named = leaves(strings, locale).filter(([, text]) => /Excel/.test(text));
      expect(named).toEqual([]);
    }
  });
});
