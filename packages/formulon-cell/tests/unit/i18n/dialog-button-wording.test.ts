import { describe, expect, it } from 'vitest';
import { de } from '../../../src/i18n/strings/de.js';
import { en } from '../../../src/i18n/strings/en.js';
import { fr } from '../../../src/i18n/strings/fr.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { ko } from '../../../src/i18n/strings/ko.js';
import { th } from '../../../src/i18n/strings/th.js';
import { zh } from '../../../src/i18n/strings/zh.js';

/**
 * `<locale>:<key path>` allowed to word OK / Cancel differently from the rest, mapped to the reason.
 * The desktop wording bundle reads the same OK / Cancel in every dialog window it has, so an entry
 * needs a per-dialog source.
 */
const ALLOWED_DIFFERENCES: Readonly<Record<string, string>> = {
  'ja:mac.data.consolidate.cancel':
    'Set to the Mac wording with the data tool dialogs; the bundle has no Consolidate window.',
};

const LOCALES = { en, ja, de, fr, ko, th, zh } as const;

/** Every `<path>.ok` / `<path>.cancel` string, found by walking the locale object. */
function buttonStrings(obj: unknown, prefix = ''): [string, string][] {
  if (obj === null || typeof obj !== 'object') return [];
  const out: [string, string][] = [];
  for (const [key, value] of Object.entries(obj as Record<string, unknown>)) {
    const path = prefix ? `${prefix}.${key}` : key;
    if (typeof value === 'string') {
      if (key === 'ok' || key === 'cancel') out.push([path, value]);
    } else out.push(...buttonStrings(value, path));
  }
  return out;
}

describe('i18n/strings — dialog OK / Cancel wording', () => {
  for (const [locale, strings] of Object.entries(LOCALES)) {
    for (const button of ['ok', 'cancel'] as const) {
      it(`${locale}: every dialog words ${button} alike`, () => {
        const entries = buttonStrings(strings).filter(
          ([path]) => path.endsWith(`.${button}`) && !(`${locale}:${path}` in ALLOWED_DIFFERENCES),
        );
        const counts = new Map<string, number>();
        for (const [, value] of entries) counts.set(value, (counts.get(value) ?? 0) + 1);
        const common = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
        const differing = entries.filter(([, value]) => value !== common);
        expect(
          differing.map(([path, value]) => `${path}=${JSON.stringify(value)} (common ${common})`),
        ).toEqual([]);
      });
    }
  }
});
