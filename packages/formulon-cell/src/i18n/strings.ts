import { builtInLocaleFor, DEFAULT_LOCALE } from './locales.js';
import { de } from './strings/de.js';
import { en } from './strings/en.js';
import { fr } from './strings/fr.js';
import { ja } from './strings/ja.js';
import { ko } from './strings/ko.js';
import { th } from './strings/th.js';
import type { Locale, Strings } from './strings/types.js';
import { zh } from './strings/zh.js';

export { BUILT_IN_LOCALES } from './locales.js';
export type { DeepPartial, Locale, Strings } from './strings/types.js';
export { de, en, fr, ja, ko, th, zh };

export const dictionaries: Record<Locale, Strings> = { en, ja, de, fr, zh, ko, th };
export const defaultStrings: Strings = en;

export const dictionaryLocaleFor = (locale: string): Locale => {
  return builtInLocaleFor(locale) ?? DEFAULT_LOCALE;
};

const isPlainRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);

const applyOverlay = (target: Record<string, unknown>, overlay: Record<string, unknown>): void => {
  for (const [key, value] of Object.entries(overlay)) {
    if (value === undefined) continue;
    const current = target[key];
    if (isPlainRecord(current) && isPlainRecord(value)) {
      applyOverlay(current, value);
    } else {
      target[key] = value;
    }
  }
};

export function mergeStrings(
  base: Strings,
  overlay?: import('./strings/types.js').DeepPartial<Strings>,
): Strings {
  if (!overlay) return base;
  const out = structuredClone(base) as Strings;
  for (const sectionKey of Object.keys(overlay) as (keyof Strings)[]) {
    const section = overlay[sectionKey];
    if (!section) continue;
    const target = out[sectionKey] as unknown;
    if (isPlainRecord(target)) applyOverlay(target, section as Record<string, unknown>);
    else (out as unknown as Record<string, unknown>)[sectionKey] = structuredClone(section);
  }
  return out;
}
