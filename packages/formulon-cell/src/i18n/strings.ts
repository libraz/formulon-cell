import { builtInLocaleFor } from './locales.js';
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

export const dictionaries: Record<Locale, Strings> = { ja, en, de, fr, zh, ko, th };
export const defaultStrings: Strings = ja;

export const dictionaryLocaleFor = (locale: string): Locale => {
  return builtInLocaleFor(locale) ?? 'ja';
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
    applyOverlay(out[sectionKey] as Record<string, unknown>, section as Record<string, unknown>);
  }
  return out;
}
