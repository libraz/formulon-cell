import { en } from './strings/en.js';
import { ja } from './strings/ja.js';
import type { Locale, Strings } from './strings/types.js';

export type { DeepPartial, Locale, Strings } from './strings/types.js';
export { en, ja };

export const dictionaries: Record<Locale, Strings> = { ja, en };
export const defaultStrings: Strings = ja;

export const dictionaryLocaleFor = (locale: string): Locale => {
  const normalized = locale.trim().toLowerCase().replace('_', '-');
  if (normalized === 'en' || normalized.startsWith('en-')) return 'en';
  if (normalized === 'ja' || normalized.startsWith('ja-')) return 'ja';
  return 'ja';
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
