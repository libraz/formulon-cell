import { en } from './strings/en.js';
import { ja } from './strings/ja.js';
import type { Locale, Strings } from './strings/types.js';

export type { DeepPartial, Locale, Strings } from './strings/types.js';
export { en, ja };

export const dictionaries: Record<Locale, Strings> = { ja, en };
export const defaultStrings: Strings = ja;

export function mergeStrings(
  base: Strings,
  overlay?: import('./strings/types.js').DeepPartial<Strings>,
): Strings {
  if (!overlay) return base;
  const out = structuredClone(base) as Strings;
  for (const sectionKey of Object.keys(overlay) as (keyof Strings)[]) {
    const section = overlay[sectionKey];
    if (!section) continue;
    Object.assign(out[sectionKey], section);
  }
  return out;
}
