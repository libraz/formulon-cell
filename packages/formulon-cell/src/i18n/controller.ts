// Observable strings controller — hands out the current locale dictionary
// and notifies subscribers when it changes. Each `attach*` function reads
// labels through this controller and re-renders on change instead of
// snapshotting strings at attach time.
//
// Inspired by Handsontable's `registerLanguageDictionary` API: locales can
// be added at runtime, then activated by id.
import {
  type DeepPartial,
  defaultStrings,
  dictionaries,
  type Locale,
  mergeStrings,
  type Strings,
} from './strings.js';

export interface I18nControllerInit {
  /** Locale to start with. Defaults to 'ja' (matches `defaultStrings`). */
  locale?: Locale | (string & {});
  /** Initial overlay (deep-merged on top of the chosen locale). */
  overlay?: DeepPartial<Strings>;
}

export interface I18nController {
  readonly locale: string;
  readonly strings: Strings;
  setLocale(locale: string): void;
  /** Deep-merge an overlay onto a registered locale's strings. Future
   *  `setLocale(locale)` calls return the merged result. If `locale` is the
   *  current one, subscribers fire immediately. */
  extend(locale: string, overlay: DeepPartial<Strings>): void;
  /** Register a brand-new locale's full dictionary at runtime. */
  register(locale: string, strings: Strings): void;
  /** Listen for strings changes — fires on setLocale + extend. The callback
   *  receives the freshly resolved strings; consumers update their DOM. */
  subscribe(fn: (s: Strings) => void): () => void;
  dispose(): void;
}

export const createI18nController = (init: I18nControllerInit = {}): I18nController => {
  const registry = new Map<string, Strings>();
  for (const [k, v] of Object.entries(dictionaries)) registry.set(k, v);

  const overlays = new Map<string, DeepPartial<Strings>>();
  let current = init.locale ?? 'ja';
  if (init.overlay) overlays.set(current, init.overlay);

  const listeners = new Set<(s: Strings) => void>();

  const resolve = (): Strings => {
    const base = registry.get(current) ?? defaultStrings;
    const overlay = overlays.get(current);
    return overlay ? mergeStrings(base, overlay) : base;
  };

  let cached = resolve();
  const recompute = (): void => {
    cached = resolve();
    for (const fn of listeners) fn(cached);
  };

  return {
    get locale() {
      return current;
    },
    get strings() {
      return cached;
    },
    setLocale(locale) {
      if (locale === current) return;
      current = locale;
      recompute();
    },
    extend(locale, overlay) {
      const prev = overlays.get(locale);
      overlays.set(locale, prev ? mergeStrings(prev as Strings, overlay) : overlay);
      if (locale === current) recompute();
    },
    register(locale, strings) {
      registry.set(locale, strings);
      if (locale === current) recompute();
    },
    subscribe(fn) {
      listeners.add(fn);
      return () => listeners.delete(fn);
    },
    dispose() {
      listeners.clear();
      registry.clear();
      overlays.clear();
    },
  };
};
