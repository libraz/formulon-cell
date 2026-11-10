// Observable strings controller — hands out the current locale dictionary
// and notifies subscribers when it changes. Each `attach*` function reads
// labels through this controller and re-renders on change instead of
// snapshotting strings at attach time.
import { DEFAULT_LOCALE, type Locale } from './locales.js';
import {
  type DeepPartial,
  defaultStrings,
  dictionaries,
  dictionaryLocaleFor,
  mergeStrings,
  type Strings,
} from './strings.js';

export interface I18nControllerInit {
  /** Locale to start with. Defaults to `DEFAULT_LOCALE` (matches `defaultStrings`). */
  locale?: Locale;
  /** Initial overlay (deep-merged on top of the chosen locale). */
  overlay?: DeepPartial<Strings>;
}

export interface I18nController {
  readonly locale: string;
  readonly strings: Strings;
  setLocale(locale: string): void;
  /** Deep-merge an overlay onto a locale's strings. Later `setLocale(locale)`
   *  calls return the merged result. If `locale` is the current one,
   *  listeners fire immediately. */
  override(locale: string, overlay: DeepPartial<Strings>): void;
  /** Listen for strings changes — fires on setLocale + override. The callback
   *  receives the freshly resolved strings; consumers update their DOM.
   *  Returns an idempotent disposer. */
  on(name: 'change', fn: (s: Strings) => void): () => void;
  dispose(): void;
}

export const createI18nController = (init: I18nControllerInit = {}): I18nController => {
  const registry = new Map<string, Strings>();
  for (const [k, v] of Object.entries(dictionaries)) registry.set(k, v);

  const overlays = new Map<string, DeepPartial<Strings>>();
  let current = init.locale ?? DEFAULT_LOCALE;
  if (init.overlay) overlays.set(current, init.overlay);

  const listeners = new Set<(s: Strings) => void>();

  const resolve = (): Strings => {
    const base =
      registry.get(current) ?? registry.get(dictionaryLocaleFor(current)) ?? defaultStrings;
    const overlay = overlays.get(current);
    return overlay ? mergeStrings(base, overlay) : base;
  };

  let cached = resolve();
  const recompute = (): void => {
    cached = resolve();
    for (const fn of listeners) fn(cached);
  };

  const overrideLocale = (locale: string, overlay: DeepPartial<Strings>): void => {
    const prev = overlays.get(locale);
    overlays.set(locale, prev ? mergeStrings(prev as Strings, overlay) : overlay);
    if (locale === current) recompute();
  };
  const listen = (fn: (s: Strings) => void): (() => void) => {
    listeners.add(fn);
    return () => {
      listeners.delete(fn);
    };
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
    override: overrideLocale,
    on: (_name, fn) => listen(fn),
    dispose() {
      listeners.clear();
      registry.clear();
      overlays.clear();
    },
  };
};
