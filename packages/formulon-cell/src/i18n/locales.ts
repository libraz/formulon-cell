/** UI locales with a built-in dictionary and matching formulon profile. */
export const BUILT_IN_LOCALES = ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'] as const;

export type BuiltInLocale = (typeof BUILT_IN_LOCALES)[number];
/**
 * A UI locale: a built-in id (`'en'`, `'ja'`, …) or any other BCP 47 tag, which
 * resolves to the built-in language it starts with and otherwise to the default.
 */
export type Locale = BuiltInLocale | (string & {});

/** Locale used when none is given or the requested one has no dictionary. */
export const DEFAULT_LOCALE: BuiltInLocale = 'en';

export const isBuiltInLocale = (locale: string): locale is BuiltInLocale =>
  (BUILT_IN_LOCALES as readonly string[]).includes(locale);

/** Resolve a BCP 47 or underscore locale to a built-in language id. */
export const builtInLocaleFor = (locale: string): BuiltInLocale | null => {
  const language = locale.trim().toLowerCase().replaceAll('_', '-').split('-')[0] ?? '';
  return isBuiltInLocale(language) ? language : null;
};
