/** UI locales with a built-in dictionary and matching formulon profile. */
export const BUILT_IN_LOCALES = ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'] as const;

export type BuiltInLocale = (typeof BUILT_IN_LOCALES)[number];
/** Short locale id used by the built-in dictionaries and profile adapter. */
export type Locale = BuiltInLocale;

export const isBuiltInLocale = (locale: string): locale is BuiltInLocale =>
  (BUILT_IN_LOCALES as readonly string[]).includes(locale);

/** Resolve a BCP 47 or underscore locale to a built-in language id. */
export const builtInLocaleFor = (locale: string): BuiltInLocale | null => {
  const language = locale.trim().toLowerCase().replaceAll('_', '-').split('-')[0] ?? '';
  return isBuiltInLocale(language) ? language : null;
};
