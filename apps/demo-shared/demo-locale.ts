import { BUILT_IN_LOCALES, type Locale } from '@libraz/formulon-cell';

const isBuiltInLocale = (value: string): value is Locale =>
  (BUILT_IN_LOCALES as readonly string[]).includes(value);

/** The built-in dictionary for a locale tag (`ja-JP` → `ja`), English for any other language. */
export const demoDictionaryLocale = (locale: string): Locale => {
  const language = locale.trim().toLowerCase().replaceAll('_', '-').split('-')[0] ?? '';
  return isBuiltInLocale(language) ? language : 'en';
};
