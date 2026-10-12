import { dictionaries, type Locale, type Strings } from './strings.js';

/** Resolve a bundled dictionary for low-level attach APIs without a locale getter. */
export function bundledLocaleForStrings(strings: Strings): Locale | null {
  for (const [locale, dictionary] of Object.entries(dictionaries)) {
    if (
      strings === dictionary ||
      (strings.a11y.spreadsheet === dictionary.a11y.spreadsheet &&
        strings.formatDialog.title === dictionary.formatDialog.title)
    ) {
      return locale as Locale;
    }
  }
  return null;
}
