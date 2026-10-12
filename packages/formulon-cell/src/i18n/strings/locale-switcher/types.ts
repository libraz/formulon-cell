import type { BuiltInLocale } from '../../locales.js';

export interface LocaleSwitcherStrings {
  localeSwitcher: {
    /** Label for the native locale select. */
    label: string;
    /** One autonym for each built-in locale. */
    options: Record<BuiltInLocale, string>;
    /** Announced when the engine rejects a locale change. */
    changeFailed: string;
  };
}
