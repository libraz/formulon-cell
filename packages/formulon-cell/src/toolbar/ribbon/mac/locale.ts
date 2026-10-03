import type { ToolbarLang } from '../../ribbon-model.js';

/** Map a BCP 47 locale tag to the ribbon language: any `ja*` tag is Japanese, everything else English. */
export const toolbarLangForLocale = (locale: string): ToolbarLang =>
  locale.toLowerCase().startsWith('ja') ? 'ja' : 'en';
