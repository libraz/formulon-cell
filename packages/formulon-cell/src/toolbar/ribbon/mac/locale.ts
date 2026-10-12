import { dictionaryLocaleFor } from '../../../i18n/strings.js';
import type { ToolbarLang } from '../../ribbon-model.js';

export const toolbarLangForLocale = (locale: string): ToolbarLang => dictionaryLocaleFor(locale);
