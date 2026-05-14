import { RIBBON_TAB_LABELS, type RibbonTab } from './model.js';
import type { ToolbarLang } from './translations.js';

export const toolbarTabs = (lang: ToolbarLang): { id: RibbonTab; label: string }[] =>
  (Object.keys(RIBBON_TAB_LABELS) as RibbonTab[])
    .filter((id) => id !== 'file')
    .map((id) => ({ id, label: RIBBON_TAB_LABELS[id][lang] }));
