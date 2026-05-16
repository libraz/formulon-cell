import { dictionaries, type Strings, type ToolbarLang } from '@libraz/formulon-cell';
import { RIBBON_TABS, type RibbonTab } from './model.js';

const resolveStrings = (strings: Strings | ToolbarLang): Strings =>
  typeof strings === 'string' ? dictionaries[strings] : strings;

export const toolbarTabs = (input: Strings | ToolbarLang): { id: RibbonTab; label: string }[] => {
  const strings = resolveStrings(input);
  return RIBBON_TABS.map((id) => ({
    id,
    label: strings.ribbon.tabs[id],
  }));
};
