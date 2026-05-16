import { dictionaries, type Strings } from '../i18n/strings.js';
import type { ToolbarLang } from './ribbon-model.js';

const resolveStrings = (input: Strings | ToolbarLang): Strings =>
  typeof input === 'string' ? dictionaries[input] : input;

export type ConditionalMenuText = Strings['conditionalMenu'];
export const conditionalMenuText = (input: Strings | ToolbarLang): ConditionalMenuText =>
  resolveStrings(input).conditionalMenu;

export type ToolbarMenuText = Strings['ribbonMenu'];
export const toolbarMenuText = (input: Strings | ToolbarLang): ToolbarMenuText =>
  resolveStrings(input).ribbonMenu;

export type RibbonDisplayText = Strings['ribbonDisplay'];
export const ribbonDisplayText = (input: Strings | ToolbarLang): RibbonDisplayText =>
  resolveStrings(input).ribbonDisplay;

export type BackstageMenuText = Strings['backstage'];
export const backstageMenuText = (input: Strings | ToolbarLang): BackstageMenuText =>
  resolveStrings(input).backstage;

export type PageScaleMenuText = Strings['pageScale'];
export const pageScaleMenuText = (input: Strings | ToolbarLang): PageScaleMenuText =>
  resolveStrings(input).pageScale;

export type ViewToggleMenuText = Strings['viewToggle'];
export const viewToggleMenuText = (input: Strings | ToolbarLang): ViewToggleMenuText =>
  resolveStrings(input).viewToggle;
