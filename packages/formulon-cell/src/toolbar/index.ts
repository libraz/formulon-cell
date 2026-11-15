import type { RibbonToolbarConfig } from '../api/spreadsheet/types.js';
import type { RibbonOptions } from './toolbar-object.js';

export type { RibbonToolbarConfig } from '../api/spreadsheet/types.js';
export { type SaveWithPromptsOptions, saveWithPrompts } from './dialogs/workbook-save.js';
export { createExcelRibbonSvg, type ExcelRibbonIconName } from './excel-ribbon-icons.js';
export {
  type DisabledReasonProjectionOptions,
  projectDisabledReason,
} from './menu-a11y.js';
export type { RibbonHooks } from './ribbon/apply-ribbon-command.js';
export type { RibbonDisplayMode } from './ribbon/render-ribbon.js';
export {
  type BuildRibbonSearchIndexOptions,
  buildRibbonSearchIndex,
  type QueryRibbonSearchIndexOptions,
  queryRibbonSearchIndex,
  type RibbonSearchItem,
  type RibbonSearchItemKind,
  type RibbonSearchUsagePrior,
} from './ribbon/search-index.js';
export type {
  BuildRibbonModelOptions,
  RibbonCommand,
  RibbonGroupModel,
  RibbonOption,
  RibbonProfile,
  RibbonTab,
  RibbonTabModel,
} from './ribbon-model.js';
export {
  buildRibbonModel,
  CHART_CONTEXTUAL_RIBBON_TABS,
  CONTEXTUAL_RIBBON_TABS,
  EXCEL365_MAC_RIBBON_TABS,
  EXCEL365_STANDARD_RIBBON_TABS,
  OPTIONAL_RIBBON_TABS,
  PICTURE_CONTEXTUAL_RIBBON_TABS,
  RIBBON_TABS,
} from './ribbon-model.js';
export {
  type InstanceRibbon,
  type MountToolbarOptions,
  mountToolbar,
  type RibbonOptions,
  type Toolbar,
  type ToolbarPatch,
  type ToolbarState,
} from './toolbar-object.js';

/**
 * The ribbon toolbar, for `mount(host, { toolbar: ribbon(options) })` and `update({ toolbar })`.
 * The value is opaque: the options ride along under the names the ribbon mount reads.
 */
export const ribbon = (options: RibbonOptions = {}): RibbonToolbarConfig => {
  const { displayMode, ...rest } = options;
  return { kind: 'ribbon', ...rest, ...(displayMode ? { ribbonDisplayMode: displayMode } : {}) };
};
