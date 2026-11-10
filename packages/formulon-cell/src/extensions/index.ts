// Extensions — the public surface for feature composition.
//
// Feature flags + presets, and one Extension factory per replaceable
// built-in. Each factory is named after the feature id it provides, and an
// extension added under that id replaces the built-in.

// Built-in factories — wrap each `attach*` interact module as an Extension.
// Names the factories carried before they were renamed after their feature ids.
export {
  allBuiltIns,
  borderDraw,
  charts,
  clipboard,
  commentDialog,
  conditional,
  conditional as conditionalDialog,
  contextMenu,
  findReplace,
  formatDialog,
  formatPainter,
  gotoSpecial,
  gotoSpecial as goToSpecialDialog,
  hoverComment,
  hyperlink,
  hyperlink as hyperlinkDialog,
  illustrations,
  iterative,
  iterative as iterativeDialog,
  namedRanges,
  namedRanges as namedRangeDialog,
  pageSetup,
  pageSetup as pageSetupDialog,
  pasteSpecial,
  pivotTableDialog,
  quickAnalysis,
  slicer,
  statusBar,
  validation,
  validation as validationList,
  viewToolbar,
  watchWindow,
  wheel,
  workbookObjects,
} from './built-ins.js';
export type { FeatureFlags, FeatureId } from './features.js';
export {
  ALL_FEATURE_IDS,
  DEFAULT_OFF_FEATURE_IDS,
  isFeatureDefaultOff,
  resolveFlags,
} from './features.js';
export { full, minimal, presets, standard } from './presets.js';
export type {
  Extension,
  ExtensionContext,
  ExtensionHandle,
  ExtensionInput,
  ExtensionSpreadsheet,
  ThemeName,
} from './types.js';
export { dedupeById, flattenExtensions, sortByPriority } from './types.js';
export type {
  ResolvedSpreadsheetPlatform,
  ResolvedSpreadsheetUiOptions,
  SpreadsheetFeatureSwitches,
  SpreadsheetPlatform,
  SpreadsheetUiOptions,
  SpreadsheetUiProfile,
} from './ui-options.js';
export { resolveSpreadsheetPlatform, resolveSpreadsheetUiOptions } from './ui-options.js';
