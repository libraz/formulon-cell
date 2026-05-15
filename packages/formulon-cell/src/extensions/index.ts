// Extensions — the public surface for feature composition.
//
// v0.1 ships `features` flags + presets and Extension factories for
// every replaceable built-in. Pair `features: {<id>: false}` with
// `extensions: [<factory>()]` to substitute your own implementation.

// Built-in factories — wrap each `attach*` interact module as an
// Extension. Consumers replace a built-in by disabling it via
// `MountOptions.features` and re-mounting their own through `extensions`.
export {
  allBuiltIns,
  borderDraw,
  charts,
  clipboard,
  commentDialog,
  conditionalDialog,
  contextMenu,
  findReplace,
  formatDialog,
  formatPainter,
  goToSpecialDialog,
  hoverComment,
  hyperlinkDialog,
  iterativeDialog,
  namedRangeDialog,
  pageSetupDialog,
  pasteSpecial,
  pivotTableDialog,
  quickAnalysis,
  slicer,
  statusBar,
  validationList,
  viewToolbar,
  watchWindow,
  wheel,
  workbookObjects,
} from './built-ins.js';
export type { FeatureFlags, FeatureId } from './features.js';
export { ALL_FEATURE_IDS, resolveFlags } from './features.js';
export { full, minimal, presets, standard } from './presets.js';
export type {
  Extension,
  ExtensionContext,
  ExtensionHandle,
  ExtensionInput,
  I18nController,
  ThemeName,
} from './types.js';
export { dedupeById, flattenExtensions, sortByPriority } from './types.js';
