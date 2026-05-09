import type { FeatureFlags } from './features.js';

// Presets express common feature bundles as a `features` shorthand. The
// names mirror Tiptap's "Kit" pattern: `minimal` is the headless surface,
// `standard` is what most apps want, `full` mirrors a full desktop
// Spreadsheet surface.
//
// In v0.1 every preset returns a `FeatureFlags` object. In v0.2 these will
// also return an `Extension[]` for compositional users; the API is
// designed so adding extensions later is purely additive.

/** Bare grid + formula bar + status bar + basic shortcuts. No menus,
 *  View toolbar, floating inspectors, dialogs, or hover comments. Useful
 *  for read-mostly views. */
export const minimal = (): FeatureFlags => ({
  viewToolbar: false,
  sheetTabs: false,
  workbookObjects: false,
  contextMenu: false,
  findReplace: false,
  formatDialog: false,
  formatPainter: false,
  conditional: false,
  iterative: false,
  gotoSpecial: false,
  namedRanges: false,
  hyperlink: false,
  commentDialog: false,
  fxDialog: false,
  pageSetup: false,
  pasteSpecial: false,
  quickAnalysis: false,
  charts: false,
  pivotTableDialog: false,
  validation: false,
  autocomplete: false,
  hoverComment: false,
  watchWindow: false,
  errorIndicators: false,
  slicer: false,
});

/** Adds clipboard, context menu, find/replace, View toolbar, Quick Analysis,
 *  workbook-object inspector, format painter, and wheel scroll. No fancy
 *  authoring dialogs (format/conditional/named-ranges/hyperlink). */
export const standard = (): FeatureFlags => ({
  formatDialog: false,
  conditional: false,
  iterative: false,
  gotoSpecial: false,
  namedRanges: false,
  hyperlink: false,
  commentDialog: false,
  fxDialog: false,
  pageSetup: false,
  pasteSpecial: false,
  pivotTableDialog: false,
  validation: false,
  hoverComment: false,
});

/** Full desktop-spreadsheet-style chrome — the default if no preset/features given. */
export const full = (): FeatureFlags => ({});

export const presets = { minimal, standard, full };
