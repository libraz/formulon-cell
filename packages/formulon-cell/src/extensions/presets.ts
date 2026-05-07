import type { FeatureFlags } from './features.js';

// Presets express common feature bundles as a `features` shorthand. The
// names mirror Tiptap's "Kit" pattern: `minimal` is the headless surface,
// `standard` is what most apps want, `excel` mirrors Excel 365's full
// chrome.
//
// In v0.1 every preset returns a `FeatureFlags` object. In v0.2 these will
// also return an `Extension[]` for compositional users; the API is
// designed so adding extensions later is purely additive.

/** Bare grid + formula bar + status bar + basic shortcuts. No menus, no
 *  dialogs, no hover comments. Useful for read-mostly views. */
export const minimal = (): FeatureFlags => ({
  sheetTabs: false,
  contextMenu: false,
  findReplace: false,
  formatDialog: false,
  formatPainter: false,
  conditional: false,
  namedRanges: false,
  hyperlink: false,
  fxDialog: false,
  pasteSpecial: false,
  validation: false,
  autocomplete: false,
  hoverComment: false,
});

/** Adds clipboard, context menu, find/replace, format painter, wheel
 *  scroll. No fancy dialogs (format/conditional/named-ranges/hyperlink). */
export const standard = (): FeatureFlags => ({
  formatDialog: false,
  conditional: false,
  namedRanges: false,
  hyperlink: false,
  fxDialog: false,
  pasteSpecial: false,
  validation: false,
  hoverComment: false,
});

/** Full Excel 365-style chrome — the default if no preset/features given. */
export const excel = (): FeatureFlags => ({});

export const presets = { minimal, standard, excel };
