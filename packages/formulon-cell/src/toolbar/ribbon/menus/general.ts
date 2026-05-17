// Generic menu primitives shared across all ribbon dropdown factories.
// Extracted from main.ts so the per-tab menu modules can build off the same
// building blocks without dragging the whole playground entry along with them.

import { prepareMenu } from '../../menu-a11y.js';

// Inverse of dynamic-dropdowns' RIBBON_DROPDOWN_MENU_FOR_COMMAND — kept here
// so menu factories that vary by command id (autosum / autosumFormula,
// sort / sortFilterHome / filter, …) can pick the DOM id their data-menu
// wirings expect without each factory hard-coding the lookup table.
const MENU_ID_FOR_COMMAND: Readonly<Record<string, string>> = {
  paste: 'menu-paste',
  pivotTableInsert: 'menu-pivot-table',
  namedRangesInsert: 'menu-defined-names-insert',
  namedRanges: 'menu-defined-names',
  links: 'menu-links-file',
  linksInsert: 'menu-links-insert',
  linksData: 'menu-links-data',
  conditional: 'menu-conditional',
  fillHome: 'menu-fill',
  textOrientation: 'menu-text-orientation',
  insertRows: 'menu-insert-cells',
  deleteRows: 'menu-delete-cells',
  formatCellsHome: 'menu-format-cells',
  pageTheme: 'menu-page-theme',
  pageBreaks: 'menu-page-breaks',
  sheetBackground: 'menu-sheet-background',
  printTitles: 'menu-print-titles',
  sortFilterHome: 'menu-sort-home',
  filter: 'menu-sort',
  findHome: 'menu-find-select',
  autosum: 'menu-autosum-home',
  autosumFormula: 'menu-autosum-formulas',
  clearArrows: 'menu-clear-arrows',
  errorChecking: 'menu-error-checking',
  watch: 'menu-watch-formulas',
  watchView: 'menu-watch-view',
  deleteCommentReview: 'menu-review-comments',
  protectReview: 'menu-protect-review',
  protect: 'menu-protect-view',
  calcOptions: 'menu-calc-options',
  chartInsert: 'menu-chart-insert',
  pictureInsert: 'menu-picture-insert',
  shapesInsert: 'menu-shapes-insert',
  screenshotInsert: 'menu-screenshot-insert',
  script: 'menu-script',
  formatTableHome: 'menu-table-style-home',
  formatTableInsert: 'menu-table-style-insert',
  cellStyles: 'menu-cell-styles-home',
  currency: 'menu-currency-home',
  textToColumns: 'menu-text-to-columns',
  dataValidation: 'menu-data-validation',
  addIn: 'menu-add-ins',
  pdf: 'menu-pdf',
};

/** Maps a ribbon command id to the DOM id its dropdown menu should use.
 *  Falls back to the command id when no entry exists (callers can pass any
 *  string and get a stable id back). */
export const menuIdForCommand = (commandId: string): string =>
  MENU_ID_FOR_COMMAND[commandId] ?? commandId;

export const createMenu = (id: string): HTMLDivElement => {
  const menu = document.createElement('div');
  menu.className = 'app__menu';
  menu.id = id;
  menu.hidden = true;
  prepareMenu(menu);
  return menu;
};

export const menuButton = (label: string, attr: string, value: string): HTMLButtonElement => {
  const button = document.createElement('button');
  button.className = 'app__menu-item';
  button.type = 'button';
  button.setAttribute('role', 'menuitem');
  button.dataset[attr] = value;
  button.textContent = label;
  return button;
};

export const menuSeparator = (): HTMLDivElement => {
  const sep = document.createElement('div');
  sep.className = 'app__menu-sep';
  sep.setAttribute('role', 'separator');
  return sep;
};

export const menuSectionHeader = (label: string): HTMLDivElement => {
  const el = document.createElement('div');
  el.className = 'app__menu-heading';
  el.setAttribute('role', 'presentation');
  el.textContent = label;
  return el;
};
