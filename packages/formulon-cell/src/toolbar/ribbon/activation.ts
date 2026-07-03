// Canonical ribbon activation model.
//
// Menu buttons have a visible chevron and an attached dropdown, gallery, or
// secondary menu. Some menu buttons still open their menu on the primary face;
// Excel-style primary-action split buttons skip that menu-first fallback and
// dispatch the command directly, while keeping the secondary menu addressable
// through the dropdown API.

export type RibbonActivationKind =
  | 'primaryAction'
  | 'splitPrimary'
  | 'splitToggle'
  | 'dropdown'
  | 'gallery'
  | 'dialog'
  | 'toggle'
  | 'disabled';

export type RibbonActivationSpec = {
  kind: RibbonActivationKind;
  menuId?: string;
};

export type RibbonActivationEntry = RibbonActivationSpec & {
  command: string;
};

export const RIBBON_MENU_FACTORY_KEYS = [
  'paste',
  'copy',
  'pivotTable',
  'definedNames',
  'links',
  'borders',
  'underline',
  'wrap',
  'merge',
  'textOrientation',
  'conditional',
  'fill',
  'insertCells',
  'deleteCells',
  'formatCells',
  'autoSum',
  'freeze',
  'clearArrows',
  'errorChecking',
  'watch',
  'reviewComments',
  'protect',
  'calcOptions',
  'sort',
  'textToColumns',
  'dataValidation',
  'findSelect',
  'pictureInsert',
  'shapesInsert',
  'screenshotInsert',
  'chartInsert',
  'tableStyle',
  'cellStyles',
  'currency',
  'pageTheme',
  'arrange',
  'printArea',
  'pageBreaks',
  'symbol',
  'script',
  'addIn',
  'pdf',
  'clear',
] as const;

export type RibbonMenuFactoryKey = (typeof RIBBON_MENU_FACTORY_KEYS)[number];

export const RIBBON_BORDERS_MENU_ID = 'menu-borders';

export const RIBBON_EXTERNAL_MENU_FOR_COMMAND: Readonly<Record<string, string>> = {
  borders: RIBBON_BORDERS_MENU_ID,
};

export const RIBBON_DROPDOWN_MENU_FOR_COMMAND: Readonly<Record<string, string>> = {
  paste: 'menu-paste',
  copy: 'menu-copy',
  pivotTableInsert: 'menu-pivot-table',
  namedRanges: 'menu-defined-names',
  linksData: 'menu-links-data',
  conditional: 'menu-conditional',
  fillHome: 'menu-fill',
  clearFormat: 'menu-clear',
  underline: 'menu-underline',
  wrap: 'menu-wrap',
  merge: 'menu-merge',
  freeze: 'menu-freeze',
  textOrientation: 'menu-text-orientation',
  insertRows: 'menu-insert-cells',
  deleteRows: 'menu-delete-cells',
  formatCellsHome: 'menu-format-cells',
  pageTheme: 'menu-page-theme',
  pageBreaks: 'menu-page-breaks',
  printArea: 'menu-print-area',
  arrangeObjectsPageLayout: 'menu-arrange-objects',
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
  symbolInsert: 'menu-symbol',
  script: 'menu-script',
  formatTableHome: 'menu-table-style-home',
  cellStyles: 'menu-cell-styles-home',
  currency: 'menu-currency-home',
  textToColumns: 'menu-text-to-columns',
  dataValidation: 'menu-data-validation',
  addIn: 'menu-add-ins',
  pdf: 'menu-pdf',
};

export const RIBBON_MENU_FOR_COMMAND: Readonly<Record<string, string>> = {
  ...RIBBON_EXTERNAL_MENU_FOR_COMMAND,
  ...RIBBON_DROPDOWN_MENU_FOR_COMMAND,
};

export const RIBBON_MENU_FACTORY_FOR_COMMAND: Readonly<Record<string, RibbonMenuFactoryKey>> = {
  paste: 'paste',
  copy: 'copy',
  pivotTableInsert: 'pivotTable',
  namedRanges: 'definedNames',
  linksData: 'links',
  borders: 'borders',
  underline: 'underline',
  wrap: 'wrap',
  merge: 'merge',
  textOrientation: 'textOrientation',
  conditional: 'conditional',
  fillHome: 'fill',
  insertRows: 'insertCells',
  deleteRows: 'deleteCells',
  formatCellsHome: 'formatCells',
  autosum: 'autoSum',
  freeze: 'freeze',
  autosumFormula: 'autoSum',
  clearArrows: 'clearArrows',
  errorChecking: 'errorChecking',
  watch: 'watch',
  watchView: 'watch',
  deleteCommentReview: 'reviewComments',
  protectReview: 'protect',
  protect: 'protect',
  calcOptions: 'calcOptions',
  filter: 'sort',
  textToColumns: 'textToColumns',
  dataValidation: 'dataValidation',
  sortFilterHome: 'sort',
  findHome: 'findSelect',
  pictureInsert: 'pictureInsert',
  shapesInsert: 'shapesInsert',
  screenshotInsert: 'screenshotInsert',
  chartInsert: 'chartInsert',
  formatTableHome: 'tableStyle',
  cellStyles: 'cellStyles',
  currency: 'currency',
  pageTheme: 'pageTheme',
  arrangeObjectsPageLayout: 'arrange',
  printArea: 'printArea',
  pageBreaks: 'pageBreaks',
  symbolInsert: 'symbol',
  script: 'script',
  addIn: 'addIn',
  pdf: 'pdf',
  clearFormat: 'clear',
};

export const RIBBON_AUDITED_PRIMARY_ACTION_SPLIT_COMMANDS = [
  'paste',
  'pivotTableInsert',
  'autosum',
  'autosumFormula',
  'copy',
  'addIn',
  'chartInsert',
  'currency',
  'dataValidation',
  'deleteCommentReview',
  'errorChecking',
  'linksData',
  'merge',
  'namedRanges',
  'protect',
  'protectReview',
  'pdf',
  'script',
  'symbolInsert',
  'textToColumns',
  'watch',
  'watchView',
] as const;

export const RIBBON_AUDITED_SPLIT_TOGGLE_COMMANDS = ['underline'] as const;

export const RIBBON_AUDITED_GALLERY_COMMANDS = [
  'conditional',
  'pageTheme',
  'pictureInsert',
  'shapesInsert',
  'screenshotInsert',
  'formatTableHome',
  'cellStyles',
] as const;

export const RIBBON_AUDITED_DROPDOWN_COMMANDS = [
  'arrangeObjectsPageLayout',
  'borders',
  'calcOptions',
  'clearArrows',
  'clearFormat',
  'deleteRows',
  'fillHome',
  'filter',
  'findHome',
  'formatCellsHome',
  'freeze',
  'insertRows',
  'pageBreaks',
  'printArea',
  'sortFilterHome',
  'textOrientation',
  'wrap',
] as const;

/** Commands intentionally classified in the activation model even though no
 *  top-level ribbon button currently renders them. They are used by shared
 *  menus, legacy shortcuts, or dialog opener aliases, so keeping the list
 *  explicit prevents stale command ids from accumulating silently. */
export const RIBBON_INTENTIONAL_NON_RENDERED_COMMANDS = [
  'drawBorder',
  'drawBorderGrid',
  'eraseBorder',
  'formatCells',
  'fxInsert',
  'gotoSpecial',
  'gotoSpecialHome',
  'links',
  'moreBorders',
  'rules',
] as const;

export const RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS: ReadonlySet<string> = new Set(
  RIBBON_AUDITED_PRIMARY_ACTION_SPLIT_COMMANDS,
);

export const RIBBON_SPLIT_TOGGLE_COMMANDS: ReadonlySet<string> = new Set(
  RIBBON_AUDITED_SPLIT_TOGGLE_COMMANDS,
);

export const RIBBON_PRIMARY_FACE_MENU_COMMANDS: ReadonlySet<string> = new Set([
  ...RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS,
  ...RIBBON_SPLIT_TOGGLE_COMMANDS,
]);

export const RIBBON_MENU_FIRST_COMMANDS: ReadonlySet<string> = new Set(
  Object.keys(RIBBON_MENU_FOR_COMMAND).filter(
    (command) => !RIBBON_PRIMARY_FACE_MENU_COMMANDS.has(command),
  ),
);

export const RIBBON_DYNAMIC_MENU_FIRST_COMMANDS: ReadonlySet<string> = new Set(
  Object.keys(RIBBON_DROPDOWN_MENU_FOR_COMMAND).filter((command) =>
    RIBBON_MENU_FIRST_COMMANDS.has(command),
  ),
);

export const RIBBON_EXTERNAL_MENU_FIRST_COMMANDS: ReadonlySet<string> = new Set(
  Object.keys(RIBBON_EXTERNAL_MENU_FOR_COMMAND).filter((command) =>
    RIBBON_MENU_FIRST_COMMANDS.has(command),
  ),
);

export const RIBBON_PRIMARY_ACTION_COMMANDS: ReadonlySet<string> = new Set([
  'accessibility',
  'alignC',
  'alignL',
  'alignR',
  'allScripts',
  'bottomAlign',
  'comma',
  'cut',
  'decDown',
  'decUp',
  'dependents',
  'drawErase',
  'drawPen',
  'findReview',
  'fontGrow',
  'fontShrink',
  'inspect',
  'indentDecrease',
  'indentIncrease',
  'middle',
  'nextCommentReview',
  'outlineGroup',
  'outlineHideDetail',
  'outlineShowDetail',
  'outlineUngroup',
  'percent',
  'precedents',
  'previousCommentReview',
  'print',
  'printPageLayout',
  'protectWorkbookReview',
  'protectionReview',
  'recalcNow',
  'recordActions',
  'removeDupes',
  'sheetBackground',
  'sheetViewDelete',
  'sheetViewSave',
  'sortAsc',
  'sortData',
  'sortDesc',
  'spellingReview',
  'top',
  'translateReview',
  'viewNormal',
  'viewPageBreakPreview',
  'viewPageLayout',
  'zoom100',
  'zoom125',
  'zoom75',
  'zoomSelection',
]);

export const RIBBON_GALLERY_COMMANDS: ReadonlySet<string> = new Set(
  RIBBON_AUDITED_GALLERY_COMMANDS,
);

export const RIBBON_DROPDOWN_COMMANDS: ReadonlySet<string> = new Set(
  RIBBON_AUDITED_DROPDOWN_COMMANDS,
);

export const RIBBON_DIALOG_COMMANDS: ReadonlySet<string> = new Set([
  'pageSetup',
  'pageSetupAdvanced',
  'formatCells',
  'moreBorders',
  'windowVisibility',
  'gotoSpecial',
  'gotoSpecialHome',
  'hyperlinkInsert',
  'commentInsert',
  'newCommentReview',
  'links',
  'rules',
  'evaluateFormula',
  'fxInsert',
  'fx',
  'sum',
  'avg',
  'ifFormula',
  'xlookupFormula',
  'concatFormula',
  'todayFormula',
  'pmtFormula',
  'roundFormula',
  'workbookObjectsView',
  'selectionPanePageLayout',
  'pivotFieldListView',
  'formatTableInsert',
  'printTitles',
  'zoomDialog',
]);

export const RIBBON_TOGGLE_COMMANDS: ReadonlySet<string> = new Set([
  'bold',
  'italic',
  'strike',
  'formatPainter',
  'drawBorder',
  'drawBorderGrid',
  'drawGrid',
  'eraseBorder',
  'pageLayoutGridlinesPrint',
  'pageLayoutGridlinesView',
  'pageLayoutHeadingsPrint',
  'pageLayoutHeadingsView',
  'showFormulasFormula',
  'viewGridlines',
  'viewHeadings',
  'viewFormulas',
  'viewFormulaBar',
  'viewR1C1',
]);

export const RIBBON_DISABLED_COMMANDS: ReadonlySet<string> = new Set(['helpSearch']);

export type RibbonActivationCategoryEntry = readonly [
  kind: RibbonActivationKind,
  commands: ReadonlySet<string>,
];

export const ribbonActivationCategories = (): readonly RibbonActivationCategoryEntry[] => [
  ['splitPrimary', RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS],
  ['splitToggle', RIBBON_SPLIT_TOGGLE_COMMANDS],
  ['gallery', RIBBON_GALLERY_COMMANDS],
  ['dropdown', RIBBON_DROPDOWN_COMMANDS],
  ['primaryAction', RIBBON_PRIMARY_ACTION_COMMANDS],
  ['dialog', RIBBON_DIALOG_COMMANDS],
  ['toggle', RIBBON_TOGGLE_COMMANDS],
  ['disabled', RIBBON_DISABLED_COMMANDS],
];

export const ribbonActivationCommandIds = (): string[] =>
  Array.from(
    new Set([
      ...Object.keys(RIBBON_MENU_FOR_COMMAND),
      ...RIBBON_PRIMARY_ACTION_COMMANDS,
      ...RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS,
      ...RIBBON_SPLIT_TOGGLE_COMMANDS,
      ...RIBBON_GALLERY_COMMANDS,
      ...RIBBON_DIALOG_COMMANDS,
      ...RIBBON_TOGGLE_COMMANDS,
      ...RIBBON_DISABLED_COMMANDS,
    ]),
  ).sort();

export const ribbonActivationEntries = (): readonly RibbonActivationEntry[] =>
  ribbonActivationCommandIds().map((command) => ({
    command,
    ...ribbonActivationForCommand(command),
  }));

export const ribbonActivationEntriesForCommands = (
  commandIds: Iterable<string>,
): readonly RibbonActivationEntry[] =>
  Array.from(new Set(commandIds))
    .sort()
    .map((command) => ({
      command,
      ...ribbonActivationForCommand(command),
    }));

export const RIBBON_SPLIT_BUTTON_COMMANDS: ReadonlySet<string> = new Set(
  Object.keys(RIBBON_MENU_FOR_COMMAND),
);

export const ribbonActivationForCommand = (commandId: string): RibbonActivationSpec => {
  const menuId = RIBBON_MENU_FOR_COMMAND[commandId];
  if (menuId) {
    if (RIBBON_SPLIT_TOGGLE_COMMANDS.has(commandId)) {
      return { kind: 'splitToggle', menuId };
    }
    if (RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS.has(commandId)) {
      return { kind: 'splitPrimary', menuId };
    }
    if (RIBBON_GALLERY_COMMANDS.has(commandId)) {
      return { kind: 'gallery', menuId };
    }
    return { kind: 'dropdown', menuId };
  }
  if (RIBBON_DISABLED_COMMANDS.has(commandId)) return { kind: 'disabled' };
  if (RIBBON_DIALOG_COMMANDS.has(commandId)) return { kind: 'dialog' };
  if (RIBBON_TOGGLE_COMMANDS.has(commandId)) return { kind: 'toggle' };
  if (RIBBON_PRIMARY_ACTION_COMMANDS.has(commandId)) return { kind: 'primaryAction' };
  return { kind: 'disabled' };
};
