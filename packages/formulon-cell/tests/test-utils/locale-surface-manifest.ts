import type { BuiltInLocale } from '../../src/i18n/locales.js';

/** The seven dictionaries that are shipped by the package. Keep this list
 * explicit: a new dictionary is a product-surface change and should update
 * the runtime matrix at the same time. */
export const LOCALE_SURFACE_LOCALES = ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'] as const;

export type LocaleSurfaceLocale = (typeof LOCALE_SURFACE_LOCALES)[number] & BuiltInLocale;

export type LocaleSurfacePolicy = 'live' | 'block' | 'transient-close' | 'startup';

/**
 * Top-level dictionary sections. The source dictionaries intentionally group
 * related controls, while the factory manifest below gives each opener its
 * own coverage id. Keeping these as data makes the coverage report useful to
 * reviewers and makes additions fail loudly in the focused test.
 */
export const LOCALE_SURFACE_SECTIONS = [
  'cellStylesGallery',
  'formatDialog',
  'conditionalDialog',
  'cfRulesDialog',
  'pivotTableDialog',
  'workbookObjects',
  'sessionCharts',
  'sessionIllustrations',
  'slicer',
  'fillSeriesDialog',
  'hyperlinkDialog',
  'commentDialog',
  'pasteSpecialDialog',
  'insertCopiedCellsDialog',
  'findReplace',
  'goToDialog',
  'protection',
  'namedRangeDialog',
  'iterativeDialog',
  'evaluateFormulaDialog',
  'externalLinksDialog',
  'fxDialog',
  'watchPanel',
  'contextMenu',
  'autoFillOptions',
  'pasteOptions',
  'autocomplete',
  'argHelper',
  'quickAnalysis',
  'filterDropdown',
  'errorMenu',
  'localeSwitcher',
  'macData',
  'macSparkline',
  'macSlicer',
  'macWorkbookStats',
  'macInk',
  'macContextMenu',
  'macHomeRibbon',
  'pageView',
  'pageSetup',
  'pageScale',
  'ribbonMenu',
  'conditionalMenu',
  'toolbar',
  'ribbon',
  'backstage',
  'ribbonDisplay',
  'viewToggle',
  'reviewReports',
  'statusBar',
  'viewToolbar',
  'sheetTabs',
  'a11y',
  'mountError',
] as const;

export type LocaleSurfaceSection = (typeof LOCALE_SURFACE_SECTIONS)[number];

export interface LocaleSurfaceFactory {
  /** Stable id used in coverage output. */
  readonly id: string;
  /** Source path relative to `src/`. */
  readonly source: string;
  /** Export or source marker that proves the opener is still wired. */
  readonly marker: string;
  readonly section: LocaleSurfaceSection;
  readonly policy: LocaleSurfacePolicy;
  /** A source-level factory may not have a cheap safe DOM opener. */
  readonly runtime: 'public' | 'shell-source' | 'manual-source' | 'unavailable';
  readonly reason?: string;
}

export type LocaleToolbarRouteKind = 'command' | 'dynamic-menu' | 'ribbon-select';

/**
 * Mounted toolbar routes whose handlers open a locale-sensitive surface (or
 * intentionally mutate a protection setting). The integration test consumes
 * this data instead of constructing the handler options itself, so a route is
 * only covered when the rendered button and its delegated menu item work.
 */
export interface LocaleSurfaceToolbarRoute {
  readonly id: string;
  readonly section: LocaleSurfaceSection;
  readonly tab: string;
  readonly command: string;
  readonly kind: LocaleToolbarRouteKind;
  readonly menuId?: string;
  readonly actionAttr?: string;
  readonly action?: string;
  readonly controlValue?: string;
  readonly titlePath: string;
  readonly factoryId: string;
  readonly cancelable: boolean;
}

export const LOCALE_SURFACE_TOOLBAR_ROUTES: readonly LocaleSurfaceToolbarRoute[] = [
  {
    id: 'toolbar.sort.custom',
    section: 'ribbonMenu',
    tab: 'data',
    command: 'sortFilterHome',
    kind: 'dynamic-menu',
    menuId: 'menu-sort-home',
    actionAttr: 'sort',
    action: 'custom',
    titlePath: 'ribbonMenu.sortDialogTitle',
    factoryId: 'toolbar.sort',
    cancelable: true,
  },
  {
    id: 'toolbar.filter.advanced',
    section: 'ribbonMenu',
    tab: 'data',
    command: 'sortFilterHome',
    kind: 'dynamic-menu',
    menuId: 'menu-sort-home',
    actionAttr: 'sort',
    action: 'filter-advanced',
    titlePath: 'ribbonMenu.advancedFilterDialogTitle',
    factoryId: 'toolbar.advancedFilter',
    cancelable: true,
  },
  {
    id: 'toolbar.removeDuplicates',
    section: 'ribbonMenu',
    tab: 'data',
    command: 'sortFilterHome',
    kind: 'dynamic-menu',
    menuId: 'menu-sort-home',
    actionAttr: 'sort',
    action: 'dedupe',
    titlePath: 'ribbonMenu.removeDuplicatesDialogTitle',
    factoryId: 'toolbar.removeDuplicates',
    cancelable: true,
  },
  {
    id: 'toolbar.textToColumns',
    section: 'ribbonMenu',
    tab: 'data',
    command: 'textToColumns',
    kind: 'dynamic-menu',
    menuId: 'menu-text-to-columns',
    actionAttr: 'textToColumnsDelimiter',
    action: 'custom',
    titlePath: 'ribbonMenu.textToColumnsDialogTitle',
    factoryId: 'toolbar.textToColumns',
    cancelable: true,
  },
  {
    id: 'toolbar.protection.protectSheet',
    section: 'protection',
    tab: 'review',
    command: 'protectReview',
    kind: 'dynamic-menu',
    menuId: 'menu-protect-review',
    actionAttr: 'protectAction',
    action: 'protect-sheet',
    titlePath: 'protection.protectSheet',
    factoryId: 'toolbar.protection',
    cancelable: true,
  },
  {
    id: 'toolbar.protection.unprotectSheet',
    section: 'protection',
    tab: 'review',
    command: 'protectReview',
    kind: 'dynamic-menu',
    menuId: 'menu-protect-review',
    actionAttr: 'protectAction',
    action: 'unprotect-sheet',
    titlePath: 'protection.unprotectSheet',
    factoryId: 'toolbar.protection',
    cancelable: true,
  },
  {
    id: 'toolbar.protection.protectWorkbook',
    section: 'protection',
    tab: 'review',
    command: 'protectReview',
    kind: 'dynamic-menu',
    menuId: 'menu-protect-review',
    actionAttr: 'protectAction',
    action: 'protect-workbook',
    titlePath: 'ribbonMenu.protectWorkbookCommand',
    factoryId: 'toolbar.protection',
    cancelable: false,
  },
  {
    id: 'toolbar.protection.unprotectWorkbook',
    section: 'protection',
    tab: 'review',
    command: 'protectReview',
    kind: 'dynamic-menu',
    menuId: 'menu-protect-review',
    actionAttr: 'protectAction',
    action: 'unprotect-workbook',
    titlePath: 'ribbonMenu.unprotectWorkbookCommand',
    factoryId: 'toolbar.protection',
    cancelable: false,
  },
  {
    id: 'toolbar.protection.allowEditRanges',
    section: 'protection',
    tab: 'review',
    command: 'protectReview',
    kind: 'dynamic-menu',
    menuId: 'menu-protect-review',
    actionAttr: 'protectAction',
    action: 'allow-edit-ranges',
    titlePath: 'ribbonMenu.allowEditRangesDialogTitle',
    factoryId: 'toolbar.protection',
    cancelable: true,
  },
  {
    id: 'toolbar.formatAsTable',
    section: 'ribbonMenu',
    tab: 'insert',
    command: 'formatTableInsert',
    kind: 'command',
    titlePath: 'pivotTableDialog.createTableTitle',
    factoryId: 'toolbar.formatAsTable',
    cancelable: true,
  },
  {
    id: 'toolbar.tableStyle',
    section: 'ribbonMenu',
    tab: 'home',
    command: 'formatTableHome',
    kind: 'dynamic-menu',
    menuId: 'menu-table-style-home',
    actionAttr: 'tableStyleFooter',
    action: 'new-table-style',
    titlePath: 'ribbonMenu.tableStyleNew',
    factoryId: 'toolbar.tableStyle',
    cancelable: true,
  },
  {
    id: 'toolbar.renameSheet',
    section: 'sheetTabs',
    tab: 'home',
    command: 'formatCellsHome',
    kind: 'dynamic-menu',
    menuId: 'menu-format-cells',
    actionAttr: 'cellFormat',
    action: 'rename-sheet',
    titlePath: 'sheetTabs.rename',
    factoryId: 'toolbar.renameSheet',
    cancelable: true,
  },
  {
    id: 'toolbar.symbol',
    section: 'ribbonMenu',
    tab: 'insert',
    command: 'symbolInsert',
    kind: 'dynamic-menu',
    menuId: 'menu-symbol',
    actionAttr: 'symbolAction',
    action: 'more',
    titlePath: 'ribbonMenu.symbolMore',
    factoryId: 'toolbar.symbol',
    cancelable: true,
  },
  {
    id: 'toolbar.fillSeries',
    section: 'fillSeriesDialog',
    tab: 'home',
    command: 'fillHome',
    kind: 'dynamic-menu',
    menuId: 'menu-fill',
    actionAttr: 'fill',
    action: 'series',
    titlePath: 'fillSeriesDialog.title',
    factoryId: 'toolbar.fillSeries',
    cancelable: true,
  },
  {
    id: 'toolbar.pageScale',
    section: 'pageScale',
    tab: 'pageLayout',
    command: 'scalePercent',
    kind: 'ribbon-select',
    controlValue: 'custom',
    titlePath: 'ribbon.scale',
    factoryId: 'toolbar.pageScalePrompt',
    cancelable: true,
  },
] as const;

/** Browser/system-owned surfaces remain source-covered but cannot be mounted
 *  deterministically by this DOM matrix. Keep the reason beside the route so
 *  reports cannot silently imply that print/capture/file input was exercised. */
export const LOCALE_SURFACE_TOOLBAR_UNAVAILABLE: readonly {
  readonly id: string;
  readonly section: LocaleSurfaceSection;
  readonly reason: string;
}[] = [
  {
    id: 'toolbar.print',
    section: 'ribbon',
    reason: 'The print command hands control to the browser/system print surface.',
  },
  {
    id: 'toolbar.capture',
    section: 'ribbon',
    reason: 'Screenshot capture requires a browser capture/user gesture boundary.',
  },
  {
    id: 'toolbar.filePicker',
    section: 'ribbon',
    reason: 'Picture/background file selection requires a browser file-picker gesture.',
  },
] as const;

const shell = (
  id: string,
  source: string,
  marker: string,
  section: LocaleSurfaceSection,
  policy: LocaleSurfacePolicy,
): LocaleSurfaceFactory => ({
  id,
  source,
  marker,
  section,
  policy,
  runtime: 'shell-source',
});

const manual = (
  id: string,
  source: string,
  marker: string,
  section: LocaleSurfaceSection,
  policy: LocaleSurfacePolicy,
  runtime: LocaleSurfaceFactory['runtime'] = 'manual-source',
  reason?: string,
): LocaleSurfaceFactory => ({ id, source, marker, section, policy, runtime, reason });

/**
 * Every module that calls the shared modal shell. This is deliberately a
 * source manifest rather than a list of test imports: it catches a new modal
 * call site before it can silently escape locale-switch coverage.
 */
export const LOCALE_SURFACE_SHELL_FACTORIES: readonly LocaleSurfaceFactory[] = [
  shell(
    'interact.cellStylesGallery',
    'interact/cell-styles-gallery.ts',
    'attachCellStylesGallery',
    'cellStylesGallery',
    'block',
  ),
  shell(
    'interact.cfRulesDialog',
    'interact/cf-rules-dialog.ts',
    'attachCfRulesDialog',
    'cfRulesDialog',
    'block',
  ),
  shell(
    'interact.conditionalDialog',
    'interact/conditional-dialog.ts',
    'attachConditionalDialog',
    'conditionalDialog',
    'block',
  ),
  shell(
    'interact.evaluateFormulaDialog',
    'interact/evaluate-formula-dialog.ts',
    'attachEvaluateFormulaDialog',
    'evaluateFormulaDialog',
    'block',
  ),
  shell(
    'interact.externalLinksDialog',
    'interact/external-links-dialog.ts',
    'attachExternalLinksDialog',
    'externalLinksDialog',
    'block',
  ),
  shell(
    'interact.formatDialogView',
    'interact/format-dialog-view.ts',
    'createFormatDialogView',
    'formatDialog',
    'block',
  ),
  shell('interact.fxDialog', 'interact/fx-dialog.ts', 'attachFxDialog', 'fxDialog', 'block'),
  shell(
    'interact.goToDialog',
    'interact/goto-dialog.ts',
    'attachGoToDialog',
    'goToDialog',
    'block',
  ),
  shell(
    'interact.hyperlinkDialog',
    'interact/hyperlink-dialog.ts',
    'attachHyperlinkDialog',
    'hyperlinkDialog',
    'block',
  ),
  shell(
    'interact.iterativeDialog',
    'interact/iterative-dialog.ts',
    'attachIterativeDialog',
    'iterativeDialog',
    'block',
  ),
  shell(
    'interact.macConsolidateDialog',
    'interact/mac-consolidate-dialog.ts',
    'openMacConsolidateDialog',
    'macData',
    'block',
  ),
  shell(
    'interact.macGoalSeekDialog',
    'interact/mac-goal-seek-dialog.ts',
    'openMacGoalSeekDialog',
    'macData',
    'block',
  ),
  shell(
    'interact.macSlicerDialog',
    'interact/mac-slicer-dialog.ts',
    'attachMacSlicerDialog',
    'macSlicer',
    'block',
  ),
  shell(
    'interact.macSparklineDialog',
    'interact/mac-sparkline-dialog.ts',
    'attachMacSparklineDialog',
    'macSparkline',
    'block',
  ),
  shell(
    'interact.macSubtotalDialog',
    'interact/mac-subtotal-dialog.ts',
    'openMacSubtotalDialog',
    'macData',
    'block',
  ),
  shell(
    'interact.macWorkbookStatistics',
    'interact/mac-workbook-statistics.ts',
    'attachMacWorkbookStatistics',
    'macWorkbookStats',
    'block',
  ),
  shell(
    'interact.namedRangeDialog',
    'interact/named-range-dialog.ts',
    'attachNamedRangeDialog',
    'namedRangeDialog',
    'block',
  ),
  shell(
    'interact.pageSetupDialog',
    'interact/page-setup-dialog.ts',
    'attachPageSetupDialog',
    'pageSetup',
    'block',
  ),
  shell(
    'interact.pasteSpecial',
    'interact/paste-special.ts',
    'attachPasteSpecial',
    'pasteSpecialDialog',
    'block',
  ),
  shell(
    'interact.pivotFieldSettings',
    'interact/pivot-field-settings.ts',
    'showPivotFilterDialog',
    'pivotTableDialog',
    'block',
  ),
  shell(
    'interact.pivotTableDialog',
    'interact/pivot-table-dialog.ts',
    'attachPivotTableDialog',
    'pivotTableDialog',
    'block',
  ),
  shell(
    'interact.validation',
    'interact/validation.ts',
    'attachValidationAlert',
    'protection',
    'block',
  ),
  shell(
    'toolbar.advancedFilter',
    'toolbar/dialogs/advanced-filter.ts',
    'showAdvancedFilterDialog',
    'ribbonMenu',
    'block',
  ),
  shell(
    'toolbar.cellStyle',
    'toolbar/dialogs/cell-style.ts',
    'showCellStyleDialog',
    'cellStylesGallery',
    'block',
  ),
  shell('toolbar.choice', 'toolbar/dialogs/choice.ts', 'showChoiceDialog', 'ribbonMenu', 'block'),
  shell(
    'toolbar.conditionalFormatStyle',
    'toolbar/dialogs/conditional-format-style.ts',
    'showConditionalFormatCustomStyleDialog',
    'conditionalMenu',
    'block',
  ),
  shell(
    'toolbar.conditionalFormat',
    'toolbar/dialogs/conditional-format.ts',
    'showConditionalFormatNumberDialog',
    'conditionalMenu',
    'block',
  ),
  shell(
    'toolbar.definedNamePicker',
    'toolbar/dialogs/defined-name-picker.ts',
    'showDefinedNamePickerDialog',
    'namedRangeDialog',
    'block',
  ),
  shell(
    'toolbar.formatAsTable',
    'toolbar/dialogs/format-as-table.ts',
    'showFormatAsTableDialog',
    'ribbonMenu',
    'block',
  ),
  shell('toolbar.prompt', 'toolbar/dialogs/prompt.ts', 'showPrompt', 'ribbonMenu', 'block'),
  shell(
    'toolbar.protection',
    'toolbar/dialogs/protection.ts',
    'showProtectSheetDialog',
    'protection',
    'block',
  ),
  shell(
    'toolbar.removeDuplicates',
    'toolbar/dialogs/remove-duplicates.ts',
    'showRemoveDuplicatesDialog',
    'ribbonMenu',
    'block',
  ),
  shell(
    'toolbar.renameSheet',
    'toolbar/dialogs/rename-sheet.ts',
    'showRenameSheetDialog',
    'sheetTabs',
    'block',
  ),
  shell('toolbar.report', 'toolbar/dialogs/report.ts', 'showReport', 'reviewReports', 'block'),
  shell(
    'toolbar.scriptCommand',
    'toolbar/dialogs/script-command.ts',
    'showScriptCommandDialog',
    'ribbonMenu',
    'block',
  ),
  shell('toolbar.sort', 'toolbar/dialogs/sort.ts', 'showSortDialog', 'ribbonMenu', 'block'),
  shell('toolbar.symbol', 'toolbar/dialogs/symbol.ts', 'showSymbolDialog', 'ribbonMenu', 'block'),
  shell(
    'toolbar.tableStyle',
    'toolbar/dialogs/table-style.ts',
    'showTableStyleDialog',
    'ribbonMenu',
    'block',
  ),
  shell(
    'toolbar.textToColumns',
    'toolbar/dialogs/text-to-columns.ts',
    'showTextToColumnsDialog',
    'ribbonMenu',
    'block',
  ),
  shell(
    'toolbar.fillSeries',
    'toolbar/ribbon/fill-series.ts',
    'showFillSeriesDialog',
    'fillSeriesDialog',
    'block',
  ),
  shell(
    'interact.macAutomationGallery',
    'toolbar/ribbon/mac/automation-gallery.ts',
    'openMacAutomationGallery',
    'macInk',
    'transient-close',
  ),
] as const;

/** Source paths that are locale-sensitive but do not use the shared shell. */
export const LOCALE_SURFACE_MANUAL_FACTORIES: readonly LocaleSurfaceFactory[] = [
  manual(
    'interact.commentDialog',
    'interact/comment-dialog.ts',
    'attachCommentDialog',
    'commentDialog',
    'block',
  ),
  manual(
    'interact.findReplace',
    'interact/find-replace.ts',
    'attachFindReplace',
    'findReplace',
    'block',
  ),
  manual(
    'interact.filterDropdown',
    'interact/filter-dropdown.ts',
    'attachFilterDropdown',
    'filterDropdown',
    'block',
  ),
  manual(
    'interact.insertCopiedCellsDialog',
    'interact/insert-copied-cells-dialog.ts',
    'openInsertCopiedCellsDialog',
    'insertCopiedCellsDialog',
    'block',
  ),
  manual(
    'interact.cellShiftDialog',
    'interact/cell-shift-dialog.ts',
    'openCellShiftDialog',
    'insertCopiedCellsDialog',
    'block',
  ),
  manual(
    'interact.quickAnalysis',
    'interact/quick-analysis.ts',
    'attachQuickAnalysis',
    'quickAnalysis',
    'transient-close',
  ),
  manual(
    'interact.workbookObjects',
    'interact/workbook-objects.ts',
    'attachWorkbookObjectsPanel',
    'workbookObjects',
    'live',
  ),
  manual(
    'interact.watchPanel',
    'interact/watch-panel.ts',
    'attachWatchPanel',
    'watchPanel',
    'live',
  ),
  manual('interact.slicer', 'interact/slicer.ts', 'attachSlicer', 'slicer', 'live'),
  manual(
    'interact.sessionCharts',
    'interact/session-charts.ts',
    'attachSessionCharts',
    'sessionCharts',
    'live',
  ),
  manual(
    'interact.sessionIllustrations',
    'interact/session-illustrations.ts',
    'attachSessionIllustrations',
    'sessionCharts',
    'live',
  ),
  manual(
    'toolbar.backstage',
    'toolbar/ribbon/backstage.ts',
    'createBackstageView',
    'backstage',
    'block',
  ),
  manual(
    'toolbar.dimensionPrompt',
    'toolbar/dialogs/dimension.ts',
    'showDimensionDialog',
    'ribbonMenu',
    'block',
  ),
  manual(
    'toolbar.pageScalePrompt',
    'toolbar/dialogs/page-scale.ts',
    'showPageScaleDialog',
    'pageScale',
    'block',
  ),
  manual('toolbar.zoomPrompt', 'toolbar/dialogs/zoom.ts', 'showZoomDialog', 'ribbon', 'block'),
  manual(
    'toolbar.mergeConfirm',
    'toolbar/dialogs/merge-confirm.ts',
    'confirmMergeLoseData',
    'ribbon',
    'block',
  ),
  manual(
    'toolbar.imagePicker',
    'toolbar/dialogs/image-file.ts',
    'pickImageFileDataUrl',
    'ribbon',
    'block',
    'unavailable',
    'Browser file picker requires a user gesture and cannot be driven safely in a unit test.',
  ),
  manual(
    'interact.localeSwitcher',
    'interact/locale-switcher.ts',
    'attachLocaleSwitcher',
    'localeSwitcher',
    'live',
  ),
  manual(
    'interact.macFormulaPalette',
    'interact/mac-formula-palette.ts',
    'attachMacFormulaPalette',
    'fxDialog',
    'block',
  ),
] as const;

export const LOCALE_SURFACE_FACTORIES: readonly LocaleSurfaceFactory[] = [
  ...LOCALE_SURFACE_SHELL_FACTORIES,
  ...LOCALE_SURFACE_MANUAL_FACTORIES,
];

export const localeSurfaceSectionSet = new Set<string>(LOCALE_SURFACE_SECTIONS);
