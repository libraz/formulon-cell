// Cell renderer / editor registry — `inst.cells.registerFormatter(...)`.
export type {
  CellEditorEntry,
  CellEditorHandle,
  CellFormatterEntry,
  CellRenderInput,
} from './cells.js';
export { CellRegistry } from './cells.js';
export type { SelectionStats, StatusAggregateEntry } from './commands/aggregate.js';
export {
  aggregateSelection,
  countUniqueRangeCells,
  STATUS_AGGREGATE_KEYS,
  statusAggregateValue,
  visibleStatusAggregates,
} from './commands/aggregate.js';
export { autoSum } from './commands/auto-sum.js';
export type { CellStyleDef, CellStyleId } from './commands/cell-styles.js';
export { applyCellStyle, CELL_STYLES, getCellStyle } from './commands/cell-styles.js';
export type { CopyResult } from './commands/clipboard/copy.js';
export { copy } from './commands/clipboard/copy.js';
export type { CSVEncodeOptions } from './commands/clipboard/csv.js';
export { encodeCSV, parseCSV } from './commands/clipboard/csv.js';
export { cut } from './commands/clipboard/cut.js';
export { encodeHtml } from './commands/clipboard/html.js';
export type {
  InsertCopiedCellsDirection,
  InsertCopiedCellsResult,
} from './commands/clipboard/insert-copied-cells.js';
export { insertCopiedCellsFromTSV } from './commands/clipboard/insert-copied-cells.js';
export type { PasteResult } from './commands/clipboard/paste.js';
export { pasteTSV } from './commands/clipboard/paste.js';
export type {
  PasteOperation,
  PasteSpecialOptions,
  PasteSpecialResult,
  PasteWhat,
} from './commands/clipboard/paste-special.js';
export { pasteSpecial as applyPasteSpecial } from './commands/clipboard/paste-special.js';
export type {
  ClipboardCell,
  ClipboardSnapshot,
} from './commands/clipboard/snapshot.js';
export { captureSnapshot } from './commands/clipboard/snapshot.js';
export { encodeTSV, parseTSV } from './commands/clipboard/tsv.js';
export type { CoercedInput } from './commands/coerce-input.js';
export {
  coerceInput,
  writeCoerced,
  writeInput,
  writeInputValidated,
} from './commands/coerce-input.js';
export type { CommentEntry } from './commands/comment.js';
export { clearComment, commentAt, listComments, setComment } from './commands/comment.js';
export {
  addConditionalRule,
  clearConditionalRules,
  clearConditionalRulesInRange,
  conditionalRulesForRange,
  listConditionalRules,
  removeConditionalRuleAt,
} from './commands/conditional-format.js';
export {
  clearIgnoredCellErrors,
  ignoreCellError,
  isCellErrorIgnored,
  restoreCellErrorIndicator,
  toggleCellErrorIgnored,
} from './commands/error-indicators.js';
export type {
  ExternalLinkKind,
  ExternalLinkRecord,
  ExternalLinksSummary,
} from './commands/external-links.js';
export { listExternalLinks, summarizeExternalLinks } from './commands/external-links.js';
export type { FillOptions } from './commands/fill.js';
export { fillDestFor, fillRange } from './commands/fill.js';
export type { FilterPredicate } from './commands/filter.js';
export { applyFilter, clearFilter, distinctValues, setAutoFilter } from './commands/filter.js';
export type { FindMatch, FindOptions } from './commands/find.js';
export {
  applySubstitution,
  findAll,
  findNext,
  replaceAll,
  replaceOne,
} from './commands/find.js';
export type { BorderPreset } from './commands/format.js';
export {
  bumpDecimals,
  bumpIndent,
  clearFormat,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  formatNumber,
  setAlign,
  setBorderPreset,
  setBorders,
  setFillColor,
  setFont,
  setFontColor,
  setNumFmt,
  setRotation,
  setVAlign,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
} from './commands/format.js';
export type {
  FormatAsTableOptions,
  TableOverlay,
  TableOverlayPatch,
  TableStyle,
} from './commands/format-as-table.js';
export {
  clearTable,
  clearTablesInRange,
  defaultTableOverlay,
  engineTableOverlays,
  formatAsTable,
  isBandedRow,
  isHeaderRow,
  isTotalRow,
  listTableOverlays,
  removeTable,
  sessionTableOverlays,
  tableForCell,
  tableOverlayAt,
  tableOverlayById,
  updateTableOverlay,
  upsertTable,
} from './commands/format-as-table.js';
export type { GoToScope, GoToSpecialKind } from './commands/goto-special.js';
export { boundingRange, findMatchingCells } from './commands/goto-special.js';
export type { HistoryEntry, LayoutSnapshot, MergesSnapshot } from './commands/history.js';
export {
  applyFormatSnapshot,
  applyLayoutSnapshot,
  applyMergesSnapshot,
  applyPageSetupSnapshot,
  applySlicersSnapshot,
  applySparklineSnapshot,
  canRedo,
  canUndo,
  captureFormatSnapshot,
  captureLayoutSnapshot,
  captureMergesSnapshot,
  capturePageSetupSnapshot,
  captureSlicersSnapshot,
  captureSparklineSnapshot,
  History,
  recordFormatChange,
  recordLayoutChange,
  recordMergesChange,
  recordMergesChangeWithEngine,
  recordPageSetupChange,
  recordSlicersChange,
  recordSparklineChange,
  redo,
  undo,
} from './commands/history.js';
export type { HyperlinkEntry } from './commands/hyperlinks.js';
export {
  clearHyperlink,
  hyperlinkAt,
  listEngineHyperlinks,
  listHyperlinks,
  setHyperlink,
} from './commands/hyperlinks.js';
export type { ExportOptions, ImportResult } from './commands/import-export.js';
export { exportCSV, importCSV } from './commands/import-export.js';
export {
  applyMerge,
  applyUnmerge,
  expandRangeWithMerges,
  mergeAnchorOf,
  mergeAt,
  stepWithMerge,
} from './commands/merge.js';
export type {
  DefinedNameDeleteResult,
  DefinedNameEntry,
  DefinedNameMutationResult,
} from './commands/named-ranges.js';
export {
  deleteDefinedName,
  listDefinedNames,
  upsertDefinedName,
} from './commands/named-ranges.js';
export {
  colGroupRangeAt,
  collapseColGroup,
  collapseRowGroup,
  expandColGroup,
  expandRowGroup,
  groupCols,
  groupRows,
  isColGroupCollapsed,
  isRowGroupCollapsed,
  MAX_OUTLINE_LEVEL,
  OUTLINE_GUTTER_PER_LEVEL,
  rowGroupRangeAt,
  ungroupCols,
  ungroupRows,
} from './commands/outline.js';
export type { MarginPreset, PageSetupEntry, PageSetupPatch } from './commands/page-setup.js';
export {
  clearPrintTitles,
  listPageSetups,
  marginPresetOf,
  marginPresetValues,
  pageSetupForSheet,
  resetPageSetup,
  setMarginPreset,
  setPageOrientation,
  setPageSetup,
  setPaperSize,
  setPrintTitleCols,
  setPrintTitleRows,
  togglePageOrientation,
} from './commands/page-setup.js';
export type {
  CreatePivotTableOptions,
  CreatePivotTableResult,
  PivotSourceField,
} from './commands/pivot-table.js';
export { createPivotTableFromRange, inferPivotSourceFields } from './commands/pivot-table.js';
export type { PrintDocument } from './commands/print.js';
export {
  buildPrintDocument,
  colLetter,
  parsePrintTitleCols,
  parsePrintTitleRows,
  printSheet,
} from './commands/print.js';
export type { SheetProtectionOptions } from './commands/protection.js';
export {
  gateProtection,
  isCellLocked,
  isCellWritable,
  isSheetProtected,
  protectedSheetPassword,
  setCellLocked,
  setProtectedSheet,
  toggleProtectedSheet,
  warnProtected,
  writableAddrs,
} from './commands/protection.js';
export type {
  QuickAnalysisAction,
  QuickAnalysisActionId,
  QuickAnalysisExecuteInput,
  QuickAnalysisExecuteResult,
  QuickAnalysisGroup,
  QuickAnalysisInput,
} from './commands/quick-analysis.js';
export {
  buildQuickAnalysisActions,
  enabledQuickAnalysisActions,
  executeQuickAnalysisAction,
  groupQuickAnalysisActions,
  isQuickAnalysisActionEnabled,
  quickAnalysisActionById,
} from './commands/quick-analysis.js';
export type { ActiveSignature, F4Result, FormulaRef } from './commands/refs.js';
export {
  extractRefs,
  FUNCTION_SIGNATURES,
  findActiveSignature,
  REF_HIGHLIGHT_COLORS,
  rotateRefAt,
  shiftFormulaRefs,
  suggestFunctions,
} from './commands/refs.js';
export type {
  CreateSessionChartOptions,
  SessionChartPatch,
  SessionChartSeriesPoint,
} from './commands/session-chart.js';
export {
  clearSessionChart,
  clearSessionChartsInRange,
  createSessionChart,
  listSessionCharts,
  sessionChartById,
  sessionChartSeries,
  sessionChartsForRange,
  updateSessionChart,
} from './commands/session-chart.js';
export {
  moveSheet,
  removeSheet,
  renameSheet,
  setSheetHidden,
} from './commands/sheet-mutate.js';
export type {
  SheetView,
  SheetViewPatch,
  SheetViewSnapshotInput,
  SheetViewSort,
  SheetViewStoreResult,
} from './commands/sheet-views.js';
export {
  activateSheetView,
  applySheetView,
  captureSheetView,
  deleteSheetView,
  findSheetView,
  removeSheetView,
  saveSheetView,
  upsertSheetView,
} from './commands/sheet-views.js';
export type { CreateSlicerOptions, CreateSlicerResult } from './commands/slicers.js';
export {
  clearSlicerSelection,
  createSlicer,
  findSlicerTable,
  listSlicers,
  listSlicerValues,
  recomputeSlicerFilters,
  removeSlicer,
  resolveSlicerSpec,
  setSlicerSelected,
  updateSlicer,
} from './commands/slicers.js';
export type { SortDirection, SortOptions } from './commands/sort.js';
export { removeDuplicates, sortRange } from './commands/sort.js';
export type { SparklineEntry } from './commands/sparkline.js';
export {
  clearSparkline,
  clearSparklinesInRange,
  listSparklines,
  setSparkline,
  sparklineAt,
} from './commands/sparkline.js';
export {
  deleteCols,
  deleteRows,
  hiddenInSelection,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  setFreezePanes,
  setSheetZoom,
  showCols,
  showRows,
} from './commands/structure.js';
export { textToColumns } from './commands/text-to-columns.js';
export {
  addTraceArrow,
  clearTraceArrows,
  traceDependents,
  tracePrecedents,
} from './commands/traces.js';
export type { ValidationOutcome } from './commands/validate.js';
export { resolveListValues, validateAgainst } from './commands/validate.js';
export {
  setGridlinesVisible,
  setHeadingsVisible,
  setR1C1ReferenceStyle,
  setShowFormulas,
  setStatusAggregates,
  setZoomPercent,
  setZoomScale,
  toggleStatusAggregate,
} from './commands/view.js';
export {
  clearWatchedCells,
  isWatched,
  setWatchWindowOpen,
  toggleWatchCell,
  unwatchCell,
  watchCell,
} from './commands/watch.js';
export type { NamedCellStyle } from './engine/cell-styles-meta.js';
export { computeNamedCellStyles } from './engine/cell-styles-meta.js';
export type {
  SpreadsheetCompatibilityId,
  SpreadsheetCompatibilityItem,
  SpreadsheetCompatibilityStatus,
  SpreadsheetCompatibilitySummary,
} from './engine/compatibility.js';
export {
  isSpreadsheetFeatureAvailable,
  isSpreadsheetFeatureWritable,
  spreadsheetCompatibilityItem,
  spreadsheetCompatibilityStatus,
  summarizeSpreadsheetCompatibility,
} from './engine/compatibility.js';
export type { LocaleOrdinal } from './engine/function-locale.js';
export { canonicalizeFormula, localizeFormula } from './engine/function-locale.js';
export type { LoadOptions } from './engine/loader.js';
export { isUsingStub } from './engine/loader.js';
export type {
  PassthroughSummary,
  PivotTableSummary,
  TableSummary,
  WorkbookObjectKind,
  WorkbookObjectRecord,
} from './engine/passthrough-sync.js';
export {
  classifyWorkbookObjectPath,
  listWorkbookObjects,
  summarizePassthroughs,
  summarizePivotTables,
  summarizeTables,
  WORKBOOK_OBJECT_KINDS,
  workbookObjectExtension,
  workbookObjectKindCounts,
  workbookObjectKindLabel,
  workbookObjectName,
  workbookObjectsByKind,
} from './engine/passthrough-sync.js';
export type { RangeResolver } from './engine/range-resolver.js';
export {
  isRangeSource,
  makeRangeResolver,
  parseRangeRef,
  resolveNumericRange,
  resolveNumericRangeFromCells,
  resolveRangeRef,
} from './engine/range-resolver.js';
export { findDependents, findPrecedents } from './engine/refs-graph.js';
export { hydrateTableOverlaysFromEngine, tableOverlaysFromEngine } from './engine/table-sync.js';
export type {
  Addr,
  CellValue,
  EngineCapabilities,
  PivotCell,
  PivotDataFieldSpec,
  PivotFieldSpec,
  PivotFilterSpec,
  PivotLayoutResult,
  Range,
  SpreadsheetProfileId,
} from './engine/types.js';
export {
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotDateGrouping,
  PivotFilterType,
  PivotFilterValueKind,
  PivotShowValuesAs,
} from './engine/types.js';
export { formatCell, fromEngineValue } from './engine/value.js';
export type { ChangeEvent, ChangeListener } from './engine/workbook-handle.js';
export { WorkbookHandle } from './engine/workbook-handle.js';
// Public event surface — adapter packages and direct consumers wire to
// these via `inst.on(...)`.
export type {
  CellChangeEvent,
  LocaleChangeEvent,
  RecalcEvent,
  SelectionChangeEvent,
  SpreadsheetEventHandler,
  SpreadsheetEventName,
  SpreadsheetEvents,
  ThemeChangeEvent,
  WorkbookChangeEvent,
} from './events.js';
// Extensions / feature gating — public composition surface.
export type {
  Extension,
  ExtensionContext,
  ExtensionHandle,
  ExtensionInput,
  FeatureFlags,
  FeatureId,
  I18nController,
  ThemeName,
} from './extensions/index.js';
export {
  ALL_FEATURE_IDS,
  allBuiltIns,
  charts,
  clipboard,
  conditionalDialog,
  contextMenu,
  dedupeById,
  findReplace,
  flattenExtensions,
  formatDialog,
  formatPainter,
  full,
  goToSpecialDialog,
  hoverComment,
  hyperlinkDialog,
  iterativeDialog,
  minimal,
  namedRangeDialog,
  pageSetupDialog,
  pasteSpecial,
  pivotTableDialog,
  presets,
  quickAnalysis,
  resolveFlags,
  slicer,
  sortByPriority,
  standard,
  statusBar,
  validationList,
  viewToolbar,
  watchWindow,
  wheel,
  workbookObjects,
} from './extensions/index.js';
// Custom function registry — `inst.formula.register(name, impl)`.
export type {
  CustomFunction,
  CustomFunctionMeta,
  CustomFunctionReturn,
} from './formula.js';
export { FormulaRegistry } from './formula.js';
export type { I18nControllerInit } from './i18n/controller.js';
export { createI18nController } from './i18n/controller.js';
export type { DeepPartial, Locale, Strings } from './i18n/strings.js';
export { defaultStrings, dictionaries, en, ja, mergeStrings } from './i18n/strings.js';
export type { ArgHelperDeps, ArgHelperHandle, ArgHelperLabels } from './interact/arg-helper.js';
export { attachArgHelper } from './interact/arg-helper.js';
export type {
  AutocompleteDeps,
  AutocompleteHandle,
  AutocompleteLabels,
  AutocompleteTable,
} from './interact/autocomplete.js';
export { attachAutocomplete } from './interact/autocomplete.js';
export type {
  BorderDrawDeps,
  BorderDrawHandle,
  BorderDrawMode,
} from './interact/border-draw.js';
export { attachBorderDraw } from './interact/border-draw.js';
export type {
  CellStylesGalleryDeps,
  CellStylesGalleryHandle,
} from './interact/cell-styles-gallery.js';
export { attachCellStylesGallery } from './interact/cell-styles-gallery.js';
export type { CfRulesDialogDeps, CfRulesDialogHandle } from './interact/cf-rules-dialog.js';
export { attachCfRulesDialog } from './interact/cf-rules-dialog.js';
export type { ClipboardHandle } from './interact/clipboard.js';
export { attachClipboard } from './interact/clipboard.js';
export type {
  ConditionalDialogDeps,
  ConditionalDialogHandle,
} from './interact/conditional-dialog.js';
export { attachConditionalDialog } from './interact/conditional-dialog.js';
export type { ContextMenuDeps } from './interact/context-menu.js';
export { attachContextMenu } from './interact/context-menu.js';
export type {
  ErrorMenuDeps,
  ErrorMenuHandle,
  ErrorMenuKind,
} from './interact/error-menu.js';
export { attachErrorMenu } from './interact/error-menu.js';
export type {
  ExternalLinksDialogDeps,
  ExternalLinksDialogHandle,
} from './interact/external-links-dialog.js';
export { attachExternalLinksDialog } from './interact/external-links-dialog.js';
export type { FilterDropdownDeps, FilterDropdownHandle } from './interact/filter-dropdown.js';
export { attachFilterDropdown } from './interact/filter-dropdown.js';
export type { FindReplaceDeps, FindReplaceHandle } from './interact/find-replace.js';
export { attachFindReplace } from './interact/find-replace.js';
export type { FormatDialogDeps, FormatDialogHandle } from './interact/format-dialog.js';
export { attachFormatDialog } from './interact/format-dialog.js';
export type { FormatPainterDeps, FormatPainterHandle } from './interact/format-painter.js';
export { attachFormatPainter } from './interact/format-painter.js';
export type { FxDialogDeps, FxDialogHandle } from './interact/fx-dialog.js';
export { attachFxDialog, FUNCTION_DESCRIPTIONS } from './interact/fx-dialog.js';
export type { GoToDialogDeps, GoToDialogHandle } from './interact/goto-dialog.js';
export { attachGoToDialog } from './interact/goto-dialog.js';
export type { HoverDeps, HoverHandle } from './interact/hover.js';
export { attachHover } from './interact/hover.js';
export type {
  IterativeDialogDeps,
  IterativeDialogHandle,
} from './interact/iterative-dialog.js';
export { attachIterativeDialog } from './interact/iterative-dialog.js';
export type {
  NamedRangeDialogDeps,
  NamedRangeDialogHandle,
} from './interact/named-range-dialog.js';
export { attachNamedRangeDialog } from './interact/named-range-dialog.js';
export type {
  PageSetupDialogDeps,
  PageSetupDialogHandle,
} from './interact/page-setup-dialog.js';
export { attachPageSetupDialog } from './interact/page-setup-dialog.js';
export type { PasteSpecialDeps, PasteSpecialHandle } from './interact/paste-special.js';
export { attachPasteSpecial } from './interact/paste-special.js';
export type { QuickAnalysisDeps, QuickAnalysisHandle } from './interact/quick-analysis.js';
export { attachQuickAnalysis } from './interact/quick-analysis.js';
export type { SessionChartLabels, SessionChartsHandle } from './interact/session-charts.js';
export { attachSessionCharts } from './interact/session-charts.js';
export type { SlicerDeps, SlicerHandle } from './interact/slicer.js';
export { attachSlicer } from './interact/slicer.js';
export type { StatusBarDeps, StatusBarHandle } from './interact/status-bar.js';
export { attachStatusBar } from './interact/status-bar.js';
export type { ValidationListDeps, ValidationListHandle } from './interact/validation.js';
export { attachValidationList } from './interact/validation.js';
export type { ViewToolbarDeps, ViewToolbarHandle } from './interact/view-toolbar.js';
export { attachViewToolbar } from './interact/view-toolbar.js';
export type { WatchPanelDeps, WatchPanelHandle } from './interact/watch-panel.js';
export { attachWatchPanel } from './interact/watch-panel.js';
export type {
  WorkbookObjectsPanelDeps,
  WorkbookObjectsPanelHandle,
} from './interact/workbook-objects.js';
export { attachWorkbookObjectsPanel } from './interact/workbook-objects.js';
export type { MountOptions, SpreadsheetInstance } from './mount.js';
export { Spreadsheet } from './mount.js';
export type { ErrorTriangleHit, ErrorTriangleKind } from './render/grid.js';
export {
  detectErrorKind,
  detectValidationViolation,
  ERROR_TRIANGLE_COLOR,
  getErrorTriangleHits,
  VALIDATION_TRIANGLE_COLOR,
} from './render/grid.js';
export type {
  CellAlign,
  CellBorderSide,
  CellBorderStyle,
  CellBorders,
  CellFormat,
  CellVAlign,
  CellValidation,
  ChartsSlice,
  ConditionalIconSet,
  ConditionalRule,
  ConditionalSlice,
  ErrorIndicatorSlice,
  FormatSlice,
  MergesSlice,
  NegativeStyle,
  NumFmt,
  PageMargins,
  PageOrientation,
  PageSetup,
  PageSetupSlice,
  PaperSize,
  ProtectionSlice,
  SessionChart,
  SessionChartKind,
  SlicerSpec,
  SlicersSlice,
  Sparkline,
  SparklineKind,
  SparklineSlice,
  SpreadsheetStore,
  State,
  StatusAggKey,
  TraceArrow,
  TracesSlice,
  ValidationErrorStyle,
  ValidationListSource,
  ValidationMeta,
  ValidationOp,
  WatchSlice,
} from './store/store.js';
export { createSpreadsheetStore, defaultPageSetup, getPageSetup, mutators } from './store/store.js';
export type { ResolvedTheme } from './theme/resolve.js';
export { resolveTheme } from './theme/resolve.js';
export type { FluentIconName } from './toolbar/fluent-icons.js';
export { FLUENT_ICON_PATHS, fluentIconPaths } from './toolbar/fluent-icons.js';
export type {
  ReviewCell,
  ReviewCellValue,
  RibbonReportItem,
  ScriptCommand,
} from './toolbar/review-tools.js';
export {
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyTextScript,
  parseScriptCommand,
} from './toolbar/review-tools.js';
export type { ActiveState } from './toolbar/ribbon-active-state.js';
export {
  BORDER_PRESETS,
  BORDER_STYLES,
  EMPTY_ACTIVE_STATE,
  projectActiveState,
} from './toolbar/ribbon-active-state.js';
export type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonOption,
  RibbonTab,
  RibbonTabModel,
  ToolbarLang,
  ToolbarText,
} from './toolbar/ribbon-model.js';
export {
  buildRibbonModel,
  FONT_FAMILIES,
  FONT_SIZES,
  RIBBON_KEYSHORTCUTS,
  RIBBON_TAB_LABELS,
  toolbarText,
} from './toolbar/ribbon-model.js';
