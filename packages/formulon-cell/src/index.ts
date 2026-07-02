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
export { type AutoSumFunction, autoSum } from './commands/auto-sum.js';
export type { DeleteCellsDirection, InsertCellsDirection } from './commands/cell-shift.js';
export { deleteCells, insertCells } from './commands/cell-shift.js';
export type {
  CellStyleDef,
  CellStyleGroupDef,
  CellStyleGroupId,
  CellStyleId,
  MergeCellStylesResult,
} from './commands/cell-styles.js';
export {
  applyCellStyle,
  applyCellStyleByName,
  CELL_STYLE_GROUPS,
  CELL_STYLES,
  createCellStyleFromActiveFormat,
  customCellStyleById,
  customCellStyleId,
  getCellStyle,
  listCustomCellStyles,
  mergeCellStylesFromWorkbook,
} from './commands/cell-styles.js';
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
export type { CoercedInput, CoerceInputOptions } from './commands/coerce-input.js';
export {
  coerceInput,
  coerceInputForCell,
  writeCoerced,
  writeInput,
  writeInputValidated,
} from './commands/coerce-input.js';
export type { CommentEntry } from './commands/comment.js';
export {
  clearComment,
  commentAt,
  listComments,
  recordCommentChange,
  setComment,
} from './commands/comment.js';
export type { ConditionalPresetAction } from './commands/conditional-format.js';
export {
  addConditionalRule,
  applyConditionalPresetAction,
  clearConditionalRules,
  clearConditionalRulesInRange,
  conditionalRulesForRange,
  listConditionalRules,
  removeConditionalRuleAt,
} from './commands/conditional-format.js';
export {
  cellValueIsFormulaError,
  circleInvalidValidationData,
  circleInvalidValidationDataInSheet,
  clearIgnoredCellErrors,
  clearValidationCircles,
  formulaErrorCellsInRange,
  ignoreCellError,
  isCellErrorIgnored,
  recordIgnoredErrorsChange,
  recordValidationCirclesChange,
  restoreCellErrorIndicator,
  selectNextFormulaError,
  toggleCellErrorIgnored,
} from './commands/error-indicators.js';
export type {
  ExternalLinkKind,
  ExternalLinkRecord,
  ExternalLinksSummary,
} from './commands/external-links.js';
export { listExternalLinks, summarizeExternalLinks } from './commands/external-links.js';
export type {
  ExecuteRibbonFillActionDeps,
  FillOptions,
  RibbonFillAction,
} from './commands/fill.js';
export { executeRibbonFillAction, fillDestFor, fillRange } from './commands/fill.js';
export type {
  AdvancedFilterCopyOptions,
  ConditionFilterOp,
  ConditionFilterOptions,
  FilterPredicate,
} from './commands/filter.js';
export {
  applyAdvancedFilter,
  applyConditionFilter,
  applyFilter,
  applyFilterColumns,
  applyValueFilter,
  clearFilter,
  copyAdvancedFilterResult,
  distinctFilterItems,
  distinctValues,
  type FilterValueItem,
  filterBySelectedCellValue,
  filterValueKey,
  inferAutoFilterRange,
  reapplyFilters,
  recordFilterChange,
  setAutoFilter,
} from './commands/filter.js';
export type { FindMatch, FindOptions } from './commands/find.js';
export {
  applySubstitution,
  findAll,
  findNext,
  replaceAll,
  replaceOne,
} from './commands/find.js';
export type { FlashFillExample, FlashFillPattern } from './commands/flash-fill.js';
export {
  applyFlashFill,
  applyFlashFillPattern,
  inferFlashFillPattern,
} from './commands/flash-fill.js';
export type { BorderPreset } from './commands/format.js';
export {
  bumpDecimals,
  bumpIndent,
  clearFormat,
  clearVisualFormat,
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
  CustomTableStyle,
  FormatAsTableOptions,
  PivotTableStyleAssignment,
  TableOverlay,
  TableOverlayPatch,
  TableStyle,
  TableStyleSwatch,
} from './commands/format-as-table.js';
export {
  applyPivotTableStyleById,
  clearTable,
  clearTablesInRange,
  createPivotTableStyleFromActivePivot,
  createTableStyleFromActiveTable,
  customPivotTableStyleById,
  customPivotTableStyleId,
  customTableStyleById,
  customTableStyleId,
  DEFAULT_TABLE_COLOR,
  defaultTableOverlay,
  engineTableOverlays,
  formatAsTable,
  formatAsTableByStyleId,
  isBandedRow,
  isFirstCol,
  isHeaderRow,
  isLastCol,
  isTotalRow,
  listCustomPivotTableStyles,
  listCustomTableStyles,
  listTableOverlays,
  pivotTableStyleAssignment,
  removeTable,
  sessionTableOverlays,
  TABLE_STYLE_COLORS,
  tableForCell,
  tableOverlayAt,
  tableOverlayById,
  tableStyleSwatch,
  tableVariantFromOptions,
  updateTableOverlay,
  upsertTable,
} from './commands/format-as-table.js';
export type {
  ExecuteRibbonFindActionDeps,
  GoToScope,
  GoToSpecialKind,
  GoToSpecialValueFilters,
  RibbonFindAction,
  RibbonFindActionReport,
  RibbonFindActionReportItem,
  RibbonFindActionResult,
} from './commands/goto-special.js';
export {
  boundingRange,
  executeRibbonFindAction,
  findMatchingCells,
  selectionFromMatches,
} from './commands/goto-special.js';
export type {
  FormatSnapshot,
  HistoryEntry,
  LayoutSnapshot,
  MergesSnapshot,
  TablesSnapshot,
} from './commands/history.js';
export {
  applyChartsSnapshot,
  applyConditionalRulesSnapshot,
  applyFormatSnapshot,
  applyIllustrationsSnapshot,
  applyLayoutSnapshot,
  applyMergesSnapshot,
  applyPageSetupSnapshot,
  applySlicersSnapshot,
  applySparklineSnapshot,
  applyTableOverlaysSnapshot,
  canRedo,
  canUndo,
  captureChartsSnapshot,
  captureConditionalRulesSnapshot,
  captureFormatSnapshot,
  captureIllustrationsSnapshot,
  captureLayoutSnapshot,
  captureMergesSnapshot,
  capturePageSetupSnapshot,
  captureSlicersSnapshot,
  captureSparklineSnapshot,
  captureTableOverlaysSnapshot,
  History,
  recordChartsChange,
  recordConditionalRulesChange,
  recordFormatChange,
  recordIllustrationsChange,
  recordLayoutChange,
  recordMergesChange,
  recordMergesChangeWithEngine,
  recordPageSetupChange,
  recordSlicersChange,
  recordSparklineChange,
  recordTablesChange,
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
  mergeWillLoseData,
  stepWithMerge,
} from './commands/merge.js';
export type {
  CreateDefinedNamesSource,
  DefinedNameDeleteResult,
  DefinedNameEntry,
  DefinedNameMutationResult,
} from './commands/named-ranges.js';
export {
  createDefinedNamesFromSelection,
  deleteDefinedName,
  insertDefinedNameFormula,
  isValidDefinedName,
  listDefinedNames,
  recordDefinedNamesChange,
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
export type {
  HostPrinterDevice,
  HostPrinterPaperOption,
  MarginPreset,
  PageBreakAxis,
  PageSetupEntry,
  PageSetupPatch,
  PrinterProfile,
} from './commands/page-setup.js';
export {
  addPrintArea,
  applyPrinterProfileBounds,
  clearPrintArea,
  clearPrintableBounds,
  clearPrintTitles,
  insertManualPageBreak,
  listPageSetups,
  marginPresetOf,
  marginPresetValues,
  normalizePrinterProfile,
  normalizePrinterProfileId,
  normalizePrinterProfiles,
  pageSetupForSheet,
  printerProfilesFromHostDevices,
  removeManualPageBreak,
  resetManualPageBreaks,
  resetPageSetup,
  resolvePrinterProfileBounds,
  setFitToPages,
  setMarginPreset,
  setPageOrientation,
  setPageScale,
  setPageSetup,
  setPaperSize,
  setPrintArea,
  setPrintableBounds,
  setPrintGridlines,
  setPrintHeadings,
  setPrintTitleCols,
  setPrintTitleRows,
  togglePageOrientation,
} from './commands/page-setup.js';
export type {
  CreatePivotTableOptions,
  CreatePivotTableResult,
  ExecuteRibbonPivotTableActionDeps,
  PivotFieldItemVisibility,
  PivotSourceField,
  RibbonPivotTableAction,
  RibbonPivotTableActionResult,
  RibbonPivotTableActionStrings,
  RibbonPivotTableReport,
  RibbonPivotTableReportItem,
} from './commands/pivot-table.js';
export {
  createPivotTableFromRange,
  executeRibbonPivotTableAction,
  inferPivotFieldItems,
  inferPivotSourceFields,
} from './commands/pivot-table.js';
export type {
  BuildPrintDocumentOptions,
  PrintAreaBounds,
  PrintDocument,
  PrintSheetOptions,
} from './commands/print.js';
export {
  buildPrintDocument,
  colLetter,
  parsePrintArea,
  parsePrintAreas,
  parsePrintTitleCols,
  parsePrintTitleRows,
  printSheet,
} from './commands/print.js';
export type {
  AllowedEditRangeOptions,
  SheetProtectionOptions,
  WorkbookStructureProtectionOptions,
} from './commands/protection.js';
export {
  addAllowedEditRange,
  allowedEditRangesForSheet,
  clearAllowedEditRanges,
  gateProtection,
  isAddrInAllowedEditRange,
  isCellLocked,
  isCellWritable,
  isSheetProtected,
  isWorkbookStructureProtected,
  protectedSheetPassword,
  protectedSheetPasswordHash,
  protectedSheetPermissions,
  recordProtectionChange,
  setCellLocked,
  setProtectedSheet,
  setWorkbookStructureProtected,
  toggleProtectedSheet,
  toggleWorkbookStructureProtected,
  verifySheetProtectionPasswordHash,
  warnProtected,
  workbookStructurePassword,
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
  ExecuteRibbonClearActionDeps,
  RibbonClearAction,
} from './commands/ribbon-clear.js';
export { executeRibbonClearAction } from './commands/ribbon-clear.js';
export type {
  ExecuteRibbonCommentActionDeps,
  RibbonCommentAction,
} from './commands/ribbon-comment.js';
export { executeRibbonCommentAction } from './commands/ribbon-comment.js';
export type {
  ExecuteRibbonFormulaAuditingActionDeps,
  RibbonFormulaAuditingAction,
  RibbonFormulaAuditingActionResult,
  RibbonFormulaAuditingReport,
} from './commands/ribbon-formula-auditing.js';
export { executeRibbonFormulaAuditingAction } from './commands/ribbon-formula-auditing.js';
export type {
  ExecuteRibbonProtectionActionDeps,
  RibbonProtectionAction,
  RibbonProtectionReport,
} from './commands/ribbon-protection.js';
export { executeRibbonProtectionAction } from './commands/ribbon-protection.js';
export type {
  CreateRibbonChartFromSelectionOptions,
  CreateSessionChartOptions,
  RibbonChartAction,
  SessionChartPatch,
  SessionChartSeriesPoint,
} from './commands/session-chart.js';
export {
  clearSessionChart,
  clearSessionChartsInRange,
  createRibbonChartFromSelection,
  createSessionChart,
  inferRecommendedChartKind,
  listSessionCharts,
  sessionChartById,
  sessionChartSeries,
  sessionChartsForRange,
  updateSessionChart,
} from './commands/session-chart.js';
export type {
  CreateSessionImageOptions,
  CreateSessionShapeOptions,
  SessionIllustrationArrangeAction,
  SessionIllustrationPatch,
} from './commands/session-illustration.js';
export {
  arrangeSessionIllustration,
  clearSessionIllustration,
  createRibbonImageFromSelection,
  createRibbonShapeFromSelection,
  createSessionImage,
  createSessionShape,
  duplicateSessionIllustration,
  listSessionIllustrations,
  sessionIllustrationById,
  updateSessionIllustration,
} from './commands/session-illustration.js';
export {
  addSheet,
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
  recordSheetViewsChange,
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
export type {
  SortActiveColumnAutoOptions,
  SortDirection,
  SortOptions,
  SortRangeWithHistoryDeps,
} from './commands/sort.js';
export {
  inferSortHasHeader,
  removeDuplicates,
  sortActiveColumnAuto,
  sortRange,
  sortRangeWithHistory,
} from './commands/sort.js';
export type { SparklineEntry } from './commands/sparkline.js';
export {
  clearSparkline,
  clearSparklinesInRange,
  listSparklines,
  setSparkline,
  sparklineAt,
} from './commands/sparkline.js';
export {
  autofitColsWidth,
  autofitRowsHeight,
  deleteCols,
  deleteRows,
  hiddenInSelection,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  setColsWidth,
  setFreezePanes,
  setRowsHeight,
  setSheetZoom,
  showCols,
  showColsAroundSelection,
  showRows,
  showRowsAroundSelection,
} from './commands/structure.js';
export { applyTextScriptToRange } from './commands/text-script.js';
export { textToColumns } from './commands/text-to-columns.js';
export {
  addTraceArrow,
  clearTraceArrows,
  clearTraceArrowsByKind,
  recordTraceChange,
  traceDependents,
  tracePrecedents,
} from './commands/traces.js';
export type { ValidationOutcome } from './commands/validate.js';
export {
  clearValidationInRange,
  clearValidationInRangeWithEngine,
  resolveListValues,
  validateAgainst,
} from './commands/validate.js';
export {
  clearSheetBackgroundImage,
  setGridlinesVisible,
  setHeadingsVisible,
  setR1C1ReferenceStyle,
  setSheetBackgroundImage,
  setShowFormulas,
  setStatusAggregates,
  setWorkbookView,
  setZoomPercent,
  setZoomScale,
  toggleStatusAggregate,
} from './commands/view.js';
export {
  clearWatchedCells,
  isWatched,
  recordWatchesChange,
  setWatchWindowOpen,
  toggleWatchCell,
  unwatchCell,
  watchCell,
  watchRange,
  watchRanges,
} from './commands/watch.js';
export type {
  AutomaticColorOption,
  ColorPaletteHandle,
  ColorPaletteOptions,
  ThemeColorColumn,
} from './components/color-palette.js';
export {
  createColorPalette,
  normalizeHex,
  PALETTE_COLUMNS,
  STANDARD_COLORS,
  THEME_COLOR_COLUMNS,
} from './components/color-palette.js';
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
  findPivotTableAtCell,
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
export type { WorkbookHandleFeatureMethods } from './engine/workbook-handle-features.js';
export type { WorkbookHandlePivotMethods } from './engine/workbook-handle-pivot.js';
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
  ResolvedSpreadsheetUiOptions,
  SpreadsheetFeatureSwitches,
  SpreadsheetUiOptions,
  SpreadsheetUiProfile,
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
  resolveSpreadsheetUiOptions,
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
export {
  defaultStrings,
  dictionaries,
  dictionaryLocaleFor,
  en,
  ja,
  mergeStrings,
} from './i18n/strings.js';
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
  ConditionalApplyFormatControls,
  ConditionalApplyFormatLabels,
  ConditionalApplyFormatOptions,
} from './interact/conditional-apply-controls.js';
export {
  appendConditionalApplyFormatControls,
  applyPatchToConditionalApplyControls,
  applyPresetPatchToConditionalApplyControls,
  collectConditionalApplyPatch,
} from './interact/conditional-apply-controls.js';
export type {
  ConditionalDialogDeps,
  ConditionalDialogHandle,
  ConditionalDialogOpenOptions,
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
export type { RangePickerControlOptions } from './interact/range-picker-control.js';
export {
  attachRangePickerButton,
  updateRangePickerLabel,
} from './interact/range-picker-control.js';
export type {
  BuildRibbonAddInReportStrings,
  RibbonAddInAction,
  RibbonAddInReport,
  RibbonAddInReportItem,
} from './interact/ribbon-add-in-report.js';
export { buildRibbonAddInReport } from './interact/ribbon-add-in-report.js';
export type {
  ExecuteRibbonFilterDataActionDeps,
  RibbonFilterDataAction,
  RibbonFilterDataActionResult,
} from './interact/ribbon-filter-data.js';
export {
  executeRibbonFilterDataAction,
  toggleAutoFilterFromSelection,
} from './interact/ribbon-filter-data.js';
export type {
  ExecuteRibbonHyperlinkActionDeps,
  RibbonHyperlinkAction,
  RibbonHyperlinkActionResult,
  RibbonHyperlinkReport,
} from './interact/ribbon-hyperlink.js';
export { executeRibbonHyperlinkAction } from './interact/ribbon-hyperlink.js';
export type {
  ResolveRibbonPdfActionStrings,
  RibbonPdfAction,
  RibbonPdfActionResult,
  RibbonPdfReport,
  RibbonPdfReportItem,
} from './interact/ribbon-pdf-report.js';
export { resolveRibbonPdfAction } from './interact/ribbon-pdf-report.js';
export type { SessionChartLabels, SessionChartsHandle } from './interact/session-charts.js';
export { attachSessionCharts } from './interact/session-charts.js';
export type { SessionIllustrationsHandle } from './interact/session-illustrations.js';
export { attachSessionIllustrations } from './interact/session-illustrations.js';
export type { SlicerDeps, SlicerHandle } from './interact/slicer.js';
export { attachSlicer } from './interact/slicer.js';
export type { StatusBarDeps, StatusBarHandle } from './interact/status-bar.js';
export { attachStatusBar } from './interact/status-bar.js';
export type {
  ValidationAlertDeps,
  ValidationAlertHandle,
  ValidationAlertLabels,
  ValidationAlertMessage,
  ValidationListDeps,
  ValidationListHandle,
  ValidationPromptDeps,
  ValidationPromptHandle,
} from './interact/validation.js';
export {
  attachValidationAlert,
  attachValidationList,
  attachValidationPrompt,
} from './interact/validation.js';
export type { ViewToolbarDeps, ViewToolbarHandle } from './interact/view-toolbar.js';
export { attachViewToolbar } from './interact/view-toolbar.js';
export type { WatchPanelDeps, WatchPanelHandle } from './interact/watch-panel.js';
export { attachWatchPanel } from './interact/watch-panel.js';
export type {
  SpreadsheetCompatibilityReportItem,
  WorkbookObjectsPanelDeps,
  WorkbookObjectsPanelHandle,
} from './interact/workbook-objects.js';
export {
  attachWorkbookObjectsPanel,
  buildSpreadsheetCompatibilityReport,
  spreadsheetCompatibilityDetail,
  spreadsheetCompatibilityLabel,
  spreadsheetCompatibilityStatusLabel,
} from './interact/workbook-objects.js';
export type { DefaultDynamicDropdownsOptions } from './mount/dynamic-dropdowns-defaults.js';
export { createDefaultDynamicDropdownsCtx } from './mount/dynamic-dropdowns-defaults.js';
export type {
  MountOptions,
  MountToolbarOptions,
  RibbonDisplayMode,
  ScreenClipCapture,
  ScreenClipCaptureResult,
  ScreenClipResult,
  SpreadsheetInstance,
  ToolbarInstance,
  ToolbarInstanceRef,
} from './mount.js';
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
  ConditionalScalePoint,
  ConditionalSlice,
  CustomCellStyle,
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
  SessionIllustration,
  SessionShapeKind,
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
  WorkbookViewMode,
} from './store/store.js';
export { createSpreadsheetStore, defaultPageSetup, getPageSetup, mutators } from './store/store.js';
export type { ResolvedTheme } from './theme/resolve.js';
export { resolveTheme } from './theme/resolve.js';
export type {
  ConditionalFormatDialogStyle,
  ConditionalFormatStyleOption,
  ConditionalFormatStyleStrings,
} from './toolbar/dialogs/conditional-format-style.js';
export {
  applyConditionalStylePreview,
  conditionalStyleFromValue,
  conditionalStyleOptions,
  showConditionalFormatCustomStyleDialog,
} from './toolbar/dialogs/conditional-format-style.js';
export type {
  ReportDialogLabels,
  ReportItem,
  ReportOptions,
} from './toolbar/dialogs/report.js';
export { reportDialogLabels, showReport } from './toolbar/dialogs/report.js';
export type { FluentIconName } from './toolbar/fluent-icons.js';
export { FLUENT_ICON_PATHS, fluentIconPaths } from './toolbar/fluent-icons.js';
export type { IconName } from './toolbar/icon-paths.js';
export { ICON_PATHS } from './toolbar/icon-paths.js';
export type { DisabledReasonProjectionOptions } from './toolbar/menu-a11y.js';
export {
  focusMenuItem,
  handleMenuKeydown,
  prepareMenu,
  projectDisabledReason,
  projectDisabledState,
} from './toolbar/menu-a11y.js';
export {
  type BackstageMenuText,
  backstageMenuText,
  type ConditionalMenuText,
  conditionalMenuText,
  type PageScaleMenuText,
  pageScaleMenuText,
  type RibbonDisplayText,
  ribbonDisplayText,
  type ToolbarMenuText,
  toolbarMenuText,
  type ViewToggleMenuText,
  viewToggleMenuText,
} from './toolbar/menu-text.js';
export type { NumberFormatAction } from './toolbar/number-format.js';
export { numberFormatForAction } from './toolbar/number-format.js';
export type {
  ReviewCell,
  ReviewCellValue,
  RibbonReportItem,
  RibbonReportLang,
  ScriptCommand,
} from './toolbar/review-tools.js';
export {
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyTextScript,
  buildTranslationReviewItems,
  formatRibbonReport,
  parseScriptCommand,
  reviewCellsFromState,
} from './toolbar/review-tools.js';
export type {
  RibbonActivationEntry,
  RibbonActivationKind,
  RibbonActivationSpec,
} from './toolbar/ribbon/activation.js';
export {
  RIBBON_BORDERS_MENU_ID,
  RIBBON_DIALOG_COMMANDS,
  RIBBON_DISABLED_COMMANDS,
  RIBBON_DROPDOWN_COMMANDS,
  RIBBON_DROPDOWN_MENU_FOR_COMMAND,
  RIBBON_DYNAMIC_MENU_FIRST_COMMANDS,
  RIBBON_EXTERNAL_MENU_FIRST_COMMANDS,
  RIBBON_EXTERNAL_MENU_FOR_COMMAND,
  RIBBON_GALLERY_COMMANDS,
  RIBBON_MENU_FACTORY_FOR_COMMAND,
  RIBBON_MENU_FACTORY_KEYS,
  RIBBON_MENU_FIRST_COMMANDS,
  RIBBON_MENU_FOR_COMMAND,
  RIBBON_PRIMARY_ACTION_COMMANDS,
  RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS,
  RIBBON_PRIMARY_FACE_MENU_COMMANDS,
  RIBBON_SPLIT_BUTTON_COMMANDS,
  RIBBON_SPLIT_TOGGLE_COMMANDS,
  RIBBON_TOGGLE_COMMANDS,
  ribbonActivationCategories,
  ribbonActivationCommandIds,
  ribbonActivationEntries,
  ribbonActivationEntriesForCommands,
  ribbonActivationForCommand,
} from './toolbar/ribbon/activation.js';
export type {
  ApplyRibbonCommandDeps,
  RibbonHooks,
  RibbonRuntime,
  RibbonUiState,
} from './toolbar/ribbon/apply-ribbon-command.js';
export { applyRibbonCommand } from './toolbar/ribbon/apply-ribbon-command.js';
export type { AutofitCellFormat } from './toolbar/ribbon/autofit.js';
export { autofitColWidth, autofitRowHeight } from './toolbar/ribbon/autofit.js';
export type {
  BackstageAction,
  BackstageDeps,
  BackstageFactories,
  BackstageItem,
  BackstageRibbonText,
  BackstageText,
} from './toolbar/ribbon/backstage.js';
export {
  backstageCardItems,
  backstageNavItems,
  createBackstageFactories,
} from './toolbar/ribbon/backstage.js';
export type {
  BackstageTitleApi,
  BackstageTitleCtx,
  BackstageTitleShellText,
} from './toolbar/ribbon/backstage-title.js';
export { createBackstageTitle } from './toolbar/ribbon/backstage-title.js';
export type { BorderPreviewSide, BorderPreviewSpec } from './toolbar/ribbon/border-icons.js';
export {
  createBorderPreview,
  createLineSamplePreview,
  LINE_STYLES_ALL,
  SVG_NS,
} from './toolbar/ribbon/border-icons.js';
export type { BorderMenuApi, BorderMenuCtx } from './toolbar/ribbon/border-menu.js';
export { createBorderMenu } from './toolbar/ribbon/border-menu.js';
export type { CellFormatActionDeps } from './toolbar/ribbon/cell-format-action.js';
export { applyCellFormatAction } from './toolbar/ribbon/cell-format-action.js';
export type {
  RibbonBorderDrawMode,
  RibbonFormatMutator as RibbonCommandTableMutator,
  RibbonViewMode,
} from './toolbar/ribbon/command-tables.js';
export {
  RIBBON_BORDER_DRAW_MODES,
  RIBBON_DIALOG_OPENERS,
  RIBBON_FORMAT_MUTATORS,
  RIBBON_FUNCTION_ARG_OPENERS,
  RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS,
  RIBBON_VIEW_MODES,
  RIBBON_ZOOM_PRESETS,
} from './toolbar/ribbon/command-tables.js';
export type {
  CfFillStyle,
  ConditionalMenuActionDeps,
} from './toolbar/ribbon/conditional-menu-action.js';
export { applyConditionalMenuAction } from './toolbar/ribbon/conditional-menu-action.js';
export type {
  ControlDispatchApi,
  ControlDispatchCtx,
  RibbonFormatMutator as ControlDispatchMutator,
} from './toolbar/ribbon/control-dispatch.js';
export { createControlDispatch } from './toolbar/ribbon/control-dispatch.js';
export type {
  DynamicDropdownMenuRefresherKey,
  DynamicDropdownsApi,
  DynamicDropdownsCtx,
  RibbonDropdownSpec,
  UiTheme as DynamicDropdownsUiTheme,
} from './toolbar/ribbon/dynamic-dropdowns.js';
export {
  createDynamicDropdowns,
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS,
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS,
  DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS,
} from './toolbar/ribbon/dynamic-dropdowns.js';
export type {
  RibbonFillDirection,
  RibbonFillSeriesMode,
} from './toolbar/ribbon/fill-series.js';
export {
  fillSeriesSourceRange,
  inferFillSeriesDirection,
  makeFillSeriesRadio,
  selectedFillSeriesRadio,
  showFillSeriesDialog,
} from './toolbar/ribbon/fill-series.js';
export {
  COMMON_FONT_VALUES,
  FONT_SUBMENU_FAMILIES,
  isFontProbablyAvailable,
  isJapaneseFontName,
  RECENT_FONT_VALUES,
  shouldShowFontOption,
  THEME_FONT_VALUES,
} from './toolbar/ribbon/font-availability.js';
export type { BordersMenuDeps } from './toolbar/ribbon/menus/borders.js';
export { createBordersMenu } from './toolbar/ribbon/menus/borders.js';
export type { CfSubmenuKey } from './toolbar/ribbon/menus/conditional.js';
export { buildCfMenuText, createConditionalMenu } from './toolbar/ribbon/menus/conditional.js';
export type {
  AutoSumFormulaName,
  FormulasMenuFactories,
} from './toolbar/ribbon/menus/formulas.js';
export { createFormulasMenuFactories } from './toolbar/ribbon/menus/formulas.js';
export {
  createMenu,
  menuSectionHeader,
  menuSeparator,
} from './toolbar/ribbon/menus/general.js';
export type { HomeMenuDeps, HomeMenuFactories } from './toolbar/ribbon/menus/home.js';
export { createHomeMenuFactories } from './toolbar/ribbon/menus/home.js';
export type { InsertMenuFactories } from './toolbar/ribbon/menus/insert.js';
export { createInsertMenuFactories } from './toolbar/ribbon/menus/insert.js';
export type { PageLayoutMenuFactories } from './toolbar/ribbon/menus/page-layout.js';
export { createPageLayoutMenuFactories } from './toolbar/ribbon/menus/page-layout.js';
export { createPasteMenu } from './toolbar/ribbon/menus/paste.js';
export type { ReviewMenuFactories } from './toolbar/ribbon/menus/review.js';
export { createReviewMenuFactories } from './toolbar/ribbon/menus/review.js';
export type {
  StylesMenuDeps,
  StylesMenuFactories,
  TableVariantId,
} from './toolbar/ribbon/menus/styles.js';
export { createStylesMenuFactories, tableVariantOptions } from './toolbar/ribbon/menus/styles.js';
export type { TextOrientationGlyph } from './toolbar/ribbon/menus/text-orientation.js';
export { createTextOrientationMenu } from './toolbar/ribbon/menus/text-orientation.js';
export type {
  RenderRibbonApi,
  RenderRibbonCtx,
  RibbonMenuFactory,
  RibbonMenus,
  RibbonRenderHelpers,
  RibbonRenderState,
} from './toolbar/ribbon/render-ribbon.js';
export {
  createRenderRibbon,
  LEGACY_COMMAND_IDS,
  SPLIT_BUTTON_COMMANDS,
} from './toolbar/ribbon/render-ribbon.js';
export type {
  BuildRibbonSearchIndexOptions,
  RibbonSearchItem,
  RibbonSearchItemKind,
  RibbonSearchUsagePrior,
} from './toolbar/ribbon/search-index.js';
export {
  buildRibbonSearchIndex,
  queryRibbonSearchIndex,
} from './toolbar/ribbon/search-index.js';
export type {
  SelectColorApi,
  SelectColorCtx,
  SelectColorPageScaleText,
  SelectColorRibbonText,
} from './toolbar/ribbon/select-color.js';
export { createSelectColorRibbon } from './toolbar/ribbon/select-color.js';
export type { ActiveState } from './toolbar/ribbon-active-state.js';
export {
  BORDER_PRESETS,
  BORDER_STYLES,
  EMPTY_ACTIVE_STATE,
  localizeBorderPresets,
  localizeBorderStyles,
  projectActiveState,
  RIBBON_ACTIVE_COMMANDS,
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
  EXCEL365_STANDARD_RIBBON_TABS,
  FONT_FAMILIES,
  FONT_SIZES,
  HOME_MIXED_LAYOUT_GROUP_VARIANTS,
  HOME_STACKED_LAYOUT_GROUP_VARIANTS,
  HOME_TILE_LAYOUT_GROUP_VARIANTS,
  isRibbonActivatableCommand,
  OPTIONAL_RIBBON_TABS,
  RIBBON_KEYSHORTCUTS,
  RIBBON_TAB_LABELS,
  RIBBON_TABS,
  ribbonActivatableCommandIds,
  ribbonActivatableCommands,
  ribbonActivatableSurfaceCommandIds,
  ribbonActivatableSurfaceCommands,
  ribbonCommandIds,
  ribbonCommands,
  ribbonSurfaceCommandIds,
  ribbonSurfaceCommands,
  ribbonTabCommandIds,
  ribbonTabLabel,
  toolbarText,
} from './toolbar/ribbon-model.js';
export type {
  ConditionalIconSetAction,
  ToolbarInsertSymbol,
} from './wrappers/conditional-menu-labels.js';
export {
  conditionalColorScaleLabel,
  conditionalColorScaleSwatchColors,
  conditionalDataBarLabel,
  conditionalDataBarSwatchColor,
  conditionalIconSetLabel,
  TOOLBAR_INSERT_SYMBOLS,
} from './wrappers/conditional-menu-labels.js';
export { cellLabel, formatA1Range, parseA1Atom, parseA1Range } from './wrappers/toolbar-a1.js';
export type {
  AutoSumAction,
  CellDeleteAction,
  CellInsertAction,
  ConditionalMenuAction,
  FreezeAction,
  MergeAction,
  PasteAction,
  WindowAction,
} from './wrappers/toolbar-actions.js';
export {
  deleteSelectedCols,
  deleteSelectedRows,
  dispatchHostClipboard,
  handleAutoSum,
  handleAutoSumAction,
  handleConditionalAction,
  handleDeleteCellsAction,
  handleFreezeAction,
  handleInsertCellsAction,
  handleMergeAction,
  handlePasteAction,
  handleWindowAction,
  insertSelectedCols,
  insertSelectedRows,
  toggleSelectedColsHidden,
  toggleSelectedRowsHidden,
} from './wrappers/toolbar-actions.js';
export type {
  AddInAction,
  AdvancedFilterDialogDraft,
  AutomationRunDraft,
  CalculationAction,
  CellFormatAction,
  CellStyleAction,
  ChartAction,
  ClearAction,
  ClearArrowsAction,
  CommentAction,
  DataValidationAction,
  DefinedNameAction,
  DimensionDialogDraft,
  FillAction,
  FilterDataAction,
  FindAction,
  FormatTableAction,
  FormulaAuditingAction,
  FunctionAction,
  HyperlinkAction,
  OutlineAxisAction,
  PageBreakAction,
  PdfAction,
  PictureAction,
  PivotTableAction,
  PrintAreaAction,
  PrintTitleAction,
  ProtectionAction,
  RemoveDuplicatesDialogDraft,
  RibbonReportDialogDraft,
  ScreenshotAction,
  ScriptDialogDraft,
  ShapeAction,
  SheetBackgroundAction,
  SheetCell,
  SheetCellFor,
  SheetRange,
  SheetRangeFor,
  SheetRenameDialogDraft,
  SortAction,
  SortDialogDraft,
  SymbolAction,
  TextOrientationAction,
  TextToColumnsAction,
  TextToColumnsDialogDraft,
  ThemeAction,
  WatchAction,
} from './wrappers/toolbar-types.js';
export {
  CELL_STYLE_SECTION_ACTION_PREFIX,
  MORE_SYMBOL_ACTION,
  SHEET_TAB_COLOR_ACTIONS,
  TEXT_TO_COLUMNS_DIALOG_KEYS,
} from './wrappers/toolbar-types.js';
