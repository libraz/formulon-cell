// Cell renderer registry — `inst.cells.registerFormatter(...)`.
export type { CellFormatterEntry, CellRenderInput } from './cells.js';
export { CellRegistry } from './cells.js';
export {
  formatChartNumber,
  formatChartValue,
  layoutChart,
  layoutUnsupportedChart,
} from './charts/index.js';
export type {
  CategoryChartData,
  CategoryChartSeries,
  ChartAxisSpec,
  ChartData,
  ChartDataLabelSpec,
  ChartDataLabelsSpec,
  ChartEmptyCellMode,
  ChartFillSpec,
  ChartFilterSpec,
  ChartGrouping,
  ChartKind,
  ChartLabelPosition,
  ChartLegendSpec,
  ChartLineSpec,
  ChartMarkerSpec,
  ChartPlotBounds,
  ChartPointState,
  ChartPointStyleSpec,
  ChartPrimitive,
  ChartPrimitiveRole,
  ChartPrimitiveTag,
  ChartScatterStyle,
  ChartScene,
  ChartSeriesDirection,
  ChartSeriesStyleSpec,
  ChartSpec,
  ChartValueFormatter,
  XYChartData,
  XYChartPoint,
  XYChartSeries,
} from './charts/types.js';
export type { SelectionStats, StatusAggregateEntry } from './commands/aggregate.js';
export {
  countUniqueRangeCells,
  STATUS_AGGREGATE_KEYS,
  statusAggregateValue,
  visibleStatusAggregates,
} from './commands/aggregate.js';
export { type AutoSumFunction, autoSum } from './commands/auto-sum.js';
export type { AutofitOptions } from './commands/autofit-measurement.js';
export { autofitColWidth, autofitRowHeight } from './commands/autofit-measurement.js';
export type { BorderPreset } from './commands/border-presets.js';
export type { CopyResult } from './commands/clipboard/copy.js';
export { copy } from './commands/clipboard/copy.js';
export type { CSVDelimiter, CSVEncodeOptions, CSVParseOptions } from './commands/clipboard/csv.js';
export { cut } from './commands/clipboard/cut.js';
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
export type { CommentEntry } from './commands/comment.js';
export { listComments } from './commands/comment.js';
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
export type { RibbonFillAction } from './commands/fill.js';
export type { FindMatch, FindOptions } from './commands/find.js';
export { findAll, replaceAll } from './commands/find.js';
export { clearFormat, formatNumber } from './commands/format.js';
export type {
  CustomTableStyle,
  FormatAsTableOptions,
  TableColumnTotals,
  TableNameError,
  TableOverlay,
  TableOverlayPatch,
  TableStyle,
  TableTotalsFunction,
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
export type {
  HistoryEntry,
  HistoryTransaction,
} from './commands/history.js';
export {
  canRedo,
  canUndo,
  History,
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
export type { ExportOptions, ImportCSVOptions, ImportResult } from './commands/import-export.js';
export { exportCSV, importCSV } from './commands/import-export.js';
export type {
  ApplyChangesOptions,
  InteractionControllerOptions,
} from './commands/interaction-controller.js';
export { InteractionController } from './commands/interaction-controller.js';
export type {
  BatchDeniedMode,
  BatchRejection,
  CellBatchCommand,
  CellBatchOperation,
  CellChangeInput,
  ChangeBatchResult,
  EditableCells,
  EditablePredicateContext,
  InteractionOperation,
  InteractionOrigin,
  InteractionPolicy,
  InteractionRestrictionContext,
  OperationEffect,
  OperationIntent,
  Origin,
  PermissionCode,
  PermissionDecision,
} from './commands/interaction-policy.js';
export {
  fixedFormPolicy,
  viewerPolicy,
} from './commands/interaction-policy.js';
export type {
  DefinedNameAttrs,
  DefinedNameDeleteResult,
  DefinedNameEntry,
  DefinedNameMutationResult,
} from './commands/named-ranges.js';
export { deleteDefinedName, listDefinedNames, upsertDefinedName } from './commands/named-ranges.js';
export type {
  HostPrinterDevice,
  HostPrinterPaperOption,
  MarginPreset,
  PageSetupEntry,
  PageSetupPatch,
  PrinterProfile,
} from './commands/page-setup.js';
export {
  listPageSetups,
  pageSetupForSheet,
  printerProfilesFromHostDevices,
  resetPageSetup,
  resolvePrinterProfileBounds,
  setPageSetup,
  setPrintTitleCols,
  setPrintTitleRows,
} from './commands/page-setup.js';
export type {
  CreatePivotTableOptions,
  CreatePivotTableResult,
  PivotFieldItemVisibility,
  PivotSourceField,
} from './commands/pivot-table.js';
export {
  createPivotTableFromRange,
  inferPivotFieldItems,
  inferPivotSourceFields,
} from './commands/pivot-table.js';
export type { BuildPrintDocumentOptions, PrintDocument } from './commands/print.js';
export { buildPrintDocument } from './commands/print.js';
export type { AllowedEditRangeOptions, SheetProtectionOptions } from './commands/protection.js';
export {
  addAllowedEditRange,
  isSheetProtected,
  protectedSheetPassword,
  setProtectedSheet,
  toggleProtectedSheet,
  whenSheetPasswordsHashed,
} from './commands/protection.js';
export type {
  QuickAnalysisAction,
  QuickAnalysisActionId,
  QuickAnalysisGroup,
  QuickAnalysisInput,
} from './commands/quick-analysis.js';
export {
  enabledQuickAnalysisActions,
  isQuickAnalysisActionEnabled,
  quickAnalysisActionById,
} from './commands/quick-analysis.js';
export type { ActiveSignature, FormulaScanOptions } from './commands/refs.js';
export { FUNCTION_SIGNATURES, findActiveSignature, suggestFunctions } from './commands/refs.js';
export { autofitColsWidth, autofitRowsHeight, setRowsHeight } from './commands/row-col-layout.js';
export type {
  ChartMutationAuthorization,
  CreateSessionChartOptions,
  SessionChartPatch,
  SessionChartSeriesInput,
} from './commands/session-chart.js';
export {
  addSessionChartSeries,
  clearSessionChart,
  clearSessionChartsInRange,
  createSessionChart,
  listSessionCharts,
  removeSessionChartSeries,
  sessionChartById,
  sessionChartsForRange,
  updateSessionChart,
} from './commands/session-chart.js';
export { addSheet } from './commands/sheet-mutate.js';
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
export { removeDuplicates } from './commands/sort.js';
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
  insertCols,
  insertRows,
} from './commands/structure.js';
export type { TableOption, TableRenameResult } from './commands/table-design.js';
export {
  renameActiveTable,
  setTableColumnFunction,
  toggleTableOption,
} from './commands/table-design.js';
export { applyTextScriptToRange } from './commands/text-script.js';
export { textToColumns } from './commands/text-to-columns.js';
export {
  addTraceArrow,
  clearTraceArrows,
  traceDependents,
  tracePrecedents,
} from './commands/traces.js';
export {
  clearWatchedCells,
  isWatched,
  setWatchWindowOpen,
  toggleWatchCell,
  unwatchCell,
  watchCell,
} from './commands/watch.js';
export type {
  WorkbookSaveFormat,
  WorkbookSaveLoss,
  WorkbookSaveOutcome,
  XlsbContentLossId,
} from './commands/workbook-save.js';
export { WORKBOOK_SAVE_FORMATS, writeWorkbook } from './commands/workbook-save.js';
export { formatA1Cell } from './engine/address.js';
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
export type {
  FileLossEntry,
  FileLossId,
  FileLossPhase,
  LoadDiagnostics,
  SaveDiagnostics,
} from './engine/file-diagnostics.js';
export { fileLossEntries, hasLoadLosses } from './engine/file-diagnostics.js';
export { LOCALE_TAGS } from './engine/function-metadata.js';
export type { LoadOptions } from './engine/loader.js';
export { isUsingStub } from './engine/loader.js';
export type {
  PivotTableSummary,
  WorkbookObjectKind,
  WorkbookObjectRecord,
} from './engine/passthrough-sync.js';
export {
  classifyWorkbookObjectPath,
  listWorkbookObjects,
  summarizePivotTables,
  WORKBOOK_OBJECT_KINDS,
  workbookObjectExtension,
  workbookObjectKindCounts,
  workbookObjectKindLabel,
  workbookObjectName,
  workbookObjectsByKind,
} from './engine/passthrough-sync.js';
export { hydrateTableOverlaysFromEngine, tableOverlaysFromEngine } from './engine/table-sync.js';
export type {
  Addr,
  CellValue,
  EngineCapabilities,
  FunctionMetadataResult,
  PhoneticProperties,
  PhoneticRun,
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
  PivotSortOrder,
  SheetVisibility,
} from './engine/types.js';
export { formatCell } from './engine/value.js';
export type {
  CellPatch,
  CellPatchAtomicResult,
  CellSnapshot,
  ChangeEvent,
  ChangeListener,
  WorkbookSaveOptions,
} from './engine/workbook-handle.js';
export { WorkbookHandle } from './engine/workbook-handle.js';
export type { WorkbookHandlePivotMethods } from './engine/workbook-handle-pivot.js';
export type { WorkbookThemePresetId } from './engine/workbook-theme-presets.js';
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
  ResolvedSpreadsheetPlatform,
  ResolvedSpreadsheetUiOptions,
  SpreadsheetFeatureSwitches,
  SpreadsheetPlatform,
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
  findReplace,
  formatDialog,
  formatPainter,
  full,
  goToSpecialDialog,
  hoverComment,
  hyperlinkDialog,
  isFeatureDefaultOff,
  iterativeDialog,
  minimal,
  namedRangeDialog,
  pageSetupDialog,
  pasteSpecial,
  pivotTableDialog,
  presets,
  quickAnalysis,
  resolveFlags,
  resolveSpreadsheetPlatform,
  resolveSpreadsheetUiOptions,
  slicer,
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
export type { DeepPartial, Locale, Strings } from './i18n/strings.js';
export { de, dictionaries, dictionaryLocaleFor, en, fr, ja, ko, th, zh } from './i18n/strings.js';
export type { ArgHelperLabels } from './interact/arg-helper.js';
export type { AutocompleteLabels } from './interact/autocomplete.js';
export type { BorderDrawHandle, BorderDrawMode } from './interact/border-draw.js';
export type { ClipboardHandle } from './interact/clipboard.js';
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
export type { ConditionalDialogOpenOptions } from './interact/conditional-dialog.js';
export type { ContextMenuDeps } from './interact/context-menu.js';
export { attachContextMenu } from './interact/context-menu.js';
export type {
  ContextMenuContext,
  ContextMenuItem,
  ContextMenuOptions,
} from './interact/context-menu-options.js';
export { openDataTableDialog } from './interact/data-table-dialog.js';
export type { FormatPainterHandle } from './interact/format-painter.js';
export { FUNCTION_DESCRIPTIONS } from './interact/function-catalog-text.js';
export type { FunctionCategory, FxDialogOpenOptions } from './interact/fx-dialog.js';
export type { ViewportOptions } from './interact/navigation-policy.js';
export type { OverlayOptions } from './interact/overlay-portal.js';
export type { RangePickerControlOptions } from './interact/range-picker-control.js';
export {
  attachRangePickerButton,
  updateRangePickerLabel,
} from './interact/range-picker-control.js';
export type { SessionChartLabels, SessionChartsHandle } from './interact/session-charts.js';
export type { SpreadsheetCompatibilityReportItem } from './interact/workbook-objects.js';
export {
  buildFileLossReport,
  buildSpreadsheetCompatibilityReport,
} from './interact/workbook-objects.js';
export type { DefaultDynamicDropdownsOptions } from './mount/dynamic-dropdowns-defaults.js';
export { createDefaultDynamicDropdownsCtx } from './mount/dynamic-dropdowns-defaults.js';
export type {
  FunctionArgumentHelp,
  FunctionArgumentHelpProvider,
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
export { openHostMenuFirstDropdown, Spreadsheet } from './mount.js';
export type {
  ErrorTriangleHit,
  ErrorTriangleKind,
  GridHitState,
  TotalsChevron,
  ValidationChevron,
} from './render/grid.js';
export { gridHitState } from './render/grid.js';
export type {
  CellAlign,
  CellBorderSide,
  CellBorderStyle,
  CellBorders,
  CellFormat,
  CellGradientFill,
  CellGradientStop,
  CellVAlign,
  CellValidation,
  ChartsSlice,
  ConditionalIconRef,
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
  SessionChartUnsupportedReason,
  SlicerNativeSettings,
  SlicerSpec,
  SlicersSlice,
  Sparkline,
  SparklineGroupSettings,
  SparklineKind,
  SparklineSlice,
  SpreadsheetStore,
  State,
  StatusAggKey,
  ThreadedCommentsSlice,
  TraceArrow,
  TracesSlice,
  ValidationErrorStyle,
  ValidationListSource,
  ValidationMeta,
  ValidationOp,
  WatchSlice,
  WorkbookViewMode,
} from './store/store.js';
export { createSpreadsheetStore, getPageSetup, mutators } from './store/store.js';
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
export type { SaveWorkbookWithPromptsOptions } from './toolbar/dialogs/workbook-save.js';
export {
  chooseWorkbookSaveFormat,
  saveWorkbookWithPrompts,
  workbookSaveFormats,
} from './toolbar/dialogs/workbook-save.js';
export type { DisabledReasonProjectionOptions } from './toolbar/menu-a11y.js';
export { projectDisabledReason, projectDisabledState } from './toolbar/menu-a11y.js';
export type { ToolbarMenuText } from './toolbar/menu-text.js';
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
  parseScriptCommand,
} from './toolbar/review-tools.js';
export type {
  RibbonActivationEntry,
  RibbonActivationKind,
  RibbonActivationSpec,
} from './toolbar/ribbon/activation.js';
export {
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
export type { RibbonFormatMutator as RibbonCommandTableMutator } from './toolbar/ribbon/command-tables.js';
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
export type { RibbonFillSeriesMode } from './toolbar/ribbon/fill-series.js';
export type { AutoSumFormulaName } from './toolbar/ribbon/menus/formulas.js';
export type { TableVariantId } from './toolbar/ribbon/menus/styles.js';
export type {
  RibbonMenuFactory,
  RibbonMenus,
  RibbonRenderHelpers,
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
export type { ActiveState } from './toolbar/ribbon-active-state.js';
export {
  BORDER_PRESETS,
  BORDER_STYLES,
  EMPTY_ACTIVE_STATE,
  localizeBorderPresets,
  localizeBorderStyles,
  projectActiveState,
} from './toolbar/ribbon-active-state.js';
export type {
  BuildRibbonModelOptions,
  RibbonCommand,
  RibbonGroupModel,
  RibbonOption,
  RibbonProfile,
  RibbonTab,
  RibbonTabModel,
  ToolbarLang,
  ToolbarText,
} from './toolbar/ribbon-model.js';
export {
  buildRibbonModel,
  CHART_CONTEXTUAL_RIBBON_TABS,
  CONTEXTUAL_RIBBON_TABS,
  EXCEL365_MAC_RIBBON_TABS,
  EXCEL365_STANDARD_RIBBON_TABS,
  FONT_FAMILIES,
  FONT_SIZES,
  isRibbonActivatableCommand,
  OPTIONAL_RIBBON_TABS,
  PICTURE_CONTEXTUAL_RIBBON_TABS,
  RIBBON_KEYSHORTCUTS,
  RIBBON_TAB_LABELS,
  RIBBON_TABS,
  ribbonActivatableCommandIds,
  ribbonActivatableCommands,
  ribbonActivatableSurfaceCommandIds,
  ribbonActivatableSurfaceCommands,
  ribbonSurfaceCommandIds,
  ribbonSurfaceCommands,
  ribbonTabLabel,
  SHAPE_CONTEXTUAL_RIBBON_TABS,
  TABLE_CONTEXTUAL_RIBBON_TABS,
} from './toolbar/ribbon-model.js';
