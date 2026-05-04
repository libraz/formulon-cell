export { Spreadsheet } from './mount.js';
export type { MountOptions, SpreadsheetInstance } from './mount.js';

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

// Custom function registry — `inst.formula.register(name, impl)`.
export type {
  CustomFunction,
  CustomFunctionMeta,
  CustomFunctionReturn,
} from './formula.js';
export { FormulaRegistry } from './formula.js';

// Cell renderer / editor registry — `inst.cells.registerFormatter(...)`.
export type {
  CellEditorEntry,
  CellEditorHandle,
  CellFormatterEntry,
  CellRenderInput,
} from './cells.js';
export { CellRegistry } from './cells.js';

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
  dedupeById,
  excel,
  flattenExtensions,
  minimal,
  presets,
  resolveFlags,
  sortByPriority,
  standard,
} from './extensions/index.js';
export { createI18nController } from './i18n/controller.js';
export type { I18nControllerInit } from './i18n/controller.js';
export { WorkbookHandle } from './engine/workbook-handle.js';
export type { ChangeEvent, ChangeListener } from './engine/workbook-handle.js';
export type { PassthroughSummary, TableSummary } from './engine/passthrough-sync.js';
export { summarizePassthroughs, summarizeTables } from './engine/passthrough-sync.js';
export type { LoadOptions } from './engine/loader.js';
export type { Addr, CellValue, EngineCapabilities, Range } from './engine/types.js';
export { formatCell, fromEngineValue } from './engine/value.js';
export type { ResolvedTheme } from './theme/resolve.js';
export { resolveTheme } from './theme/resolve.js';
export type {
  CellAlign,
  CellBorderSide,
  CellBorders,
  CellFormat,
  CellValidation,
  CellVAlign,
  ConditionalRule,
  ConditionalSlice,
  FormatSlice,
  MergesSlice,
  NegativeStyle,
  NumFmt,
  Sparkline,
  SparklineKind,
  SparklineSlice,
  SpreadsheetStore,
  State,
  StatusAggKey,
  ValidationErrorStyle,
  ValidationListSource,
  ValidationMeta,
  ValidationOp,
  WatchSlice,
} from './store/store.js';
export { createSpreadsheetStore, mutators } from './store/store.js';
export { clearComment, commentAt, setComment } from './commands/comment.js';
export {
  moveSheet,
  removeSheet,
  renameSheet,
  setSheetHidden,
} from './commands/sheet-mutate.js';
export {
  MAX_OUTLINE_LEVEL,
  OUTLINE_GUTTER_PER_LEVEL,
  collapseColGroup,
  collapseRowGroup,
  colGroupRangeAt,
  expandColGroup,
  expandRowGroup,
  groupCols,
  groupRows,
  isColGroupCollapsed,
  isRowGroupCollapsed,
  rowGroupRangeAt,
  ungroupCols,
  ungroupRows,
} from './commands/outline.js';
export type { CoercedInput } from './commands/coerce-input.js';
export {
  coerceInput,
  writeCoerced,
  writeInput,
  writeInputValidated,
} from './commands/coerce-input.js';
export type { ValidationOutcome } from './commands/validate.js';
export { resolveListValues, validateAgainst } from './commands/validate.js';
export type { RangeResolver } from './engine/range-resolver.js';
export {
  isRangeSource,
  makeRangeResolver,
  parseRangeRef,
  resolveNumericRange,
  resolveNumericRangeFromCells,
  resolveRangeRef,
} from './engine/range-resolver.js';
export type { SelectionStats } from './commands/aggregate.js';
export { aggregateSelection } from './commands/aggregate.js';
export { autoSum } from './commands/auto-sum.js';
export type { CopyResult } from './commands/clipboard/copy.js';
export { copy } from './commands/clipboard/copy.js';
export { cut } from './commands/clipboard/cut.js';
export type { PasteResult } from './commands/clipboard/paste.js';
export { pasteTSV } from './commands/clipboard/paste.js';
export { encodeTSV, parseTSV } from './commands/clipboard/tsv.js';
export type { CSVEncodeOptions } from './commands/clipboard/csv.js';
export { encodeCSV, parseCSV } from './commands/clipboard/csv.js';
export type { ExportOptions, ImportResult } from './commands/import-export.js';
export { exportCSV, importCSV } from './commands/import-export.js';
export { encodeHtml } from './commands/clipboard/html.js';
export {
  History,
  applyFormatSnapshot,
  applyLayoutSnapshot,
  applyMergesSnapshot,
  applySparklineSnapshot,
  canRedo,
  canUndo,
  captureFormatSnapshot,
  captureLayoutSnapshot,
  captureMergesSnapshot,
  captureSparklineSnapshot,
  recordFormatChange,
  recordLayoutChange,
  recordMergesChange,
  recordSparklineChange,
  redo,
  undo,
} from './commands/history.js';
export type { HistoryEntry, LayoutSnapshot, MergesSnapshot } from './commands/history.js';
export {
  bumpDecimals,
  bumpIndent,
  clearFormat,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  formatNumber,
  setAlign,
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
export type { FindMatch, FindOptions } from './commands/find.js';
export {
  applySubstitution,
  findAll,
  findNext,
  replaceAll,
  replaceOne,
} from './commands/find.js';
export type {
  AutocompleteDeps,
  AutocompleteHandle,
  AutocompleteTable,
} from './interact/autocomplete.js';
export { attachAutocomplete } from './interact/autocomplete.js';
export type { ArgHelperDeps, ArgHelperHandle } from './interact/arg-helper.js';
export { attachArgHelper } from './interact/arg-helper.js';
export type { ActiveSignature } from './commands/refs.js';
export { FUNCTION_SIGNATURES, findActiveSignature } from './commands/refs.js';
export type {
  ConditionalDialogDeps,
  ConditionalDialogHandle,
} from './interact/conditional-dialog.js';
export { attachConditionalDialog } from './interact/conditional-dialog.js';
export type { ContextMenuDeps } from './interact/context-menu.js';
export { attachContextMenu } from './interact/context-menu.js';
export type {
  NamedRangeDialogDeps,
  NamedRangeDialogHandle,
} from './interact/named-range-dialog.js';
export { attachNamedRangeDialog } from './interact/named-range-dialog.js';
export type { FilterDropdownDeps, FilterDropdownHandle } from './interact/filter-dropdown.js';
export { attachFilterDropdown } from './interact/filter-dropdown.js';
export type { FindReplaceDeps, FindReplaceHandle } from './interact/find-replace.js';
export { attachFindReplace } from './interact/find-replace.js';
export type { FormatDialogDeps, FormatDialogHandle } from './interact/format-dialog.js';
export { attachFormatDialog } from './interact/format-dialog.js';
export type { FormatPainterDeps, FormatPainterHandle } from './interact/format-painter.js';
export { attachFormatPainter } from './interact/format-painter.js';
export type { FxDialogDeps, FxDialogHandle } from './interact/fx-dialog.js';
export { FUNCTION_DESCRIPTIONS, attachFxDialog } from './interact/fx-dialog.js';
export type { HoverDeps, HoverHandle } from './interact/hover.js';
export { attachHover } from './interact/hover.js';
export type {
  IterativeDialogDeps,
  IterativeDialogHandle,
} from './interact/iterative-dialog.js';
export { attachIterativeDialog } from './interact/iterative-dialog.js';
export type { CellStyleDef, CellStyleId } from './commands/cell-styles.js';
export { CELL_STYLES, applyCellStyle, getCellStyle } from './commands/cell-styles.js';
export type {
  CellStylesGalleryDeps,
  CellStylesGalleryHandle,
} from './interact/cell-styles-gallery.js';
export { attachCellStylesGallery } from './interact/cell-styles-gallery.js';
export type { PasteSpecialDeps, PasteSpecialHandle } from './interact/paste-special.js';
export { attachPasteSpecial } from './interact/paste-special.js';
export type { StatusBarDeps, StatusBarHandle } from './interact/status-bar.js';
export { attachStatusBar } from './interact/status-bar.js';
export type { WatchPanelDeps, WatchPanelHandle } from './interact/watch-panel.js';
export { attachWatchPanel } from './interact/watch-panel.js';
export type { ValidationListDeps, ValidationListHandle } from './interact/validation.js';
export { attachValidationList } from './interact/validation.js';
export type {
  ClipboardCell,
  ClipboardSnapshot,
} from './commands/clipboard/snapshot.js';
export { captureSnapshot } from './commands/clipboard/snapshot.js';
export type {
  PasteOperation,
  PasteSpecialOptions,
  PasteSpecialResult,
  PasteWhat,
} from './commands/clipboard/paste-special.js';
export { pasteSpecial } from './commands/clipboard/paste-special.js';
export type { ClipboardHandle } from './interact/clipboard.js';
export { attachClipboard } from './interact/clipboard.js';
export type { FillOptions } from './commands/fill.js';
export { fillDestFor, fillRange } from './commands/fill.js';
export type { SortDirection, SortOptions } from './commands/sort.js';
export { removeDuplicates, sortRange } from './commands/sort.js';
export type { F4Result, FormulaRef } from './commands/refs.js';
export {
  REF_HIGHLIGHT_COLORS,
  extractRefs,
  rotateRefAt,
  shiftFormulaRefs,
  suggestFunctions,
} from './commands/refs.js';
export type { FilterPredicate } from './commands/filter.js';
export { applyFilter, clearFilter, distinctValues, setAutoFilter } from './commands/filter.js';
export {
  applyMerge,
  applyUnmerge,
  expandRangeWithMerges,
  mergeAnchorOf,
  mergeAt,
  stepWithMerge,
} from './commands/merge.js';
export { textToColumns } from './commands/text-to-columns.js';
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
export type { DeepPartial, Locale, Strings } from './i18n/strings.js';
export { defaultStrings, dictionaries, en, ja, mergeStrings } from './i18n/strings.js';
