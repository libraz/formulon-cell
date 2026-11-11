// Public entry of @libraz/formulon-cell: mount(), the Workbook facade and the types their
// signatures use. Ribbon pieces live in ./toolbar, features in ./extensions.

export type { Chart } from './api/chart.js';
export type { Drawing, DrawingRef, Picture, Shape, Slicer } from './api/drawing.js';
export { type SpreadsheetFailure, spreadsheetFailureOf } from './api/errors.js';
export type { Disposer } from './api/events.js';
export type { SheetOoxml, WorkbookOoxml } from './api/ooxml.js';
export type { Pivot, PivotInput, PivotRef } from './api/pivot.js';
export type { PivotCache, PivotCacheRef, PivotWorksheetSourceInput } from './api/pivot-cache.js';
export type { ExportCsvOptions, ImportCsvOptions } from './api/sheet-cells.js';
export type { FindCellsOptions, ReplaceCellsOptions } from './api/sheet-find.js';
export type { ProtectOptions } from './api/sheet-format.js';
export type { CommandId } from './api/spreadsheet/command-ids.js';
export type { DialogId, DialogOptions, Dialogs } from './api/spreadsheet/dialogs.js';
export {
  type Audit,
  type CellFormatter,
  type CellRegistry,
  type CellTarget,
  type CommandIntent,
  type Commands,
  type Extensions,
  type FormulaHints,
  type History,
  type I18n,
  mount,
  type Selection,
  SpreadsheetInstance,
  type Viewport,
  type Watches,
} from './api/spreadsheet/instance.js';
export type {
  ApplyChangesOptions,
  CellChange,
  CellEvent,
  ChangeEvent,
  ErrorEvent,
  FormulaHint,
  HostProviders,
  RecalcEvent,
  RibbonToolbarConfig,
  SelectionState,
  SheetPrintDocument,
  SpreadsheetInstanceEvents,
  SpreadsheetOptions,
  SpreadsheetPatch,
  UpdateEvent,
  ViewportConfig,
  WorkbookEvent,
} from './api/spreadsheet/types.js';
export type { Table, TableRef } from './api/table.js';
export {
  type ChangedCell,
  type ChangeOrigin,
  type FileLoss,
  SAVE_FORMATS,
  type SaveFormat,
  type SaveOptions,
  type SaveResult,
  Sheet,
  type SheetCopyOptions,
  type SheetRef,
  Workbook,
  type WorkbookChange,
  type WorkbookEvents,
  type WorkbookOptions,
} from './api/workbook.js';
export type { CellFormatterEntry, CellRenderInput } from './cells.js';
export type { CategoryChartPoint, ChartCategory } from './charts/types.js';
export type {
  AutoCorrectEntry,
  AutoCorrectExceptions,
  AutoCorrectOptions,
  AutoCorrectState,
} from './commands/autocorrect.js';
export type {
  BatchCell,
  BatchRejection,
  ChangeBatchResult,
  EditableCells,
  EditablePredicateContext,
  InteractionOperation,
  InteractionOrigin,
  InteractionPolicy,
  InteractionRestrictionContext,
  PermissionCode,
  PermissionDecision,
} from './commands/interaction-policy.js';
export { fixedFormPolicy, viewerPolicy } from './commands/interaction-policy.js';
export type {
  HostPrinterDevice,
  HostPrinterPaperOption,
  PrinterProfile,
} from './commands/page-setup.js';
export { printerProfilesFromHostDevices } from './commands/page-setup.js';
export type { SpellingService } from './commands/spelling.js';
export { formatA1Cell } from './engine/address.js';
export type { WorkbookThemePresetId } from './engine/workbook-theme-presets.js';
export type { ExtensionSpreadsheet, ThemeName } from './extensions/types.js';
export type {
  ResolvedSpreadsheetPlatform,
  SpreadsheetFeatureSwitches,
  SpreadsheetPlatform,
  SpreadsheetUiOptions,
  SpreadsheetUiProfile,
} from './extensions/ui-options.js';
export { resolveSpreadsheetPlatform } from './extensions/ui-options.js';
export { BUILT_IN_LOCALES } from './i18n/locales.js';
export type { DeepPartial, Locale, Strings } from './i18n/strings.js';
export type { BorderDrawHandle, BorderDrawMode } from './interact/border-draw.js';
export type { ClipboardHandle } from './interact/clipboard.js';
export type {
  ContextMenuContext,
  ContextMenuItem,
  ContextMenuOptions,
} from './interact/context-menu-options.js';
export type { FormatPainterHandle } from './interact/format-painter.js';
export type { ViewportOptions } from './interact/navigation-policy.js';
export type { OverlayOptions } from './interact/overlay-portal.js';
export type {
  FunctionArgumentHelp,
  FunctionArgumentHelpProvider,
  ScreenClipCapture,
  ScreenClipCaptureResult,
  ScreenClipResult,
} from './mount.js';
export type { TraceArrow } from './store/store.js';
