// Shared toolbar action / dialog draft / constant declarations consumed by
// the React and Vue wrappers. Kept here so the per-button dispatch types only
// live in one place. Behaviour-bearing handlers go in `toolbar-actions.ts`;
// this file is type-and-constant only.

import type { CellStyleId } from '../commands/cell-styles.js';
import type { TableStyle } from '../commands/format-as-table.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import type { SessionChartKind } from '../store/types.js';
import type { RibbonReportItem, ScriptCommand } from '../toolbar/review-tools.js';

export type CellFormatAction =
  | 'dialog'
  | 'rowHeight'
  | 'autoFitRowHeight'
  | 'colWidth'
  | 'autoFitColWidth'
  | 'hideRows'
  | 'showRows'
  | 'hideCols'
  | 'showCols'
  | 'renameSheet'
  | 'hideSheet'
  | 'unhideSheet'
  | 'moveSheetLeft'
  | 'moveSheetRight'
  | 'tabColorNone'
  | 'tabColorRed'
  | 'tabColorOrange'
  | 'tabColorYellow'
  | 'tabColorGreen'
  | 'tabColorBlue'
  | 'tabColorPurple'
  | 'tabColorGray'
  | 'protectSheet';

export type FillAction =
  | 'down'
  | 'right'
  | 'up'
  | 'left'
  | 'flash'
  | 'series'
  | 'days'
  | 'weekdays'
  | 'months'
  | 'years';

export type ClearAction =
  | 'all'
  | 'formats'
  | 'contents'
  | 'comments'
  | 'hyperlinks'
  | 'conditional';

export type SortAction =
  | 'asc'
  | 'desc'
  | 'custom'
  | 'filter'
  | 'filter-clear'
  | 'filter-reapply'
  | 'filter-by-selected'
  | 'filter-advanced'
  | 'dedupe'
  | 'conditional'
  | 'named';

export type FilterDataAction = 'toggle' | 'clear' | 'reapply' | 'filter-by-selected' | 'advanced';

export type FindAction =
  | 'find'
  | 'replace'
  | 'go-to'
  | 'go-to-special'
  | 'formulas'
  | 'constants'
  | 'numbers'
  | 'text'
  | 'errors'
  | 'conditional-format'
  | 'data-validation'
  | 'comments';

export type CommentAction = 'delete-active' | 'delete-all';
export type ProtectionAction = 'allow-edit-range' | 'clear-allowed-edit-ranges';
export type HyperlinkAction = 'edit' | 'open' | 'clear' | 'external';
export type OutlineAxisAction = 'rows' | 'cols';

export type FunctionAction =
  | 'IF'
  | 'IFS'
  | 'AND'
  | 'OR'
  | 'XLOOKUP'
  | 'VLOOKUP'
  | 'INDEX'
  | 'MATCH'
  | 'CONCAT'
  | 'TEXT'
  | 'LEFT'
  | 'RIGHT'
  | 'TODAY'
  | 'NOW'
  | 'DATE'
  | 'YEAR'
  | 'PMT'
  | 'NPV'
  | 'IRR'
  | 'RATE'
  | 'ROUND'
  | 'SUMIF'
  | 'COUNTIF'
  | 'ABS';

export type TextOrientationAction =
  | 'angleCounterclockwise'
  | 'angleClockwise'
  | 'verticalText'
  | 'rotateTextUp'
  | 'rotateTextDown'
  | 'horizontalText'
  | 'formatAlignment';

export type TextToColumnsAction = 'comma' | 'tab' | 'semicolon' | 'space' | 'custom';

export type DataValidationAction =
  | 'settings'
  | 'circleInvalid'
  | 'clearCircles'
  | 'clearValidation';

export type FormulaAuditingAction =
  | 'errorChecking'
  | 'traceError'
  | 'ignoreError'
  | 'circleInvalid'
  | 'clearCircles';

export type ClearArrowsAction = 'clear-all' | 'clear-precedents' | 'clear-dependents';

export type PivotTableAction = 'dialog' | 'recommended' | 'new-sheet' | 'existing-sheet';
export type ChartAction = SessionChartKind | 'recommended';
export type PictureAction = 'device' | 'online';
export type ShapeAction = 'rectangle' | 'rounded-rectangle' | 'oval' | 'line' | 'arrow';
export type ScreenshotAction = 'current-view' | 'screen-clipping';
export type SymbolAction = string;

export type PrintAreaAction = 'set' | 'clear';
export type PrintTitleAction = 'rows' | 'cols' | 'clear';
export type PageBreakAction = 'insert-row' | 'insert-col' | 'remove-row' | 'remove-col' | 'reset';
export type SheetBackgroundAction = 'set' | 'clear';
export type ThemeAction = 'paper' | 'ink' | 'contrast';

export type CellStyleAction = CellStyleId;
export type FormatTableAction = TableStyle;

export type DefinedNameAction =
  | 'manager'
  | 'define'
  | 'createTopRow'
  | 'createBottomRow'
  | 'createLeftColumn'
  | 'createRightColumn'
  | `use:${string}`;

export type CalculationAction = 'auto' | 'autoNoTable' | 'manual' | 'iterative';
export type WatchAction = 'open' | 'add' | 'delete' | 'delete-all';
export type AddInAction = 'get' | 'my' | 'manage';
export type PdfAction = 'create' | 'share' | 'preferences';

export type { AutoSumAction } from './toolbar-actions.js';

export interface AdvancedFilterDialogDraft {
  listRange: string;
  criteriaRange: string;
  copyTo: string;
  uniqueOnly: boolean;
}

export interface RibbonReportDialogDraft {
  title: string;
  items: readonly RibbonReportItem[];
}

export interface SortDialogDraft {
  byCol: number;
  direction: 'asc' | 'desc';
  hasHeader: boolean;
}

export interface RemoveDuplicatesDialogDraft {
  columns: number[];
  hasHeader: boolean;
}

export interface DimensionDialogDraft {
  kind: 'rowHeight' | 'colWidth';
  value: string;
}

export interface SheetRenameDialogDraft {
  value: string;
}

export interface ScriptDialogDraft {
  command: ScriptCommand;
}

export interface AutomationRunDraft {
  command: ScriptCommand;
  range: string;
  changed: number;
}

export interface TextToColumnsDialogDraft {
  comma: boolean;
  tab: boolean;
  semicolon: boolean;
  space: boolean;
  collapseConsecutive: boolean;
}

export type SheetCellFor<I extends SpreadsheetInstance> =
  ReturnType<I['store']['getState']>['data']['cells'] extends Map<string, infer Cell>
    ? Cell
    : never;

export type SheetRangeFor<I extends SpreadsheetInstance> = ReturnType<
  I['store']['getState']
>['selection']['range'];

export type SheetCell = SheetCellFor<SpreadsheetInstance>;
export type SheetRange = SheetRangeFor<SpreadsheetInstance>;

export const MORE_SYMBOL_ACTION = '__more-symbols__';
export const CELL_STYLE_SECTION_ACTION_PREFIX = '__cell-style-section__';
export const TEXT_TO_COLUMNS_DIALOG_KEYS = ['comma', 'tab', 'semicolon', 'space'] as const;

export const SHEET_TAB_COLOR_ACTIONS = [
  { action: 'tabColorRed', color: '#c00000' },
  { action: 'tabColorOrange', color: '#ed7d31' },
  { action: 'tabColorYellow', color: '#ffc000' },
  { action: 'tabColorGreen', color: '#70ad47' },
  { action: 'tabColorBlue', color: '#4472c4' },
  { action: 'tabColorPurple', color: '#7030a0' },
  { action: 'tabColorGray', color: '#7f7f7f' },
] as const satisfies readonly { action: CellFormatAction; color: string }[];
