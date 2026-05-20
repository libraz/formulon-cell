import type {
  CustomTableStyle,
  PivotTableStyleAssignment,
  TableOverlay,
} from '../commands/format-as-table.js';
import type { SheetView } from '../commands/sheet-views.js';
import type { Addr, CellValue, Range } from '../engine/types.js';

export type EditorMode =
  | { kind: 'idle' }
  | { kind: 'enter'; raw: string }
  | { kind: 'edit'; raw: string; caret: number };

export type NumFmt =
  | { kind: 'general' }
  | { kind: 'fixed'; decimals: number; thousands?: boolean; negativeStyle?: NegativeStyle }
  | {
      kind: 'currency';
      decimals: number;
      symbol?: string;
      negativeStyle?: NegativeStyle;
    }
  | { kind: 'percent'; decimals: number }
  | { kind: 'scientific'; decimals: number }
  | { kind: 'accounting'; decimals: number; symbol?: string }
  | { kind: 'date'; pattern: string }
  | { kind: 'time'; pattern: string }
  | { kind: 'datetime'; pattern: string }
  | { kind: 'special'; pattern: string }
  | { kind: 'text' }
  | { kind: 'custom'; pattern: string };

/** How negative numbers display. */
export type NegativeStyle = 'minus' | 'parens' | 'red' | 'red-parens';

export type CellAlign =
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';
export type CellVAlign = 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
export type TextDirection = 'context' | 'ltr' | 'rtl';
export type FillPattern =
  | 'gray125'
  | 'gray25'
  | 'gray50'
  | 'horizontal'
  | 'vertical'
  | 'diagonalDown'
  | 'diagonalUp';

/** Per-side border style. The renderer treats `false`/missing as "no border"
 *  and `true` as the legacy single-line border (back-compat). Object form
 *  carries a spreadsheet-style style + optional color. The full OOXML
 *  repertoire is supported; common spreadsheet border ordinals map to these
 *  names verbatim. */
export type CellBorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export type CellBorderSide =
  | boolean
  | {
      style: CellBorderStyle;
      color?: string;
    };

export interface CellBorders {
  top?: CellBorderSide;
  right?: CellBorderSide;
  bottom?: CellBorderSide;
  left?: CellBorderSide;
  /** Diagonal border directions: `\` runs top-left → bottom-right and `/`
   *  runs bottom-left → top-right. */
  diagonalDown?: CellBorderSide;
  diagonalUp?: CellBorderSide;
}

export interface CellFormat {
  /** Named cell style id last applied through the Cell Styles gallery. Direct
   *  formatting may still override individual fields, mirroring spreadsheets'
   *  style + local-format layering. */
  cellStyle?: string;
  numFmt?: NumFmt;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  align?: CellAlign;
  /** Vertical alignment. Default is 'bottom'. */
  vAlign?: CellVAlign;
  /** Wrap text within the cell — paint multi-line with hard wrapping. */
  wrap?: boolean;
  /** Shrink single-line text to fit the available cell width. */
  shrinkToFit?: boolean;
  /** Indent level (left-align padding in increments of ~8px). 0..15. */
  indent?: number;
  /** Text rotation in degrees, -90..90. 0 = horizontal. */
  rotation?: number;
  /** Text reading direction. Context maps to Excel's default/context mode. */
  textDirection?: TextDirection;
  borders?: CellBorders;
  /** Foreground (font) color as a CSS color string. */
  color?: string;
  /** Background fill color as a CSS color string. */
  fill?: string;
  /** Spreadsheet fill pattern drawn over `fill` using `fillPatternColor`. */
  fillPattern?: FillPattern;
  /** Foreground color for `fillPattern`. */
  fillPatternColor?: string;
  /** Override theme font family for this cell. */
  fontFamily?: string;
  /** Font size in CSS pixels. */
  fontSize?: number;
  /** Hyperlink URL. When set, the cell paints text underlined+blue and
   *  Ctrl/Cmd+click opens the link. */
  hyperlink?: string;
  /** Free-form note attached to the cell. Surfaced as a small triangle
   *  marker + hover tooltip; not exported to .xlsx for now. */
  comment?: string;
  /** Data validation. When kind === 'list', the cell paints a small ▼ on its
   *  right edge; clicking it opens a dropdown of `source` values. Other
   *  kinds (whole/decimal/date/time/textLength/custom) constrain typed input
   *  through `validateAgainst()` — the chevron is list-only. */
  validation?: CellValidation;
  /** Sheet-protection lock flag. Default is `true` (locked) — `undefined`
   *  is treated as locked. Set to `false` to opt the cell out of the
   *  per-sheet protection gate via `setCellLocked(range, false)`. The flag
   *  only takes effect when the containing sheet is also marked protected
   *  via `setSheetProtected`. */
  locked?: boolean;
  /** Formula-hidden flag. When true and the containing sheet is protected,
   *  formula text is suppressed from the formula bar, matching the desktop
   *  Format Cells > Protection > Hidden behavior. */
  formulaHidden?: boolean;
}

export interface CustomCellStyle {
  id: string;
  label: string;
  format: Partial<CellFormat>;
}

/** Comparison ordinals match OOXML data-validation `op`:
 *  0 between, 1 notBetween, 2 equal, 3 notEqual,
 *  4 lessThan, 5 lessThanOrEqual, 6 greaterThan, 7 greaterThanOrEqual. */
export type ValidationOp = 'between' | 'notBetween' | '=' | '<>' | '<' | '<=' | '>' | '>=';

/** OOXML errorStyle: 0 stop, 1 warning, 2 information. `stop` rejects the
 *  input outright; `warning` and `information` let the user keep the entry. */
export type ValidationErrorStyle = 'stop' | 'warning' | 'information';

/** Optional metadata that every validation kind carries. Mirrors the
 *  upstream `DataValidationEntry` shape minus `type` / `op` / formulas which
 *  the discriminated cases own. */
export interface ValidationMeta {
  /** Allow empty input regardless of constraint. Default true. */
  allowBlank?: boolean;
  errorStyle?: ValidationErrorStyle;
  errorTitle?: string;
  errorMessage?: string;
  promptTitle?: string;
  promptMessage?: string;
  /** Suppress the prompt tooltip even when the metadata is set. Default true. */
  showInputMessage?: boolean;
  /** Suppress the error dialog on invalid entry. Default true. */
  showErrorMessage?: boolean;
}

/** A list-source can be either an inline literal array of strings or a range
 *  reference (`Sheet1!$A$1:$A$10` or `$A$1:$A$10`). Range refs are resolved
 *  lazily by the dropdown / validator so the source-of-truth stays a single
 *  string in the OOXML formula1 slot. */
export type ValidationListSource = string[] | { ref: string };

/** Discriminated union — `kind` mirrors the OOXML `type` ordinal:
 *  list (3), whole (1), decimal (2), date (4), time (5), textLength (6),
 *  custom (7). */
export type CellValidation =
  | ({ kind: 'list'; source: ValidationListSource } & ValidationMeta)
  | ({ kind: 'whole'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'decimal'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'date'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'time'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'textLength'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'custom'; formula: string } & ValidationMeta);

export interface ViewportSlice {
  /** First row visible (zero-indexed). */
  rowStart: number;
  rowCount: number;
  colStart: number;
  colCount: number;
  zoom: number;
}

export interface SelectionSlice {
  active: Addr;
  range: Range;
  /** Anchor point for shift-click extension. */
  anchor: Addr;
  /** Disjoint ranges added via Ctrl/Cmd+click. The primary `range` plus these
   *  form a non-contiguous selection. Aggregations and the renderer iterate
   *  over `[range, ...extraRanges]`. Optional so legacy callers that build a
   *  selection literal don't need to opt into multi-range. */
  extraRanges?: Range[];
}

export interface LayoutSlice {
  /** Sheet 0 column widths in CSS pixels, indexed by col. */
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  defaultColWidth: number;
  defaultRowHeight: number;
  headerColWidth: number;
  headerRowHeight: number;
  /** Number of rows pinned at the top (the desktop-spreadsheet "Freeze Panes"). 0 = none. */
  freezeRows: number;
  /** Number of cols pinned at the left. 0 = none. */
  freezeCols: number;
  /** Rows hidden by the user. Geometry returns height 0; renderer skips them. */
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  /** Outline (group) level per row, 1..7. Absent or 0 means no group. The
   *  bracket gutter widens with the maximum level; collapse/expand toggle
   *  hides/shows the rows in a contiguous group. spreadsheet parity. */
  outlineRows: Map<number, number>;
  outlineCols: Map<number, number>;
  /** Width of the row outline gutter in CSS px — derived from
   *  `outlineRows` (max level × per-level slot). Maintained by outline
   *  mutators; renderer treats this as authoritative. */
  outlineRowGutter: number;
  outlineColGutter: number;
  /** Sheets whose tab is hidden (the desktop-spreadsheet "Hide Sheet"). Indexed by sheet
   *  index. Hidden sheets keep their data; only the tab is suppressed. */
  hiddenSheets: Set<number>;
  /** Sheet tab fill colors keyed by sheet index, matching Excel's Tab Color affordance. */
  sheetTabColors: Map<number, string>;
}

/** Snapshot of every populated cell on the active sheet. The store does not
 *  fan out engine reads to every render — it caches and invalidates on
 *  change events. */
export interface DataSlice {
  sheetIndex: number;
  cells: Map<string, { value: CellValue; formula: string | null }>;
}

/** Reference highlight surfaced by the editor while a formula is being
 *  authored. Mirrors `commands/refs.FormulaRef` shape; kept local to avoid
 *  a circular import between store and commands. */
export interface EditorRefHighlight {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  colorIndex: number;
}

export interface UiSlice {
  editor: EditorMode;
  hover: Addr | null;
  /** Theme id stamped on the host (`data-fc-theme`). Built-ins ship `paper`,
   *  `ink`, and `contrast`; consumers can register additional themes via
   *  custom CSS keyed off the same attribute. */
  theme: 'paper' | 'ink' | (string & {});
  /** Live preview range while dragging the fill handle. Painted as a dashed
   *  marquee; cleared when the drag ends. Null at rest. */
  fillPreview: Range | null;
  /** Source range currently held by the internal clipboard. Painted as a
   *  dashed copy marquee, similar to "marching ants". */
  copyRange: Range | null;
  /** Disjoint clipboard ranges for Ctrl/Cmd multi-row or multi-column copies. */
  copyRanges?: Range[] | null;
  /** When false, the renderer skips drawing inter-cell hairline gridlines. */
  showGridLines: boolean;
  /** When false, the renderer hides the row-number / column-letter strips. */
  showHeaders: boolean;
  /** When true, formula cells display the formula text instead of the
   *  evaluated value. Equivalent to the desktop-spreadsheet "Show Formulas" (Ctrl+`). */
  showFormulas: boolean;
  /** Workbook view mode surfaced by View > Workbook Views. The renderer keeps
   *  the grid model identical for now; chrome stamps the mode on the host so
   *  themes and wrappers can distinguish Normal, Page Layout, and Page Break Preview. */
  workbookView: WorkbookViewMode;
  /** Display refs in R1C1 form instead of A1 (headers, name box). Underlying
   *  storage stays A1 — only the rendered representation changes. */
  r1c1: boolean;
  /** Live formula-reference highlights (desktop spreadsheets: colored borders on referenced
   *  cells while editing a formula). Empty when no formula edit is active. */
  editorRefs: EditorRefHighlight[];
  /** Which aggregate stats appear in the status bar for the active selection. */
  statusAggs: StatusAggKey[];
  /** Excel-style right-click status bar toggles beyond aggregates. */
  statusOptions: StatusBarOptions;
  /** Range with autofilter enabled. Header row inside this range paints a
   *  small filter button (▼). null = no autofilter. */
  filterRange: Range | null;
  /** Value-filter criteria keyed by filter range + column. Reapply uses this
   *  to recompute hidden rows after the sheet data changes. */
  filterCriteria: ValueFilterCriteria[];
  /** Visibility flag for the Watch Window panel. Session-only state — the
   *  panel itself reads `watch.watches` for content. */
  watchPanelOpen: boolean;
  /** Excel-style sheet background image URLs keyed by sheet index. These are
   *  painted behind cells for on-screen use and intentionally excluded from print. */
  sheetBackgroundImages: Map<number, string>;
}

export interface ValueFilterCriteria {
  range: Range;
  byCol: number;
  hiddenValues: string[];
}

export type WorkbookViewMode = 'normal' | 'pageLayout' | 'pageBreakPreview';

/** Aggregate readouts available in the status bar. Spreadsheets ship these six. */
export type StatusAggKey = 'sum' | 'average' | 'count' | 'countNumbers' | 'min' | 'max';
export type StatusBarOptionKey =
  | 'capsLock'
  | 'numLock'
  | 'scrollLock'
  | 'uploadStatus'
  | 'macroRecording'
  | 'viewShortcuts'
  | 'zoom'
  | 'zoomSlider';

export interface StatusBarOptions {
  capsLock: boolean;
  numLock: boolean;
  scrollLock: boolean;
  uploadStatus: boolean;
  macroRecording: boolean;
  viewShortcuts: boolean;
  zoom: boolean;
  zoomSlider: boolean;
}

export interface FormatSlice {
  /** Per-cell format keyed by `addrKey`. Missing entries → defaults. */
  formats: Map<string, CellFormat>;
  /** Session-scoped custom named styles created from the Cell Styles gallery. */
  customCellStyles?: CustomCellStyle[];
}

export interface MergesSlice {
  /** Per-anchor (top-left of merge) → range. Anchor key is the addrKey of the
   *  top-left cell. Cells inside the merge but not the anchor are tracked via
   *  `byCell` for fast hit-test. */
  byAnchor: Map<string, Range>;
  /** Reverse index: any cell inside a merge → its anchor key. */
  byCell: Map<string, string>;
}

/** Icon-set artwork name. 3-slot families classify by [0.33, 0.67];
 *  5-slot families classify by [0.20, 0.40, 0.60, 0.80]. */
export type ConditionalIconSet =
  | 'arrows3'
  | 'arrows5'
  | 'triangles3'
  | 'traffic3'
  | 'trafficRim3'
  | 'symbols3'
  | 'flags3'
  | 'stars3'
  | 'quarters5'
  | 'ratings5'
  | 'bars5'
  | 'boxes5';

export type ConditionalScalePoint =
  | { kind: 'min' | 'max' }
  | { kind: 'number' | 'percent' | 'percentile'; value: number };

/** Conditional formatting rule. Evaluated by the renderer against cell
 *  values; the predicate kinds (cell-value, top-bottom, formula, blanks,
 *  duplicates, etc.) skip cells that don't satisfy their type-specific
 *  filter. */
export type ConditionalRule =
  | {
      kind: 'cell-value';
      range: Range;
      op: '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between';
      a: number;
      b?: number;
      /** Format applied when the predicate matches. Same shape as CellFormat
       *  but only `fill`, `color`, `bold`, `italic`, `underline`, `strike`
       *  are honored by the renderer. */
      apply: Partial<CellFormat>;
    }
  | {
      kind: 'color-scale';
      range: Range;
      /** Two- or three-stop gradient. Stops are CSS color strings; the values
       *  are interpolated linearly across [min, max] (or [min, mid, max]) of
       *  the range. */
      stops: [string, string] | [string, string, string];
      /** Excel-style threshold metadata for each color stop. Omitted rules use
       *  min/max for two-color scales and min/50th percentile/max for
       *  three-color scales. */
      thresholds?:
        | [ConditionalScalePoint, ConditionalScalePoint]
        | [ConditionalScalePoint, ConditionalScalePoint, ConditionalScalePoint];
    }
  | {
      kind: 'data-bar';
      range: Range;
      color: string;
      /** True for Excel's gradient-fill data bars, false for solid-fill bars.
       *  Omitted legacy rules render as solid bars. */
      gradient?: boolean;
      /** When true, paint the bar across the whole cell with the text on top
       *  (like the spreadsheet's "Show Bar Only" being false). */
      showValue?: boolean;
    }
  | {
      kind: 'icon-set';
      range: Range;
      /** Icon family. 3-slot or 5-slot determined by the suffix. */
      icons: ConditionalIconSet;
      /** When false, render only the icon and suppress the cell value text. */
      showValue?: boolean;
      /** Boundaries between icon slots, ordered from low to high. Omitted
       *  rules use Excel's default percent thresholds. */
      thresholds?: ConditionalScalePoint[];
      /** Invert slot index so the highest values get the "low" icon. */
      reverseOrder?: boolean;
    }
  | {
      kind: 'top-bottom';
      range: Range;
      /** `top` selects the N largest values, `bottom` the N smallest. */
      mode: 'top' | 'bottom';
      n: number;
      /** When true, `n` is interpreted as a percentage (0..100) of the
       *  range's numeric-cell count. */
      percent?: boolean;
      apply: Partial<CellFormat>;
    }
  | {
      kind: 'average';
      range: Range;
      /** Spreadsheet average rules over numeric cells in the range. */
      mode: 'above' | 'below' | 'equal-or-above' | 'equal-or-below';
      apply: Partial<CellFormat>;
    }
  | {
      kind: 'text-contains';
      range: Range;
      text: string;
      caseSensitive?: boolean;
      apply: Partial<CellFormat>;
    }
  | {
      kind: 'date-occurring';
      range: Range;
      period:
        | 'yesterday'
        | 'today'
        | 'tomorrow'
        | 'last7'
        | 'last-week'
        | 'this-week'
        | 'next-week'
        | 'last-month'
        | 'this-month'
        | 'next-month';
      apply: Partial<CellFormat>;
    }
  | {
      kind: 'formula';
      range: Range;
      /** Lightweight predicate. Supports comparator-prefix forms
       *  (`>10`, `<>"foo"`, `<= 0`, `=42`) and an `=`-prefixed cell formula
       *  evaluated through `wb.evaluateText` when the engine exposes one;
       *  otherwise the rule is a no-op. */
      formula: string;
      apply: Partial<CellFormat>;
    }
  | {
      kind: 'duplicates' | 'unique';
      range: Range;
      apply: Partial<CellFormat>;
    }
  | {
      kind: 'blanks' | 'non-blanks' | 'errors' | 'no-errors';
      range: Range;
      apply: Partial<CellFormat>;
    };

export interface ConditionalSlice {
  rules: ConditionalRule[];
}

/** Inline mini-chart attached to a single cell. The renderer paints `kind`
 *  inside the cell rect using the resolved numeric series at `source`. */
export type SparklineKind = 'line' | 'column' | 'win-loss';

export interface Sparkline {
  kind: SparklineKind;
  /** A1-style range, e.g. `B2:B12` or `Sheet2!B2:B12`. */
  source: string;
  /** Stroke for line, fill for column. CSS color. Default: `#0078d4`. */
  color?: string;
  /** When true, paint negatives in `negativeColor` (column / win-loss). */
  showNegative?: boolean;
  negativeColor?: string;
}

export interface SparklineSlice {
  /** Per-host-cell sparkline keyed by `addrKey`. */
  sparklines: Map<string, Sparkline>;
}

export type SessionChartKind = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'scatter';

/** Session chart overlay. This is intentionally UI-owned until the engine
 *  exposes chart authoring; the source range can later map to a persisted
 *  chart definition without changing the public command shape. */
export interface SessionChart {
  id: string;
  kind: SessionChartKind;
  source: Range;
  title?: string;
  color?: string;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

export interface ChartsSlice {
  charts: readonly SessionChart[];
}

export type SessionShapeKind = 'rectangle' | 'rounded-rectangle' | 'oval' | 'line' | 'arrow';

/** Session illustration overlay. Like session charts, this is UI-owned until
 *  writable drawing parts exist in the engine. */
export interface SessionIllustration {
  id: string;
  kind: 'shape' | 'image';
  shape?: SessionShapeKind;
  src?: string;
  alt?: string;
  sheet: number;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  color?: string;
}

export interface IllustrationsSlice {
  illustrations: readonly SessionIllustration[];
}

/** Cells the user has pinned in the Watch Window. Session-only — desktop spreadsheets
 *  parity: watches don't survive workbook close, and they aren't recorded
 *  in the undo stack. Order is insertion order. */
export interface WatchSlice {
  watches: readonly Addr[];
}

/** A single trace arrow drawn from a precedent or to a dependent.
 *  `kind: 'precedent'` arrows flow `from` (source cell) → `to` (active cell);
 *  `kind: 'dependent'` arrows flow `from` (active cell) → `to` (cell that
 *  reads from it). Painters distinguish the two visually. */
export interface TraceArrow {
  kind: 'precedent' | 'dependent';
  from: Addr;
  to: Addr;
}

/** Trace-precedents / trace-dependents arrows currently visible. Session-only;
 *  not recorded in the undo stack — spreadsheets keep trace arrows out of the
 *  history journal too. Each `tracePrecedents()` / `traceDependents()` call
 *  appends to `items`; `clearTraces()` empties the list. */
export interface TracesSlice {
  items: readonly TraceArrow[];
}

/** Per-cell triangle suppression for the error-indicator overlay. Key is
 *  `addrKey` (`sheet:row:col`). Session-only, NOT history-tracked — the spreadsheet's
 *  "Ignore Error" affordance only suppresses the marker for the current
 *  session and doesn't survive a reload. */
export interface ErrorIndicatorSlice {
  ignoredErrors: Set<string>;
  /** Excel-style "Circle Invalid Data" marks. Session-only; populated on demand
   *  from Data Validation > Circle Invalid Data and cleared by Clear
   *  Validation Circles. */
  validationCircles: Set<string>;
}

/** Page orientation for print / PDF export. */
export type PageOrientation = 'portrait' | 'landscape';

/** Paper size — covers the common ISO + ANSI sheets. The print document
 *  emits `@page { size: <paperSize> <orientation> }` which all major browsers
 *  honour for the print preview / PDF rendering. */
export type PaperSize = 'A4' | 'A3' | 'A5' | 'letter' | 'legal' | 'tabloid';

/** Margins in inches — spreadsheet parity. The dialog renders text inputs in inches;
 *  the print-CSS converts to `in` units verbatim. */
export interface PageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export type PrintCommentsMode = 'none' | 'asDisplayed' | 'endOfSheet';
export type PrintCellErrorsMode = 'displayed' | 'blank' | 'dash' | 'na';
export type PrintPageOrder = 'downThenOver' | 'overThenDown';
export type PrintQuality = 'automatic' | '300' | '600' | '1200';

/** Per-sheet page-setup configuration. Drives both the Page Setup dialog and
 *  the print document builder. Default values come from `defaultPageSetup()`;
 *  unset fields fall back to that default — `getPageSetup` always returns a
 *  fully-populated record. */
export interface PageSetup {
  orientation: PageOrientation;
  paperSize: PaperSize;
  margins: PageMargins;
  /** Minimum printable insets from the physical page edge, in inches. This is
   *  distinct from `printArea`: hosts may fill it from a printer profile or
   *  preview preset so content is laid out inside the device's non-printable
   *  border. Browser print cannot discover this automatically. */
  printableBounds?: PageMargins;
  /** Distance from page edge to header/footer text, in inches. */
  headerMargin?: number;
  footerMargin?: number;
  /** Center printed content within the page margins. */
  centerHorizontally?: boolean;
  centerVertically?: boolean;
  /** Header / footer text — desktop spreadsheets splits the strip into three slots
   *  (left / center / right). Empty / missing strings render as nothing. */
  headerLeft?: string;
  headerCenter?: string;
  headerRight?: string;
  footerLeft?: string;
  footerCenter?: string;
  footerRight?: string;
  /** Header/Footer tab options. */
  differentOddEvenPages?: boolean;
  differentFirstPage?: boolean;
  scaleHeaderFooterWithDocument?: boolean;
  alignHeaderFooterWithMargins?: boolean;
  /** A1-style print area, e.g. "A1:D20" or "A1:B2,D4:E5".
   *  Empty means print the used range. */
  printArea?: string;
  /** A1-style row range ("1:3" or "$1:$3") whose rows repeat at the top of
   *  every printed page. Single-row form ("2") is allowed. */
  printTitleRows?: string;
  /** A1-style column range ("A:B"). Repeats those columns on the left of
   *  every printed page. */
  printTitleCols?: string;
  /** Fit-to-N-pages-wide. 0 means no width constraint. */
  fitWidth?: number;
  /** Fit-to-N-pages-tall. 0 means no height constraint. */
  fitHeight?: number;
  /** Manual page breaks before the given zero-based rows / columns. */
  manualPageBreakRows?: number[];
  manualPageBreakCols?: number[];
  /** Print scale, 0.10..4.00 (1 = 100%). When `fitWidth`/`fitHeight` is set
   *  the browser ignores the explicit scale. */
  scale?: number;
  /** Printer quality and first printed page number from Excel's Page tab.
   *  `firstPageNumber` undefined means Auto. */
  printQuality?: PrintQuality;
  firstPageNumber?: number;
  /** Paint inter-cell hairline gridlines on the print document. */
  showGridlines?: boolean;
  /** Paint row-numbers and column-letters on the print document. */
  showHeadings?: boolean;
  /** Excel Sheet tab print options. Some are preserved for parity even when
   *  browser print has no exact equivalent. */
  blackAndWhite?: boolean;
  draftQuality?: boolean;
  comments?: PrintCommentsMode;
  cellErrorsAs?: PrintCellErrorsMode;
  pageOrder?: PrintPageOrder;
}

export interface PageSetupSlice {
  /** Per-sheet page-setup map, keyed by sheet index. Sheets without an
   *  entry fall back to `defaultPageSetup()`. History-tracked. */
  setupBySheet: Map<number, PageSetup>;
}

/** Default page-setup record. Returned by `getPageSetup` when the sheet has
 *  no explicit entry, and used as the baseline for partial-patch merges.
 *
 *  Margins match the "Normal" preset surfaced by the Page Setup dialog
 *  (`commands/page-setup.ts`) so the chrome can faithfully reflect the
 *  active preset in its dropdown. */
export function defaultPageSetup(): PageSetup {
  return {
    orientation: 'portrait',
    paperSize: 'A4',
    margins: { top: 0.75, right: 0.7, bottom: 0.75, left: 0.7 },
    headerMargin: 0.3,
    footerMargin: 0.3,
    centerHorizontally: false,
    centerVertically: false,
    differentOddEvenPages: false,
    differentFirstPage: false,
    scaleHeaderFooterWithDocument: true,
    alignHeaderFooterWithMargins: true,
    scale: 1,
    printQuality: 'automatic',
    showGridlines: false,
    showHeadings: false,
    blackAndWhite: false,
    draftQuality: false,
    comments: 'none',
    cellErrorsAs: 'displayed',
    pageOrder: 'downThenOver',
  };
}

/** A single spreadsheet-style slicer attached to one column of one spreadsheet Table.
 *  `selected` is the user's current chip selection — empty array means "all
 *  values pass" (no filter). The optional `x`/`y` coordinates anchor the
 *  floating panel relative to the host; absent = default offset. */
export interface SlicerSpec {
  /** Unique id within the workbook. Used as React-style key + state map key. */
  id: string;
  /** Engine-side `TableSummary.name`. */
  tableName: string;
  /** Column header text (matches one of `TableSummary.columns`). */
  column: string;
  /** Current chip selection. Empty array == include-all. */
  selected: readonly string[];
  /** Optional anchor x relative to the host. */
  x?: number;
  /** Optional anchor y relative to the host. */
  y?: number;
}

/** History-tracked slice carrying every active slicer. The collection is
 *  immutable — mutators rebuild the array. */
export interface SlicersSlice {
  slicers: readonly SlicerSpec[];
}

/** Session-level Format-as-Table overlays. Full ListObject authoring is
 *  engine-gated; this slice gives the UI spreadsheet-style table visuals today. */
export interface TablesSlice {
  tables: readonly TableOverlay[];
  customTableStyles?: readonly CustomTableStyle[];
  customPivotTableStyles?: readonly CustomTableStyle[];
  pivotTableStyles?: readonly PivotTableStyleAssignment[];
}

export interface SheetViewsSlice {
  views: readonly SheetView[];
  activeViewId: string | null;
}

export interface AllowedEditRange {
  id: string;
  title: string;
  range: Range;
  password?: string;
}

/** Workbook-level sheet-protection state. Each protected sheet is keyed by
 *  its index; the value records whether a password was supplied (currently
 *  stored verbatim, not enforced — v1 ships without password validation).
 *  NOT history-tracked: spreadsheets expose protection as a workbook-level
 *  setting and toggling it doesn't appear in undo. Cell-level locks live
 *  on `CellFormat.locked`; this slice only owns the sheet-side flag. */
export interface ProtectionSlice {
  protectedSheets: Map<number, { password?: string }>;
  workbookStructure?: { password?: string };
  allowedEditRanges: readonly AllowedEditRange[];
}

export interface State {
  viewport: ViewportSlice;
  selection: SelectionSlice;
  layout: LayoutSlice;
  data: DataSlice;
  ui: UiSlice;
  format: FormatSlice;
  merges: MergesSlice;
  conditional: ConditionalSlice;
  sparkline: SparklineSlice;
  charts: ChartsSlice;
  illustrations: IllustrationsSlice;
  watch: WatchSlice;
  traces: TracesSlice;
  errorIndicators: ErrorIndicatorSlice;
  pageSetup: PageSetupSlice;
  slicers: SlicersSlice;
  tables: TablesSlice;
  sheetViews: SheetViewsSlice;
  protection: ProtectionSlice;
}
