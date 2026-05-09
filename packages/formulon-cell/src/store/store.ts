import { createStore } from 'zustand/vanilla';
import type { TableOverlay } from '../commands/format-as-table.js';
import type { SheetView, SheetViewPatch } from '../commands/sheet-views.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { addrKey } from '../engine/workbook-handle.js';

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
  | { kind: 'text' }
  | { kind: 'custom'; pattern: string };

/** How negative numbers display. */
export type NegativeStyle = 'minus' | 'parens' | 'red' | 'red-parens';

export type CellAlign = 'left' | 'center' | 'right';
export type CellVAlign = 'top' | 'middle' | 'bottom';

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
  /** Indent level (left-align padding in increments of ~8px). 0..15. */
  indent?: number;
  /** Text rotation in degrees, -90..90. 0 = horizontal. */
  rotation?: number;
  borders?: CellBorders;
  /** Foreground (font) color as a CSS color string. */
  color?: string;
  /** Background fill color as a CSS color string. */
  fill?: string;
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
  /** Number of rows pinned at the top (Excel "Freeze Panes"). 0 = none. */
  freezeRows: number;
  /** Number of cols pinned at the left. 0 = none. */
  freezeCols: number;
  /** Rows hidden by the user. Geometry returns height 0; renderer skips them. */
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  /** Outline (group) level per row, 1..7. Absent or 0 means no group. The
   *  bracket gutter widens with the maximum level; collapse/expand toggle
   *  hides/shows the rows in a contiguous group. Excel parity. */
  outlineRows: Map<number, number>;
  outlineCols: Map<number, number>;
  /** Width of the row outline gutter in CSS px — derived from
   *  `outlineRows` (max level × per-level slot). Maintained by outline
   *  mutators; renderer treats this as authoritative. */
  outlineRowGutter: number;
  outlineColGutter: number;
  /** Sheets whose tab is hidden (Excel "Hide Sheet"). Indexed by sheet
   *  index. Hidden sheets keep their data; only the tab is suppressed. */
  hiddenSheets: Set<number>;
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
   *  dashed copy marquee, similar to Excel's marching ants. */
  copyRange: Range | null;
  /** When false, the renderer skips drawing inter-cell hairline gridlines. */
  showGridLines: boolean;
  /** When false, the renderer hides the row-number / column-letter strips. */
  showHeaders: boolean;
  /** When true, formula cells display the formula text instead of the
   *  evaluated value. Equivalent to Excel "Show Formulas" (Ctrl+`). */
  showFormulas: boolean;
  /** Display refs in R1C1 form instead of A1 (headers, name box). Underlying
   *  storage stays A1 — only the rendered representation changes. */
  r1c1: boolean;
  /** Live formula-reference highlights (Excel: colored borders on referenced
   *  cells while editing a formula). Empty when no formula edit is active. */
  editorRefs: EditorRefHighlight[];
  /** Which aggregate stats appear in the status bar for the active selection. */
  statusAggs: StatusAggKey[];
  /** Range with autofilter enabled. Header row inside this range paints a
   *  small filter button (▼). null = no autofilter. */
  filterRange: Range | null;
  /** Visibility flag for the Watch Window panel. Session-only state — the
   *  panel itself reads `watch.watches` for content. */
  watchPanelOpen: boolean;
}

/** Aggregate readouts available in the status bar. Excel ships these six. */
export type StatusAggKey = 'sum' | 'average' | 'count' | 'countNumbers' | 'min' | 'max';

export interface FormatSlice {
  /** Per-cell format keyed by `addrKey`. Missing entries → defaults. */
  formats: Map<string, CellFormat>;
}

export interface MergesSlice {
  /** Per-anchor (top-left of merge) → range. Anchor key is the addrKey of the
   *  top-left cell. Cells inside the merge but not the anchor are tracked via
   *  `byCell` for fast hit-test. */
  byAnchor: Map<string, Range>;
  /** Reverse index: any cell inside a merge → its anchor key. */
  byCell: Map<string, string>;
}

/** Icon-set artwork name. `arrows3` / `traffic3` / `stars3` use 3 slots
 *  classified by [0.33, 0.67]; `arrows5` uses 5 slots classified by
 *  [0.20, 0.40, 0.60, 0.80]. */
export type ConditionalIconSet = 'arrows3' | 'arrows5' | 'traffic3' | 'stars3';

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
    }
  | {
      kind: 'data-bar';
      range: Range;
      color: string;
      /** When true, paint the bar across the whole cell with the text on top
       *  (like Excel's "Show Bar Only" being false). */
      showValue?: boolean;
    }
  | {
      kind: 'icon-set';
      range: Range;
      /** Icon family. 3-slot or 5-slot determined by the suffix. */
      icons: ConditionalIconSet;
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

export type SessionChartKind = 'column' | 'line';

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

/** Cells the user has pinned in the Watch Window. Session-only — Excel
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
 *  not recorded in the undo stack — Excel keeps trace arrows out of the
 *  history journal too. Each `tracePrecedents()` / `traceDependents()` call
 *  appends to `items`; `clearTraces()` empties the list. */
export interface TracesSlice {
  items: readonly TraceArrow[];
}

/** Per-cell triangle suppression for the error-indicator overlay. Key is
 *  `addrKey` (`sheet:row:col`). Session-only, NOT history-tracked — Excel's
 *  "Ignore Error" affordance only suppresses the marker for the current
 *  session and doesn't survive a reload. */
export interface ErrorIndicatorSlice {
  ignoredErrors: Set<string>;
}

/** Page orientation for print / PDF export. */
export type PageOrientation = 'portrait' | 'landscape';

/** Paper size — covers the common ISO + ANSI sheets. The print document
 *  emits `@page { size: <paperSize> <orientation> }` which all major browsers
 *  honour for the print preview / PDF rendering. */
export type PaperSize = 'A4' | 'A3' | 'A5' | 'letter' | 'legal' | 'tabloid';

/** Margins in inches — Excel parity. The dialog renders text inputs in inches;
 *  the print-CSS converts to `in` units verbatim. */
export interface PageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

/** Per-sheet page-setup configuration. Drives both the Page Setup dialog and
 *  the print document builder. Default values come from `defaultPageSetup()`;
 *  unset fields fall back to that default — `getPageSetup` always returns a
 *  fully-populated record. */
export interface PageSetup {
  orientation: PageOrientation;
  paperSize: PaperSize;
  margins: PageMargins;
  /** Header / footer text — Excel splits the strip into three slots
   *  (left / center / right). Empty / missing strings render as nothing. */
  headerLeft?: string;
  headerCenter?: string;
  headerRight?: string;
  footerLeft?: string;
  footerCenter?: string;
  footerRight?: string;
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
  /** Print scale, 0.10..4.00 (1 = 100%). When `fitWidth`/`fitHeight` is set
   *  the browser ignores the explicit scale. */
  scale?: number;
  /** Paint inter-cell hairline gridlines on the print document. */
  showGridlines?: boolean;
  /** Paint row-numbers and column-letters on the print document. */
  showHeadings?: boolean;
}

export interface PageSetupSlice {
  /** Per-sheet page-setup map, keyed by sheet index. Sheets without an
   *  entry fall back to `defaultPageSetup()`. History-tracked. */
  setupBySheet: Map<number, PageSetup>;
}

/** Default page-setup record. Returned by `getPageSetup` when the sheet has
 *  no explicit entry, and used as the baseline for partial-patch merges. */
export function defaultPageSetup(): PageSetup {
  return {
    orientation: 'portrait',
    paperSize: 'A4',
    margins: { top: 0.7, right: 0.7, bottom: 0.7, left: 0.7 },
    scale: 1,
    showGridlines: false,
    showHeadings: false,
  };
}

/** A single Excel-style slicer attached to one column of one Excel Table.
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
 *  engine-gated; this slice gives the UI Excel-style table visuals today. */
export interface TablesSlice {
  tables: readonly TableOverlay[];
}

export interface SheetViewsSlice {
  views: readonly SheetView[];
  activeViewId: string | null;
}

/** Workbook-level sheet-protection state. Each protected sheet is keyed by
 *  its index; the value records whether a password was supplied (currently
 *  stored verbatim, not enforced — v1 ships without password validation).
 *  NOT history-tracked: Excel exposes protection as a workbook-level
 *  setting and toggling it doesn't appear in undo. Cell-level locks live
 *  on `CellFormat.locked`; this slice only owns the sheet-side flag. */
export interface ProtectionSlice {
  protectedSheets: Map<number, { password?: string }>;
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
  watch: WatchSlice;
  traces: TracesSlice;
  errorIndicators: ErrorIndicatorSlice;
  pageSetup: PageSetupSlice;
  slicers: SlicersSlice;
  tables: TablesSlice;
  sheetViews: SheetViewsSlice;
  protection: ProtectionSlice;
}

const initialAddr = (sheet = 0): Addr => ({ sheet, row: 0, col: 0 });
const initialRange = (sheet = 0): Range => ({ sheet, r0: 0, c0: 0, r1: 0, c1: 0 });

function rangesIntersect(a: Range, b: Range): boolean {
  return a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);
}

function keyInRange(key: string, range: Range): boolean {
  const [sheet, row, col] = key.split(':').map(Number);
  return (
    sheet === range.sheet &&
    row !== undefined &&
    col !== undefined &&
    row >= range.r0 &&
    row <= range.r1 &&
    col >= range.c0 &&
    col <= range.c1
  );
}

export const createSpreadsheetStore = () =>
  createStore<State>(() => ({
    viewport: { rowStart: 0, rowCount: 40, colStart: 0, colCount: 16, zoom: 1 },
    selection: {
      active: initialAddr(),
      range: initialRange(),
      anchor: initialAddr(),
      extraRanges: [],
    },
    layout: {
      colWidths: new Map(),
      rowHeights: new Map(),
      defaultColWidth: 64,
      defaultRowHeight: 20,
      headerColWidth: 46,
      headerRowHeight: 22,
      freezeRows: 0,
      freezeCols: 0,
      hiddenRows: new Set(),
      hiddenCols: new Set(),
      outlineRows: new Map(),
      outlineCols: new Map(),
      outlineRowGutter: 0,
      outlineColGutter: 0,
      hiddenSheets: new Set(),
    },
    data: { sheetIndex: 0, cells: new Map() },
    ui: {
      editor: { kind: 'idle' },
      hover: null,
      theme: 'paper',
      fillPreview: null,
      copyRange: null,
      showGridLines: true,
      showHeaders: true,
      showFormulas: false,
      editorRefs: [],
      r1c1: false,
      statusAggs: ['sum', 'average', 'count'],
      filterRange: null,
      watchPanelOpen: false,
    },
    format: { formats: new Map() },
    merges: { byAnchor: new Map(), byCell: new Map() },
    conditional: { rules: [] },
    sparkline: { sparklines: new Map() },
    charts: { charts: [] },
    watch: { watches: [] },
    traces: { items: [] },
    errorIndicators: { ignoredErrors: new Set() },
    pageSetup: { setupBySheet: new Map() },
    slicers: { slicers: [] },
    tables: { tables: [] },
    sheetViews: { views: [], activeViewId: null },
    protection: { protectedSheets: new Map() },
  }));

export type SpreadsheetStore = ReturnType<typeof createSpreadsheetStore>;

// Tiny mutation helpers — single source of truth for state shape changes.
export const mutators = {
  setActive(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => ({
      ...s,
      selection: {
        active: addr,
        anchor: addr,
        range: { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col },
        extraRanges: [],
      },
    }));
  },

  /** Append a single-cell range to the current multi-selection. The cell
   *  becomes the new active/anchor so a follow-up shift-click extends from it.
   *  No-op if `addr` is the same sheet/row/col as the current active cell. */
  addExtraCell(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const sameAsActive =
        s.selection.active.sheet === addr.sheet &&
        s.selection.active.row === addr.row &&
        s.selection.active.col === addr.col;
      if (sameAsActive) return s;
      const prevPrimary = s.selection.range;
      // Demote the current primary range into extraRanges, promote the new
      // cell to primary so future shift-extends widen the new band.
      return {
        ...s,
        selection: {
          active: addr,
          anchor: addr,
          range: {
            sheet: addr.sheet,
            r0: addr.row,
            c0: addr.col,
            r1: addr.row,
            c1: addr.col,
          },
          extraRanges: [...(s.selection.extraRanges ?? []), prevPrimary],
        },
      };
    });
  },

  extendRangeTo(store: SpreadsheetStore, to: Addr): void {
    store.setState((s) => {
      const a = s.selection.anchor;
      return {
        ...s,
        selection: {
          ...s.selection,
          active: to,
          range: {
            sheet: to.sheet,
            r0: Math.min(a.row, to.row),
            c0: Math.min(a.col, to.col),
            r1: Math.max(a.row, to.row),
            c1: Math.max(a.col, to.col),
          },
        },
      };
    });
  },

  setEditor(store: SpreadsheetStore, mode: EditorMode): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, editor: mode } }));
  },

  setHover(store: SpreadsheetStore, addr: Addr | null): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, hover: addr } }));
  },

  setTheme(store: SpreadsheetStore, theme: 'paper' | 'ink' | (string & {})): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, theme } }));
  },

  setShowGridLines(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, showGridLines: on } }));
  },

  setShowHeaders(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, showHeaders: on } }));
  },

  setShowFormulas(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, showFormulas: on } }));
  },

  setR1C1(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, r1c1: on } }));
  },

  toggleStatusAgg(store: SpreadsheetStore, key: StatusAggKey): void {
    store.setState((s) => {
      const set = new Set(s.ui.statusAggs);
      if (set.has(key)) set.delete(key);
      else set.add(key);
      return { ...s, ui: { ...s.ui, statusAggs: Array.from(set) } };
    });
  },

  setStatusAggs(store: SpreadsheetStore, keys: StatusAggKey[]): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, statusAggs: [...keys] } }));
  },

  setFilterRange(store: SpreadsheetStore, range: Range | null): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, filterRange: range } }));
  },

  setEditorRefs(store: SpreadsheetStore, refs: EditorRefHighlight[]): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, editorRefs: refs } }));
  },

  setZoom(store: SpreadsheetStore, zoom: number): void {
    const z = Math.max(0.5, Math.min(4, zoom));
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, zoom: z } }));
  },

  setViewportSize(store: SpreadsheetStore, rowCount: number, colCount: number): void {
    const rows = Math.max(1, Math.floor(rowCount));
    const cols = Math.max(1, Math.floor(colCount));
    const MAX_ROW = 1_048_575;
    const MAX_COL = 16_383;
    store.setState((s) => {
      if (s.viewport.rowCount === rows && s.viewport.colCount === cols) return s;
      const maxRowStart = Math.max(s.layout.freezeRows, MAX_ROW + 1 - rows);
      const maxColStart = Math.max(s.layout.freezeCols, MAX_COL + 1 - cols);
      return {
        ...s,
        viewport: {
          ...s.viewport,
          rowCount: rows,
          colCount: cols,
          rowStart: Math.min(maxRowStart, Math.max(s.layout.freezeRows, s.viewport.rowStart)),
          colStart: Math.min(maxColStart, Math.max(s.layout.freezeCols, s.viewport.colStart)),
        },
      };
    });
  },

  setFillPreview(store: SpreadsheetStore, range: Range | null): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, fillPreview: range } }));
  },

  setCopyRange(store: SpreadsheetStore, range: Range | null): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, copyRange: range ? { ...range } : null } }));
  },

  setCell(
    store: SpreadsheetStore,
    addr: Addr,
    value: CellValue,
    formula: string | null = null,
  ): void {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      if (value.kind === 'blank' && !formula) cells.delete(addrKey(addr));
      else cells.set(addrKey(addr), { value, formula });
      return { ...s, data: { ...s.data, cells } };
    });
  },

  replaceCells(
    store: SpreadsheetStore,
    entries: Iterable<{ addr: Addr; value: CellValue; formula: string | null }>,
  ): void {
    const cells = new Map<string, { value: CellValue; formula: string | null }>();
    for (const e of entries) cells.set(addrKey(e.addr), { value: e.value, formula: e.formula });
    store.setState((s) => ({ ...s, data: { ...s.data, cells } }));
  },

  /** Switch the active sheet index. Cells must be re-hydrated separately
   *  via `replaceCells` after calling this. Resets selection to A1 on the
   *  new sheet. */
  setSheetIndex(store: SpreadsheetStore, idx: number): void {
    store.setState((s) => ({
      ...s,
      data: { ...s.data, sheetIndex: idx, cells: new Map() },
      selection: {
        active: { sheet: idx, row: 0, col: 0 },
        anchor: { sheet: idx, row: 0, col: 0 },
        range: { sheet: idx, r0: 0, c0: 0, r1: 0, c1: 0 },
      },
    }));
  },

  setColWidth(store: SpreadsheetStore, col: number, px: number): void {
    store.setState((s) => {
      const colWidths = new Map(s.layout.colWidths);
      colWidths.set(col, Math.max(28, Math.min(800, px)));
      return { ...s, layout: { ...s.layout, colWidths } };
    });
  },

  setRowHeight(store: SpreadsheetStore, row: number, px: number): void {
    store.setState((s) => {
      const rowHeights = new Map(s.layout.rowHeights);
      rowHeights.set(row, Math.max(16, Math.min(400, px)));
      return { ...s, layout: { ...s.layout, rowHeights } };
    });
  },

  /** Set entire row/col selection without an active-cell address change.
   *  Used when the user clicks a row/col header. */
  selectRow(store: SpreadsheetStore, row: number): void {
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: s.data.sheetIndex, row, col: 0 },
        anchor: { sheet: s.data.sheetIndex, row, col: 0 },
        range: { sheet: s.data.sheetIndex, r0: row, c0: 0, r1: row, c1: 16383 },
        extraRanges: [],
      },
    }));
  },

  selectCol(store: SpreadsheetStore, col: number): void {
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: s.data.sheetIndex, row: 0, col },
        anchor: { sheet: s.data.sheetIndex, row: 0, col },
        range: { sheet: s.data.sheetIndex, r0: 0, c0: col, r1: 1048575, c1: col },
        extraRanges: [],
      },
    }));
  },

  selectAll(store: SpreadsheetStore): void {
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: s.data.sheetIndex, row: 0, col: 0 },
        anchor: { sheet: s.data.sheetIndex, row: 0, col: 0 },
        range: { sheet: s.data.sheetIndex, r0: 0, c0: 0, r1: 1048575, c1: 16383 },
        extraRanges: [],
      },
    }));
  },

  /** Replace the primary selection range without touching active/anchor. Used
   *  by merge-aware navigation to grow a shift-extend so it covers an entire
   *  merge rectangle. */
  setRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => ({
      ...s,
      selection: { ...s.selection, range: { ...range } },
    }));
  },

  /** Merge a partial format into the cell at `addr`. Pass `null` to clear. */
  setCellFormat(store: SpreadsheetStore, addr: Addr, patch: Partial<CellFormat> | null): void {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      const key = addrKey(addr);
      if (patch === null) {
        formats.delete(key);
      } else {
        const prev = formats.get(key) ?? {};
        const next: CellFormat = { ...prev, ...patch };
        if (patch.borders) next.borders = { ...(prev.borders ?? {}), ...patch.borders };
        formats.set(key, next);
      }
      return { ...s, format: { formats } };
    });
  },

  /** Apply `patch` to every cell in `range`. Pass `null` to clear.
   *  Skips no-op when the range is huge (full-row/column or selectAll) — until
   *  we have row/column-level format storage, painting per-cell entries for
   *  millions of empty cells would OOM. */
  setRangeFormat(store: SpreadsheetStore, range: Range, patch: Partial<CellFormat> | null): void {
    const area = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
    if (area > 100_000) return;
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      const sheet = range.sheet;
      for (let r = range.r0; r <= range.r1; r += 1) {
        for (let c = range.c0; c <= range.c1; c += 1) {
          const key = addrKey({ sheet, row: r, col: c });
          if (patch === null) {
            formats.delete(key);
          } else {
            const prev = formats.get(key) ?? {};
            const next: CellFormat = { ...prev, ...patch };
            if (patch.borders) next.borders = { ...(prev.borders ?? {}), ...patch.borders };
            formats.set(key, next);
          }
        }
      }
      return { ...s, format: { formats } };
    });
  },

  scrollBy(store: SpreadsheetStore, dRow: number, dCol: number): void {
    // Excel sheet bounds — keep at least one body row/col visible past the
    // freeze zone, otherwise the viewport disappears off the right/bottom.
    const MAX_ROW = 1_048_575;
    const MAX_COL = 16_383;
    store.setState((s) => {
      const maxRowStart = Math.max(s.layout.freezeRows, MAX_ROW + 1 - s.viewport.rowCount);
      const maxColStart = Math.max(s.layout.freezeCols, MAX_COL + 1 - s.viewport.colCount);
      return {
        ...s,
        viewport: {
          ...s.viewport,
          rowStart: Math.min(
            maxRowStart,
            Math.max(s.layout.freezeRows, s.viewport.rowStart + dRow),
          ),
          colStart: Math.min(
            maxColStart,
            Math.max(s.layout.freezeCols, s.viewport.colStart + dCol),
          ),
        },
      };
    });
  },

  /** Pin the first `rows` rows / `cols` columns. Pass 0/0 to unfreeze.
   *  Scrolls past the frozen zone if the body viewport is currently inside it. */
  setFreezePanes(store: SpreadsheetStore, rows: number, cols: number): void {
    const fr = Math.max(0, Math.floor(rows));
    const fc = Math.max(0, Math.floor(cols));
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, freezeRows: fr, freezeCols: fc },
      viewport: {
        ...s.viewport,
        rowStart: Math.max(fr, s.viewport.rowStart),
        colStart: Math.max(fc, s.viewport.colStart),
      },
    }));
  },

  /** Merge a range into a single cell. The top-left becomes the anchor;
   *  every other cell in the range is mapped back to that anchor. If the
   *  range is 1×1 it's a no-op. */
  mergeRange(store: SpreadsheetStore, range: Range): void {
    if (range.r0 === range.r1 && range.c0 === range.c1) return;
    store.setState((s) => {
      const byAnchor = new Map(s.merges.byAnchor);
      const byCell = new Map(s.merges.byCell);
      // Strip any existing merges that touch the range.
      for (const [anchorKey, r] of byAnchor) {
        if (
          r.sheet === range.sheet &&
          !(r.r1 < range.r0 || r.r0 > range.r1 || r.c1 < range.c0 || r.c0 > range.c1)
        ) {
          byAnchor.delete(anchorKey);
          for (let row = r.r0; row <= r.r1; row += 1) {
            for (let col = r.c0; col <= r.c1; col += 1) {
              byCell.delete(addrKey({ sheet: r.sheet, row, col }));
            }
          }
        }
      }
      const anchor = { sheet: range.sheet, row: range.r0, col: range.c0 };
      const ak = addrKey(anchor);
      byAnchor.set(ak, range);
      for (let row = range.r0; row <= range.r1; row += 1) {
        for (let col = range.c0; col <= range.c1; col += 1) {
          if (row === range.r0 && col === range.c0) continue;
          byCell.set(addrKey({ sheet: range.sheet, row, col }), ak);
        }
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  },

  addConditionalRule(store: SpreadsheetStore, rule: ConditionalRule): void {
    store.setState((s) => ({
      ...s,
      conditional: { rules: [...s.conditional.rules, rule] },
    }));
  },

  removeConditionalRuleAt(store: SpreadsheetStore, idx: number): void {
    store.setState((s) => ({
      ...s,
      conditional: { rules: s.conditional.rules.filter((_, i) => i !== idx) },
    }));
  },

  clearConditionalRules(store: SpreadsheetStore): void {
    store.setState((s) => ({ ...s, conditional: { rules: [] } }));
  },

  clearConditionalRulesInRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const rules = s.conditional.rules.filter((rule) => !rangesIntersect(rule.range, range));
      if (rules.length === s.conditional.rules.length) return s;
      return { ...s, conditional: { rules } };
    });
  },

  /** Attach a sparkline spec to `addr`. Pass `null` to remove. */
  setSparkline(store: SpreadsheetStore, addr: Addr, spec: Sparkline | null): void {
    store.setState((s) => {
      const sparklines = new Map(s.sparkline.sparklines);
      const key = addrKey(addr);
      if (spec === null) sparklines.delete(key);
      else sparklines.set(key, { ...spec });
      return { ...s, sparkline: { sparklines } };
    });
  },

  clearSparkline(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const sparklines = new Map(s.sparkline.sparklines);
      sparklines.delete(addrKey(addr));
      return { ...s, sparkline: { sparklines } };
    });
  },

  clearSparklinesInRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const sparklines = new Map(s.sparkline.sparklines);
      for (const key of sparklines.keys()) {
        if (keyInRange(key, range)) sparklines.delete(key);
      }
      if (sparklines.size === s.sparkline.sparklines.size) return s;
      return { ...s, sparkline: { sparklines } };
    });
  },

  upsertChart(store: SpreadsheetStore, chart: SessionChart): void {
    store.setState((s) => {
      const next = s.charts.charts.filter((c) => c.id !== chart.id);
      return { ...s, charts: { charts: [...next, { ...chart, source: { ...chart.source } }] } };
    });
  },

  removeChart(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const next = s.charts.charts.filter((c) => c.id !== id);
      if (next.length === s.charts.charts.length) return s;
      return { ...s, charts: { charts: next } };
    });
  },

  updateChart(store: SpreadsheetStore, id: string, patch: Partial<Omit<SessionChart, 'id'>>): void {
    store.setState((s) => {
      let changed = false;
      const next = s.charts.charts.map((chart) => {
        if (chart.id !== id) return chart;
        changed = true;
        return { ...chart, ...patch, source: patch.source ? { ...patch.source } : chart.source };
      });
      if (!changed) return s;
      return { ...s, charts: { charts: next } };
    });
  },

  clearChartsInRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const next = s.charts.charts.filter((chart) => !rangesIntersect(chart.source, range));
      if (next.length === s.charts.charts.length) return s;
      return { ...s, charts: { charts: next } };
    });
  },

  /** Pin `addr` to the Watch Window. No-op when the same sheet/row/col is
   *  already watched — duplicate entries would render redundant rows. */
  addWatch(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const exists = s.watch.watches.some(
        (w) => w.sheet === addr.sheet && w.row === addr.row && w.col === addr.col,
      );
      if (exists) return s;
      return {
        ...s,
        watch: { watches: [...s.watch.watches, { ...addr }] },
      };
    });
  },

  /** Unpin `addr` from the Watch Window. No-op when not present. */
  removeWatch(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const next = s.watch.watches.filter(
        (w) => !(w.sheet === addr.sheet && w.row === addr.row && w.col === addr.col),
      );
      if (next.length === s.watch.watches.length) return s;
      return { ...s, watch: { watches: next } };
    });
  },

  /** Drop every watched cell. */
  clearWatches(store: SpreadsheetStore): void {
    store.setState((s) => (s.watch.watches.length === 0 ? s : { ...s, watch: { watches: [] } }));
  },

  /** Show or hide the Watch Window panel. */
  setWatchPanelOpen(store: SpreadsheetStore, open: boolean): void {
    store.setState((s) =>
      s.ui.watchPanelOpen === open ? s : { ...s, ui: { ...s.ui, watchPanelOpen: open } },
    );
  },

  /** Append a trace arrow to the visible set. Duplicates (same kind +
   *  identical endpoints) are dropped so repeated `tracePrecedents()` calls
   *  on the same active cell don't pile up overlapping arrows. */
  addTrace(store: SpreadsheetStore, item: TraceArrow): void {
    store.setState((s) => {
      const exists = s.traces.items.some(
        (t) =>
          t.kind === item.kind &&
          t.from.sheet === item.from.sheet &&
          t.from.row === item.from.row &&
          t.from.col === item.from.col &&
          t.to.sheet === item.to.sheet &&
          t.to.row === item.to.row &&
          t.to.col === item.to.col,
      );
      if (exists) return s;
      return {
        ...s,
        traces: { items: [...s.traces.items, { kind: item.kind, from: item.from, to: item.to }] },
      };
    });
  },

  /** Empty the trace-arrow set. */
  clearTraces(store: SpreadsheetStore): void {
    store.setState((s) => (s.traces.items.length === 0 ? s : { ...s, traces: { items: [] } }));
  },

  /** Suppress the error-indicator triangle for `addr` for the rest of the
   *  session. Idempotent — re-ignoring a cell is a no-op. NOT history-tracked. */
  ignoreError(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const key = addrKey(addr);
      if (s.errorIndicators.ignoredErrors.has(key)) return s;
      const next = new Set(s.errorIndicators.ignoredErrors);
      next.add(key);
      return { ...s, errorIndicators: { ignoredErrors: next } };
    });
  },

  /** Re-enable the error-indicator triangle for `addr` when it was ignored. */
  unignoreError(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const key = addrKey(addr);
      if (!s.errorIndicators.ignoredErrors.has(key)) return s;
      const next = new Set(s.errorIndicators.ignoredErrors);
      next.delete(key);
      return { ...s, errorIndicators: { ignoredErrors: next } };
    });
  },

  /** Drop every ignored-error suppression. */
  clearIgnoredErrors(store: SpreadsheetStore): void {
    store.setState((s) =>
      s.errorIndicators.ignoredErrors.size === 0
        ? s
        : { ...s, errorIndicators: { ignoredErrors: new Set() } },
    );
  },

  /** Append a fresh slicer to the slice. Caller is responsible for picking a
   *  unique `id` — duplicates are rejected (the older spec wins). */
  addSlicer(store: SpreadsheetStore, spec: SlicerSpec): void {
    store.setState((s) => {
      if (s.slicers.slicers.some((sp) => sp.id === spec.id)) return s;
      return {
        ...s,
        slicers: { slicers: [...s.slicers.slicers, { ...spec, selected: [...spec.selected] }] },
      };
    });
  },

  /** Remove the slicer with `id`. No-op when not present. */
  removeSlicer(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const next = s.slicers.slicers.filter((sp) => sp.id !== id);
      if (next.length === s.slicers.slicers.length) return s;
      return { ...s, slicers: { slicers: next } };
    });
  },

  /** Merge a partial patch onto the slicer with `id`. Skips the `id` field —
   *  ids stay immutable. */
  updateSlicer(store: SpreadsheetStore, id: string, patch: Partial<Omit<SlicerSpec, 'id'>>): void {
    store.setState((s) => {
      let changed = false;
      const next = s.slicers.slicers.map((sp) => {
        if (sp.id !== id) return sp;
        changed = true;
        return {
          ...sp,
          ...patch,
          selected: patch.selected ? [...patch.selected] : sp.selected,
        };
      });
      if (!changed) return s;
      return { ...s, slicers: { slicers: next } };
    });
  },

  /** Replace the chip selection for slicer `id`. Empty array = "include all". */
  setSlicerSelected(store: SpreadsheetStore, id: string, values: readonly string[]): void {
    store.setState((s) => {
      let changed = false;
      const next = s.slicers.slicers.map((sp) => {
        if (sp.id !== id) return sp;
        changed = true;
        return { ...sp, selected: [...values] };
      });
      if (!changed) return s;
      return { ...s, slicers: { slicers: next } };
    });
  },

  /** Toggle sheet-level protection for `sheet`. When `on` is `true` the sheet
   *  enters protected mode and the command layer gates writes against
   *  per-cell `locked` flags. The optional `password` is stored verbatim —
   *  v1 does NOT enforce it (no challenge dialog) but the value round-trips
   *  through the slice so callers can persist it. NOT history-tracked. */
  setSheetProtected(
    store: SpreadsheetStore,
    sheet: number,
    on: boolean,
    options?: { password?: string },
  ): void {
    store.setState((s) => {
      const next = new Map(s.protection.protectedSheets);
      if (on) {
        const entry: { password?: string } = {};
        if (options?.password !== undefined) entry.password = options.password;
        next.set(sheet, entry);
      } else if (next.has(sheet)) {
        next.delete(sheet);
      } else {
        return s;
      }
      return { ...s, protection: { protectedSheets: next } };
    });
  },

  /** Merge a partial patch into the page-setup for `sheet`. Pass `null` to
   *  reset that sheet back to defaults. The merge is shallow except for
   *  `margins`, which is deep-merged — so a patch like `{ margins: { top: 1 } }`
   *  preserves the other three sides. */
  setPageSetup(store: SpreadsheetStore, sheet: number, patch: Partial<PageSetup> | null): void {
    store.setState((s) => {
      const setupBySheet = new Map(s.pageSetup.setupBySheet);
      if (patch === null) {
        setupBySheet.delete(sheet);
      } else {
        const prev = setupBySheet.get(sheet) ?? defaultPageSetup();
        const next: PageSetup = { ...prev, ...patch };
        if (patch.margins) {
          next.margins = { ...prev.margins, ...patch.margins };
        }
        setupBySheet.set(sheet, next);
      }
      return { ...s, pageSetup: { setupBySheet } };
    });
  },

  /** Remove any merges that intersect the range. */
  unmergeRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const byAnchor = new Map(s.merges.byAnchor);
      const byCell = new Map(s.merges.byCell);
      for (const [anchorKey, r] of byAnchor) {
        if (
          r.sheet === range.sheet &&
          !(r.r1 < range.r0 || r.r0 > range.r1 || r.c1 < range.c0 || r.c0 > range.c1)
        ) {
          byAnchor.delete(anchorKey);
          for (let row = r.r0; row <= r.r1; row += 1) {
            for (let col = r.c0; col <= r.c1; col += 1) {
              byCell.delete(addrKey({ sheet: r.sheet, row, col }));
            }
          }
        }
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  },

  upsertTableOverlay(store: SpreadsheetStore, next: TableOverlay): void {
    store.setState((s) => {
      const filtered = s.tables.tables.filter((t) => t.id !== next.id);
      return { ...s, tables: { tables: [...filtered, next] } };
    });
  },

  removeTableOverlay(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const next = s.tables.tables.filter((t) => t.source === 'engine' || t.id !== id);
      if (next.length === s.tables.tables.length) return s;
      return { ...s, tables: { tables: next } };
    });
  },

  clearTableOverlaysInRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const next = s.tables.tables.filter(
        (t) => t.source === 'engine' || !rangesIntersect(t.range, range),
      );
      if (next.length === s.tables.tables.length) return s;
      return { ...s, tables: { tables: next } };
    });
  },

  replaceEngineTableOverlays(store: SpreadsheetStore, tables: readonly TableOverlay[]): void {
    store.setState((s) => {
      const session = s.tables.tables.filter((t) => t.source !== 'engine');
      return { ...s, tables: { tables: [...tables, ...session] } };
    });
  },

  upsertSheetView(store: SpreadsheetStore, view: SheetView): void {
    store.setState((s) => {
      const next = s.sheetViews.views.filter((v) => v.id !== view.id);
      return { ...s, sheetViews: { ...s.sheetViews, views: [...next, view] } };
    });
  },

  removeSheetView(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const views = s.sheetViews.views.filter((v) => v.id !== id);
      if (views.length === s.sheetViews.views.length) return s;
      return {
        ...s,
        sheetViews: {
          views,
          activeViewId: s.sheetViews.activeViewId === id ? null : s.sheetViews.activeViewId,
        },
      };
    });
  },

  applySheetViewPatch(
    store: SpreadsheetStore,
    patch: SheetViewPatch,
    activeViewId: string | null = null,
  ): void {
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        freezeRows: Math.max(0, Math.floor(patch.freezeRows)),
        freezeCols: Math.max(0, Math.floor(patch.freezeCols)),
        hiddenRows: new Set(patch.hiddenRows),
        hiddenCols: new Set(patch.hiddenCols),
      },
      viewport: {
        ...s.viewport,
        rowStart: Math.max(Math.max(0, Math.floor(patch.freezeRows)), s.viewport.rowStart),
        colStart: Math.max(Math.max(0, Math.floor(patch.freezeCols)), s.viewport.colStart),
      },
      ui: { ...s.ui, filterRange: patch.filterRange },
      sheetViews: { ...s.sheetViews, activeViewId },
    }));
  },
};

/** Pure read-helper: return the page-setup for `sheet`, falling back to
 *  `defaultPageSetup()` when no entry exists. Always returns a fully-populated
 *  record so callers can read every field without optional-chaining. */
export function getPageSetup(state: State, sheet: number): PageSetup {
  const entry = state.pageSetup.setupBySheet.get(sheet);
  if (!entry) return defaultPageSetup();
  // Merge missing fields onto the default so callers always see a complete
  // record even if `setPageSetup` was called with a sparse patch.
  const def = defaultPageSetup();
  return {
    ...def,
    ...entry,
    margins: { ...def.margins, ...entry.margins },
  };
}
