import type {
  CustomTableStyle,
  PivotTableStyleAssignment,
  TableOverlay,
} from '../commands/format-as-table.js';
import type { SheetView } from '../commands/sheet-views.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import type { CellFormat, CustomCellStyle } from './types/cell-format.js';
import type { ConditionalRule } from './types/conditional-rule.js';
import type { PageSetup } from './types/page-setup.js';

export type * from './types/cell-format.js';
export type * from './types/conditional-rule.js';
export type * from './types/page-setup.js';
export { defaultPageSetup } from './types/page-setup.js';

export type EditorMode =
  | { kind: 'idle' }
  | { kind: 'enter'; raw: string }
  | { kind: 'edit'; raw: string; caret: number };

export interface ViewportSlice {
  /** First row visible (zero-indexed). */
  rowStart: number;
  rowCount: number;
  colStart: number;
  colCount: number;
  zoom: number;
  /** Active navigation policy's fixed view rectangle, when an embedded host
   *  has installed one. This is session state only; workbook data and
   *  calculation remain unrestricted. */
  navigationRange?: Range;
  /** Grid canvas width in CSS pixels. Published by the renderer because it is
   *  the axis a right-to-left sheet mirrors about — geometry cannot place a
   *  column from the right edge without knowing where that edge is. Zero
   *  before the first measure, which reads as "no mirror yet". */
  widthPx: number;
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
  /** The subset of `hiddenSheets` a workbook marked very hidden. A spreadsheet
   *  leaves these out of its Unhide list as well as its tab bar, which is how
   *  a settings or lookup sheet is kept out of a user's reach. Every member is
   *  also in `hiddenSheets`. */
  veryHiddenSheets: Set<number>;
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

export interface PendingFormat {
  addr: Addr;
  format: Partial<CellFormat>;
}

/** Which clipboard command raised the marquee. */
export type CopyMode = 'copy' | 'cut';

export interface UiSlice {
  editor: EditorMode;
  hover: Addr | null;
  /** Excel-style input format staged for a single empty active cell. It is
   *  committed only when the user enters a value in that cell and is cleared
   *  on cancel or navigation away. */
  pendingFormat?: PendingFormat | null;
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
  /** Whether the marquee came from a copy or a cut. A copy survives pasting so
   *  the same source can be pasted repeatedly; a cut is consumed by its first
   *  paste. Null whenever there is no marquee. */
  copyMode?: CopyMode | null;
  /** Monotonic generation for the internal clipboard marquee. Unlike the
   *  marquee coordinates themselves, this changes whenever a public copy
   *  mutator replaces or clears the session, so async clipboard fallbacks can
   *  tell a stale payload from the live one. */
  copyRevision?: number;
  /** When false, the renderer skips drawing inter-cell hairline gridlines. */
  showGridLines: boolean;
  /** When false, the renderer hides the row-number / column-letter strips. */
  showHeaders: boolean;
  /** When false, numeric zero values are painted as blank cells. */
  showZeros: boolean;
  /** Sheet direction. When true, column A sits at the right edge and columns
   *  advance leftward — the `<sheetView rightToLeft>` flag. Governs the grid's
   *  own axis, not the direction of text inside a cell (`CellFormat.textDirection`). */
  rightToLeft: boolean;
  /** When true, formula cells display the formula text instead of the
   *  evaluated value. Equivalent to the desktop-spreadsheet "Show Formulas" (Ctrl+`). */
  showFormulas: boolean;
  /** Excel End mode: the next arrow key jumps to the edge of the current data region. */
  endMode: boolean;
  /** Workbook view mode surfaced by View > Workbook Views. Page Layout folds
   *  the page margins into the grid geometry and Page Break Preview overlays
   *  the page boundaries; chrome also stamps the mode on the host so themes
   *  and wrappers can react. */
  workbookView: WorkbookViewMode;
  /** Zoom in force before Page Break Preview pulled the sheet out to fit a
   *  whole page. Restored when the view closes; null outside the preview. */
  zoomBeforePreview?: number | null;
  /** Screen position of the page boundary currently being dragged in Page
   *  Break Preview. Purely transient: the pages reflow only when the pointer
   *  is released, so until then this is the whole of the drag feedback. */
  pageBreakDrag?: { axis: 'row' | 'col'; position: number } | null;
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
  /** Value filter: rows whose column value key is in this set are hidden. */
  hiddenValues: string[];
  /** Condition filter (number/text comparison). When present the criterion is
   *  evaluated per row and `hiddenValues` is ignored. Persisting it lets
   *  Reapply, sheet-view snapshots, and cross-column AND recompute honour
   *  condition filters the same way as value filters. */
  condition?: FilterConditionSpec;
  /** Color filter. When present, rows survive only when the target cell's fill
   *  or font color matches `color`. */
  color?: FilterColorSpec;
}

export interface FilterColorSpec {
  kind: 'cellColor' | 'fontColor';
  color: string;
}

/** A number/text comparison used by an AutoFilter condition filter. Mirrors
 *  the operator set of the filter dropdown. */
export interface FilterConditionSpec {
  op:
    | 'equals'
    | 'notEquals'
    | 'contains'
    | 'notContains'
    | 'greaterThan'
    | 'greaterThanOrEqual'
    | 'lessThan'
    | 'lessThanOrEqual';
  value: string;
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

export type SessionShapeKind =
  | 'rectangle'
  | 'rounded-rectangle'
  | 'oval'
  | 'triangle'
  | 'diamond'
  | 'line'
  | 'arrow';

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
  radius?: number;
  lineWidth?: number;
  opacity?: number;
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

export interface PageSetupSlice {
  /** Per-sheet page-setup map, keyed by sheet index. Sheets without an
   *  entry fall back to `defaultPageSetup()`. History-tracked. */
  setupBySheet: Map<number, PageSetup>;
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
 *  its index; the value records the supplied legacy password plus the
 *  permission flags mirrored from `<sheetProtection>`. Cell-level locks live
 *  on `CellFormat.locked`; this slice owns the sheet-side protection state. */
export interface SheetProtectionPermissions {
  objects?: boolean;
  scenarios?: boolean;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
}

export interface SheetProtectionPasswordHash {
  algorithmName: string;
  hashValue: string;
  saltValue: string;
  spinCount: number;
}

export interface SheetProtectionState {
  password?: string;
  passwordHash?: SheetProtectionPasswordHash;
  permissions?: SheetProtectionPermissions;
}

export interface ProtectionSlice {
  protectedSheets: Map<number, SheetProtectionState>;
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
