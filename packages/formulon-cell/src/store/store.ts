import { createStore } from 'zustand/vanilla';
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
 *  carries an Excel-style style + optional color. The full OOXML repertoire
 *  is supported; Excel's 13 ordinals map to these names verbatim. */
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
  /** Diagonal — Excel supports both directions; `\` runs top-left → bottom-right
   *  and `/` runs bottom-left → top-right. */
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
  /** Vertical alignment. Excel default is 'bottom'. */
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
  /** Allow empty input regardless of constraint. Default true (Excel parity). */
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

/** Conditional formatting rule. Evaluated by the renderer against numeric
 *  cell values; non-numeric cells are skipped. */
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

/** Cells the user has pinned in the Watch Window. Session-only — Excel
 *  parity: watches don't survive workbook close, and they aren't recorded
 *  in the undo stack. Order is insertion order. */
export interface WatchSlice {
  watches: readonly Addr[];
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
  watch: WatchSlice;
}

const initialAddr = (sheet = 0): Addr => ({ sheet, row: 0, col: 0 });
const initialRange = (sheet = 0): Range => ({ sheet, r0: 0, c0: 0, r1: 0, c1: 0 });

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
      defaultColWidth: 104,
      defaultRowHeight: 28,
      headerColWidth: 52,
      headerRowHeight: 30,
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
    watch: { watches: [] },
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

  setFillPreview(store: SpreadsheetStore, range: Range | null): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, fillPreview: range } }));
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
};
