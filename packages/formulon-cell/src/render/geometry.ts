import { pageLayoutGaps } from '../commands/pagination.js';
import type { Range } from '../engine/types.js';
import type { LayoutSlice, State, UiSlice, ViewportSlice } from '../store/store.js';

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export type HitZone =
  | { kind: 'cell'; row: number; col: number }
  | { kind: 'col-header'; col: number }
  | { kind: 'row-header'; row: number }
  | { kind: 'corner' }
  | { kind: 'col-resize'; col: number }
  | { kind: 'row-resize'; row: number }
  | { kind: 'col-filter-btn'; col: number };

const RESIZE_SLACK = 4;
/** Chevron sits flush with the right edge of the header cell, just inboard of
 *  the col-resize handle. ~14px wide for a comfortable click target. */
export const FILTER_BTN_SIZE = 14;
export const FILTER_BTN_INSET = 4;

function viewportZoom(viewport?: ViewportSlice): number {
  return viewport?.zoom && Number.isFinite(viewport.zoom) ? viewport.zoom : 1;
}

/** Spreadsheet-style column letter ("A", "Z", "AA", "AB", ...). */
export function colLabel(idx: number): string {
  let n = idx;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

export function colWidth(layout: LayoutSlice, col: number, viewport?: ViewportSlice): number {
  if (layout.hiddenCols.has(col)) return 0;
  return (layout.colWidths.get(col) ?? layout.defaultColWidth) * viewportZoom(viewport);
}

export function rowHeight(layout: LayoutSlice, row: number, viewport?: ViewportSlice): number {
  if (layout.hiddenRows.has(row)) return 0;
  return (layout.rowHeights.get(row) ?? layout.defaultRowHeight) * viewportZoom(viewport);
}

/** Return the geometry layout for the current view. Excel's "Headings" view
 *  toggle removes the row/column header rails from the sheet viewport; the
 *  underlying layout metrics are kept intact so turning headings back on
 *  restores the original header chrome. */
/** The renderer's projection of the layout slice. Beyond the store's layout it
 *  carries the two facts needed to place a column on screen rather than in the
 *  abstract: whether the sheet runs right-to-left, and the width it mirrors
 *  about. Every function below that emits or consumes a screen x takes this
 *  type, so a caller cannot accidentally hand it a raw `state.layout` and get
 *  left-to-right coordinates on a right-to-left sheet. */
export interface ViewLayout extends PagedLayout {
  /** `<sheetView rightToLeft>`: column A sits at the right edge. */
  rtl: boolean;
  /** Grid width in CSS pixels — the axis `rtl` mirrors about. */
  viewWidth: number;
}

/** A layout that may carry Page Layout view's page gutters.
 *
 *  In that view the grid is interrupted by paper: each printed page is framed
 *  by its margins and separated from the next by a gutter. Rather than teach
 *  every paint pass about pages, the gutters are folded into the axis as extra
 *  leading pixels before the row / column that opens a page, so all existing
 *  geometry — cell rects, hit-testing, the inline editor's anchor — keeps
 *  working unchanged. Both maps are absent in every other view, and their
 *  values are already multiplied by the viewport zoom. */
export interface PagedLayout extends LayoutSlice {
  pageGapRows?: ReadonlyMap<number, number>;
  pageGapCols?: ReadonlyMap<number, number>;
  /** Height / width of the Page Layout rulers, in screen pixels. The rulers
   *  sit outboard of the outline gutters, so reserving the band here shifts
   *  the headers and the whole grid without any painter having to know why. */
  rulerRowHeight?: number;
  rulerColWidth?: number;
}

/** Height of the horizontal ruler band; zero outside Page Layout view. */
export function rulerTop(layout: PagedLayout): number {
  return layout.rulerRowHeight ?? 0;
}

/** Width of the vertical ruler band; zero outside Page Layout view. */
export function rulerLeft(layout: PagedLayout): number {
  return layout.rulerColWidth ?? 0;
}

/** Leading gutter before `col`, in screen pixels. Zero outside Page Layout. */
export function colGap(layout: PagedLayout, col: number): number {
  return layout.pageGapCols?.get(col) ?? 0;
}

/** Leading gutter above `row`, in screen pixels. Zero outside Page Layout. */
export function rowGap(layout: PagedLayout, row: number): number {
  return layout.pageGapRows?.get(row) ?? 0;
}

/** Application state as the renderer sees it: the same slices, with the layout
 *  already projected for the current view. Paint passes take this so they
 *  cannot reach a raw `state.layout`. */
export type ViewState = Omit<State, 'layout'> & { layout: ViewLayout };

export function layoutForView(
  state: LayoutForViewInput,
  overrides: { showHeaders?: boolean } = {},
): ViewLayout {
  const showHeaders = overrides.showHeaders ?? state.ui.showHeaders !== false;
  const base = {
    ...state.layout,
    rtl: state.ui.rightToLeft === true,
    viewWidth: state.viewport.widthPx,
    ...pageGapsFor(state),
  };
  if (showHeaders) return base;
  return { ...base, headerColWidth: 0, headerRowHeight: 0 };
}

/** What `layoutForView` needs. The slices beyond layout/ui/viewport are only
 *  read in Page Layout view, so callers that never enter it — and the unit
 *  tests — may hand over the three-slice subset. */
export type LayoutForViewInput = {
  layout: LayoutSlice;
  ui: UiSlice;
  viewport: ViewportSlice;
} & Partial<State>;

function pageGapsFor(state: LayoutForViewInput): Partial<PagedLayout> {
  if (state.ui.workbookView !== 'pageLayout') return {};
  if (!state.data || !state.format || !state.merges || !state.pageSetup) return {};
  const viewport = state.viewport;
  return {
    ...pageLayoutGaps(state as State, state.data.sheetIndex, {
      throughRow: viewport.rowStart + viewport.rowCount,
      throughCol: viewport.colStart + viewport.colCount,
    }),
    rulerRowHeight: RULER_BAND,
    rulerColWidth: RULER_BAND,
  };
}

/** Thickness of either ruler band, in screen pixels. */
export const RULER_BAND = 18;

/** Screen x of a rect's trailing edge — the corner a range's bottom-corner
 *  affordances hang off, which the mirror moves to the physical left. */
export function trailingEdgeX(rect: Rect, rtl: boolean): number {
  return rtl ? rect.x : rect.x + rect.w;
}

/** Mirror a horizontal span within a cell rect. Lets a painter keep writing
 *  its offsets from the cell's left edge and still land on the trailing side
 *  when the sheet runs right-to-left. */
export function mirrorInRect(bounds: Rect, x: number, w: number, rtl: boolean): number {
  return rtl ? bounds.x + bounds.w - (x - bounds.x) - w : x;
}

/** Mirror a horizontal span about the viewport's right edge; identity in a
 *  left-to-right sheet. The transform is its own inverse, so the same call
 *  converts a laid-out x to a screen x and a screen x back again. */
export function mirrorX(layout: ViewLayout, x: number, w = 0): number {
  return layout.rtl ? layout.viewWidth - x - w : x;
}

/** Total left offset before the first data column. Includes the row-outline
 *  bracket gutter (when rows are grouped) plus the row-number header strip. */
export function gridOriginX(layout: PagedLayout): number {
  return rulerLeft(layout) + layout.outlineRowGutter + layout.headerColWidth;
}

/** Total top offset before the first data row. Includes the col-outline
 *  bracket gutter plus the col-letter header strip. */
export function gridOriginY(layout: PagedLayout): number {
  return rulerTop(layout) + layout.outlineColGutter + layout.headerRowHeight;
}

/** Total width occupied by frozen columns. Zero if no freeze. */
export function frozenColsWidth(layout: PagedLayout, viewport?: ViewportSlice): number {
  let w = 0;
  for (let c = 0; c < layout.freezeCols; c += 1)
    w += colGap(layout, c) + colWidth(layout, c, viewport);
  return w;
}

/** Total height occupied by frozen rows. Zero if no freeze. */
export function frozenRowsHeight(layout: PagedLayout, viewport?: ViewportSlice): number {
  let h = 0;
  for (let r = 0; r < layout.freezeRows; r += 1)
    h += rowGap(layout, r) + rowHeight(layout, r, viewport);
  return h;
}

/** Cumulative pixel x for a column, relative to data origin (excludes header).
 *  Frozen columns are positioned from the data origin; non-frozen columns sit
 *  to the right of the frozen band, offset by the body viewport scroll. */
export function colX(layout: PagedLayout, viewport: ViewportSlice, col: number): number {
  const fc = layout.freezeCols;
  const walk = (from: number, base: number): number => {
    let x = base;
    for (let c = from; c <= col; c += 1) {
      // The gutter that opens a page precedes its first column, so it counts
      // towards that column's own offset as well as every later one.
      x += colGap(layout, c);
      if (c < col) x += colWidth(layout, c, viewport);
    }
    return x;
  };
  if (col < fc) return walk(0, 0);
  return walk(Math.max(viewport.colStart, fc), frozenColsWidth(layout, viewport));
}

export function rowY(layout: PagedLayout, viewport: ViewportSlice, row: number): number {
  const fr = layout.freezeRows;
  const walk = (from: number, base: number): number => {
    let y = base;
    for (let r = from; r <= row; r += 1) {
      y += rowGap(layout, r);
      if (r < row) y += rowHeight(layout, r, viewport);
    }
    return y;
  };
  if (row < fr) return walk(0, 0);
  return walk(Math.max(viewport.rowStart, fr), frozenRowsHeight(layout, viewport));
}

export function cellRect(
  layout: ViewLayout,
  viewport: ViewportSlice,
  row: number,
  col: number,
): Rect {
  const x = gridOriginX(layout) + colX(layout, viewport, col);
  const y = gridOriginY(layout) + rowY(layout, viewport, row);
  const w = colWidth(layout, col, viewport);
  return { x: mirrorX(layout, x, w), y, w, h: rowHeight(layout, row, viewport) };
}

/** Like {@link cellRect}, but returns the true (possibly negative) offset for
 *  cells scrolled off the leading edge instead of clamping them to the data
 *  origin. `colX` / `rowY` clamp because paint passes never draw those cells;
 *  DOM overlays anchored to a cell — the inline editor — need the real offset
 *  so they scroll out of view with their cell rather than sticking to the
 *  first visible row/column. */
export function cellRectUnclamped(
  layout: ViewLayout,
  viewport: ViewportSlice,
  row: number,
  col: number,
): Rect {
  const rect = cellRect(layout, viewport, row, col);
  const colStart = Math.max(viewport.colStart, layout.freezeCols);
  // Walking back past a page gutter skips the gutter of the cell being
  // resolved: that one sits ahead of the cell, not between it and the edge.
  if (col >= layout.freezeCols && col < colStart) {
    let back = 0;
    for (let c = col; c < colStart; c += 1) {
      back += colWidth(layout, c, viewport) + (c > col ? colGap(layout, c) : 0);
    }
    rect.x += layout.rtl ? back : -back;
  }
  const rowStart = Math.max(viewport.rowStart, layout.freezeRows);
  if (row >= layout.freezeRows && row < rowStart) {
    for (let r = row; r < rowStart; r += 1) {
      rect.y -= rowHeight(layout, r, viewport) + (r > row ? rowGap(layout, r) : 0);
    }
  }
  return rect;
}

/** The boundary between the frozen/header chrome and the scrollable body band,
 *  in screen coordinates. A non-frozen cell that renders past this boundary —
 *  above it, or leading of it on the horizontal axis — has scrolled underneath
 *  the chrome. On a right-to-left sheet the horizontal boundary is the band's
 *  right edge, so "leading of it" means further right. */
export function bodyBandOrigin(
  layout: ViewLayout,
  viewport: ViewportSlice,
): { x: number; y: number } {
  return {
    x: mirrorX(layout, gridOriginX(layout) + frozenColsWidth(layout, viewport)),
    y: gridOriginY(layout) + frozenRowsHeight(layout, viewport),
  };
}

/** Hit-test a pointer position against the data area. Returns { row, col }
 *  inside the visible viewport, or null if the point is in a header / outside.
 *  Freeze-aware: a click in the frozen band resolves to a frozen row/col. */
export function hitTest(
  layout: ViewLayout,
  viewport: ViewportSlice,
  screenX: number,
  y: number,
): { row: number; col: number } | null {
  const x = mirrorX(layout, screenX);
  const ox = gridOriginX(layout);
  const oy = gridOriginY(layout);
  if (x < ox || y < oy) return null;
  const fc = layout.freezeCols;
  const fr = layout.freezeRows;
  const fcw = frozenColsWidth(layout, viewport);
  const frh = frozenRowsHeight(layout, viewport);

  // A pointer inside a page gutter belongs to no cell — the gutter is paper
  // margin, not sheet — so each step charges the gap before testing the cell
  // and bails out when the pointer never reaches the cell itself.
  let col: number;
  let cx = ox;
  if (fc > 0 && x < ox + fcw) {
    col = 0;
    while (col < fc) {
      cx += colGap(layout, col);
      if (x < cx) return null;
      const w = colWidth(layout, col, viewport);
      if (x < cx + w) break;
      cx += w;
      col += 1;
    }
    if (col >= fc) return null;
  } else {
    cx = ox + fcw;
    col = Math.max(viewport.colStart, fc);
    const end = viewport.colStart + viewport.colCount;
    while (col < end) {
      cx += colGap(layout, col);
      if (x < cx) return null;
      const w = colWidth(layout, col, viewport);
      if (x < cx + w) break;
      cx += w;
      col += 1;
    }
    if (col >= end) return null;
  }

  let row: number;
  let cy = oy;
  if (fr > 0 && y < oy + frh) {
    row = 0;
    while (row < fr) {
      cy += rowGap(layout, row);
      if (y < cy) return null;
      const h = rowHeight(layout, row, viewport);
      if (y < cy + h) break;
      cy += h;
      row += 1;
    }
    if (row >= fr) return null;
  } else {
    cy = oy + frh;
    row = Math.max(viewport.rowStart, fr);
    const end = viewport.rowStart + viewport.rowCount;
    while (row < end) {
      cy += rowGap(layout, row);
      if (y < cy) return null;
      const h = rowHeight(layout, row, viewport);
      if (y < cy + h) break;
      cy += h;
      row += 1;
    }
    if (row >= end) return null;
  }

  return { row, col };
}

/** Resolve which column index lies under x in the data area, and the pixel
 *  position of its right edge. Returns null when x is outside visible cols.
 *  Freeze-aware. */
function colAtX(
  layout: ViewLayout,
  viewport: ViewportSlice,
  x: number,
): { col: number; rightEdge: number; leftEdge: number } | null {
  const fc = layout.freezeCols;
  const fcw = frozenColsWidth(layout, viewport);
  const ox = gridOriginX(layout);
  if (fc > 0 && x < ox + fcw) {
    let cx = ox;
    for (let col = 0; col < fc; col += 1) {
      cx += colGap(layout, col);
      const w = colWidth(layout, col, viewport);
      if (x < cx + w) return { col, leftEdge: cx, rightEdge: cx + w };
      cx += w;
    }
    return null;
  }
  let cx = ox + fcw;
  let col = Math.max(viewport.colStart, fc);
  const end = viewport.colStart + viewport.colCount;
  while (col < end) {
    cx += colGap(layout, col);
    const w = colWidth(layout, col, viewport);
    if (x < cx + w) return { col, leftEdge: cx, rightEdge: cx + w };
    cx += w;
    col += 1;
  }
  return null;
}

function rowAtY(
  layout: PagedLayout,
  viewport: ViewportSlice,
  y: number,
): { row: number; bottomEdge: number; topEdge: number } | null {
  const fr = layout.freezeRows;
  const frh = frozenRowsHeight(layout, viewport);
  const oy = gridOriginY(layout);
  if (fr > 0 && y < oy + frh) {
    let cy = oy;
    for (let row = 0; row < fr; row += 1) {
      cy += rowGap(layout, row);
      const h = rowHeight(layout, row, viewport);
      if (y < cy + h) return { row, topEdge: cy, bottomEdge: cy + h };
      cy += h;
    }
    return null;
  }
  let cy = oy + frh;
  let row = Math.max(viewport.rowStart, fr);
  const end = viewport.rowStart + viewport.rowCount;
  while (row < end) {
    cy += rowGap(layout, row);
    const h = rowHeight(layout, row, viewport);
    if (y < cy + h) return { row, topEdge: cy, bottomEdge: cy + h };
    cy += h;
    row += 1;
  }
  return null;
}

/** Whether `col` is currently rendered (frozen band or scrolled body). */
export function isColVisible(layout: LayoutSlice, viewport: ViewportSlice, col: number): boolean {
  if (col < 0) return false;
  if (layout.hiddenCols.has(col)) return false;
  if (col < layout.freezeCols) return true;
  const start = Math.max(viewport.colStart, layout.freezeCols);
  return col >= start && col < viewport.colStart + viewport.colCount;
}

export function isRowVisible(layout: LayoutSlice, viewport: ViewportSlice, row: number): boolean {
  if (row < 0) return false;
  if (layout.hiddenRows.has(row)) return false;
  if (row < layout.freezeRows) return true;
  const start = Math.max(viewport.rowStart, layout.freezeRows);
  return row >= start && row < viewport.rowStart + viewport.rowCount;
}

/** Rich hit-test that resolves headers, resize edges, the corner chip, and
 *  cells. Returns null only when the point is past the last visible col/row.
 *  When `filterRange` is supplied, a chevron hot-zone is returned for the
 *  rightmost ~18px (excluding the resize slack) of headers inside the range. */
export function hitZone(
  layout: ViewLayout,
  viewport: ViewportSlice,
  screenX: number,
  y: number,
  filterRange?: Range | null,
  opts?: { resizeHandles?: boolean },
): HitZone | null {
  // Everything below reasons in laid-out space, where column A is always
  // leftmost. `colAtX` and `hitTest` do the same, so the pointer only has to
  // cross the mirror once.
  const x = mirrorX(layout, screenX);
  // Page Layout's rulers take the outermost band, outboard of the header
  // rails. They belong to the page, not to any row or column, so the header
  // fall-through below must not claim them.
  if (y < rulerTop(layout) || x < rulerLeft(layout)) return null;
  const resizeHandles = opts?.resizeHandles !== false;
  // Outline gutters sit outboard of the row/col header strips. Treat them as
  // header zones for now — the pointer layer routes outline-toggle clicks
  // through a dedicated hit-test before reaching this fall-through.
  const inHeaderCols = x < gridOriginX(layout);
  const inHeaderRows = y < gridOriginY(layout);

  if (inHeaderCols && inHeaderRows) return { kind: 'corner' };

  if (inHeaderRows) {
    const found = colAtX(layout, viewport, x);
    if (!found) return null;
    if (resizeHandles && found.rightEdge - x <= RESIZE_SLACK) {
      return { kind: 'col-resize', col: found.col };
    }
    if (
      resizeHandles &&
      x - found.leftEdge <= RESIZE_SLACK &&
      isColVisible(layout, viewport, found.col - 1)
    ) {
      return { kind: 'col-resize', col: found.col - 1 };
    }
    if (filterRange && found.col >= filterRange.c0 && found.col <= filterRange.c1) {
      const btnRight = found.rightEdge - RESIZE_SLACK;
      const btnLeft = btnRight - FILTER_BTN_SIZE;
      if (x >= btnLeft && x < btnRight) return { kind: 'col-filter-btn', col: found.col };
    }
    return { kind: 'col-header', col: found.col };
  }

  if (inHeaderCols) {
    const found = rowAtY(layout, viewport, y);
    if (!found) return null;
    if (resizeHandles && found.bottomEdge - y <= RESIZE_SLACK) {
      return { kind: 'row-resize', row: found.row };
    }
    if (
      resizeHandles &&
      y - found.topEdge <= RESIZE_SLACK &&
      isRowVisible(layout, viewport, found.row - 1)
    ) {
      return { kind: 'row-resize', row: found.row - 1 };
    }
    return { kind: 'row-header', row: found.row };
  }

  const cell = hitTest(layout, viewport, screenX, y);
  if (!cell) return null;
  return { kind: 'cell', row: cell.row, col: cell.col };
}

/** Screen x of a column's leading edge — physically its left edge on a
 *  left-to-right sheet, its right edge on a right-to-left one. */
export function colLeadingEdge(layout: ViewLayout, viewport: ViewportSlice, col: number): number {
  return mirrorX(layout, gridOriginX(layout) + colX(layout, viewport, col));
}

/** Return the absolute y of a row's top edge (header-inclusive coords). */
export function rowTopEdge(layout: LayoutSlice, viewport: ViewportSlice, row: number): number {
  return gridOriginY(layout) + rowY(layout, viewport, row);
}

/** Indices of every row currently rendered, in render order. Frozen first,
 *  then the body slice. Hidden rows are omitted. */
export function visibleRows(layout: LayoutSlice, viewport: ViewportSlice): number[] {
  const out: number[] = [];
  for (let r = 0; r < layout.freezeRows; r += 1) {
    if (!layout.hiddenRows.has(r)) out.push(r);
  }
  const start = Math.max(viewport.rowStart, layout.freezeRows);
  const end = viewport.rowStart + viewport.rowCount;
  for (let r = start; r < end; r += 1) {
    if (!layout.hiddenRows.has(r)) out.push(r);
  }
  return out;
}

export function visibleCols(layout: LayoutSlice, viewport: ViewportSlice): number[] {
  const out: number[] = [];
  for (let c = 0; c < layout.freezeCols; c += 1) {
    if (!layout.hiddenCols.has(c)) out.push(c);
  }
  const start = Math.max(viewport.colStart, layout.freezeCols);
  const end = viewport.colStart + viewport.colCount;
  for (let c = start; c < end; c += 1) {
    if (!layout.hiddenCols.has(c)) out.push(c);
  }
  return out;
}

/** Per-axis layout cache for one paint cycle. Replaces O(visibleAxis) loops
 *  inside cellRect with O(1) map lookups. Build once per paint via
 *  `buildColLayout` / `buildRowLayout`. */
export interface AxisLayout {
  /** Visible indices in render order: frozen band first, then body slice.
   *  Hidden indices are excluded. */
  visible: number[];
  /** Index → starting pixel, relative to the data origin (header excluded). */
  positionAt: Map<number, number>;
  /** Index → pixel size. Mirrors `colWidth` / `rowHeight` for visible indices. */
  sizeAt: Map<number, number>;
  /** Sum of frozen-band sizes. Matches `frozenColsWidth` / `frozenRowsHeight`. */
  frozenTotal: number;
  /** Index → leading page gutter, for the visible indices that open a page.
   *  Empty outside Page Layout view. `positionAt` already sits past the
   *  gutter, so painters that need to fill the paper read it from here. */
  gapAt: Map<number, number>;
}

export function buildColLayout(layout: PagedLayout, viewport: ViewportSlice): AxisLayout {
  const visible: number[] = [];
  const positionAt = new Map<number, number>();
  const sizeAt = new Map<number, number>();
  const gapAt = new Map<number, number>();

  let x = 0;
  const place = (c: number): void => {
    const gap = colGap(layout, c);
    x += gap;
    const w = colWidth(layout, c, viewport);
    if (w > 0) {
      visible.push(c);
      positionAt.set(c, x);
      sizeAt.set(c, w);
      if (gap > 0) gapAt.set(c, gap);
    }
    x += w;
  };

  for (let c = 0; c < layout.freezeCols; c += 1) place(c);
  const frozenTotal = x;

  const start = Math.max(viewport.colStart, layout.freezeCols);
  const end = viewport.colStart + viewport.colCount;
  for (let c = start; c < end; c += 1) place(c);

  return { visible, positionAt, sizeAt, frozenTotal, gapAt };
}

export function buildRowLayout(layout: PagedLayout, viewport: ViewportSlice): AxisLayout {
  const visible: number[] = [];
  const positionAt = new Map<number, number>();
  const sizeAt = new Map<number, number>();
  const gapAt = new Map<number, number>();

  let y = 0;
  const place = (r: number): void => {
    const gap = rowGap(layout, r);
    y += gap;
    const h = rowHeight(layout, r, viewport);
    if (h > 0) {
      visible.push(r);
      positionAt.set(r, y);
      sizeAt.set(r, h);
      if (gap > 0) gapAt.set(r, gap);
    }
    y += h;
  };

  for (let r = 0; r < layout.freezeRows; r += 1) place(r);
  const frozenTotal = y;

  const start = Math.max(viewport.rowStart, layout.freezeRows);
  const end = viewport.rowStart + viewport.rowCount;
  for (let r = start; r < end; r += 1) place(r);

  return { visible, positionAt, sizeAt, frozenTotal, gapAt };
}

/** Constant-time cellRect using precomputed AxisLayouts. Caller guarantees
 *  the (row, col) pair is in `cols.visible` × `rows.visible`; otherwise the
 *  rect is anchored at the data origin with the cell's nominal size. */
export function cellRectIn(
  layout: ViewLayout,
  cols: AxisLayout,
  rows: AxisLayout,
  row: number,
  col: number,
): Rect {
  const w = cols.sizeAt.get(col) ?? colWidth(layout, col);
  return {
    x: mirrorX(layout, gridOriginX(layout) + (cols.positionAt.get(col) ?? 0), w),
    y: gridOriginY(layout) + (rows.positionAt.get(row) ?? 0),
    w,
    h: rows.sizeAt.get(row) ?? rowHeight(layout, row),
  };
}

/** Up to four rectangles covering the visible portion of `range`, one per
 *  freeze quadrant the range overlaps. Returns an empty list if the range
 *  is entirely outside the visible area. */
export function rangeRects(
  layout: ViewLayout,
  viewport: ViewportSlice,
  range: { r0: number; r1: number; c0: number; c1: number },
): Rect[] {
  const fr = layout.freezeRows;
  const fc = layout.freezeCols;
  const lastRow = viewport.rowStart + viewport.rowCount - 1;
  const lastCol = viewport.colStart + viewport.colCount - 1;

  const rowSegs: [number, number][] = [];
  if (fr > 0 && range.r0 < fr) {
    rowSegs.push([range.r0, Math.min(range.r1, fr - 1)]);
  }
  const bodyRowStart = Math.max(viewport.rowStart, fr);
  if (range.r1 >= bodyRowStart && range.r0 <= lastRow) {
    rowSegs.push([Math.max(range.r0, bodyRowStart), Math.min(range.r1, lastRow)]);
  }

  const colSegs: [number, number][] = [];
  if (fc > 0 && range.c0 < fc) {
    colSegs.push([range.c0, Math.min(range.c1, fc - 1)]);
  }
  const bodyColStart = Math.max(viewport.colStart, fc);
  if (range.c1 >= bodyColStart && range.c0 <= lastCol) {
    colSegs.push([Math.max(range.c0, bodyColStart), Math.min(range.c1, lastCol)]);
  }

  const rects: Rect[] = [];
  for (const [r0, r1] of rowSegs) {
    for (const [c0, c1] of colSegs) {
      const tl = cellRect(layout, viewport, r0, c0);
      const br = cellRect(layout, viewport, r1, c1);
      // The two corner rects come back already mirrored, so on a right-to-left
      // sheet c0's rect is the rightmost one. Span from whichever is leading.
      const left = Math.min(tl.x, br.x);
      const right = Math.max(tl.x + tl.w, br.x + br.w);
      rects.push({ x: left, y: tl.y, w: right - left, h: br.y + br.h - tl.y });
    }
  }
  return rects;
}
