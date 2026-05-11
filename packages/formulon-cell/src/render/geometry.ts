import type { Range } from '../engine/types.js';
import type { LayoutSlice, ViewportSlice } from '../store/store.js';

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

/** Total left offset before the first data column. Includes the row-outline
 *  bracket gutter (when rows are grouped) plus the row-number header strip. */
export function gridOriginX(layout: LayoutSlice): number {
  return layout.outlineRowGutter + layout.headerColWidth;
}

/** Total top offset before the first data row. Includes the col-outline
 *  bracket gutter plus the col-letter header strip. */
export function gridOriginY(layout: LayoutSlice): number {
  return layout.outlineColGutter + layout.headerRowHeight;
}

/** Total width occupied by frozen columns. Zero if no freeze. */
export function frozenColsWidth(layout: LayoutSlice, viewport?: ViewportSlice): number {
  let w = 0;
  for (let c = 0; c < layout.freezeCols; c += 1) w += colWidth(layout, c, viewport);
  return w;
}

/** Total height occupied by frozen rows. Zero if no freeze. */
export function frozenRowsHeight(layout: LayoutSlice, viewport?: ViewportSlice): number {
  let h = 0;
  for (let r = 0; r < layout.freezeRows; r += 1) h += rowHeight(layout, r, viewport);
  return h;
}

/** Cumulative pixel x for a column, relative to data origin (excludes header).
 *  Frozen columns are positioned from the data origin; non-frozen columns sit
 *  to the right of the frozen band, offset by the body viewport scroll. */
export function colX(layout: LayoutSlice, viewport: ViewportSlice, col: number): number {
  const fc = layout.freezeCols;
  if (col < fc) {
    let x = 0;
    for (let c = 0; c < col; c += 1) x += colWidth(layout, c, viewport);
    return x;
  }
  let x = frozenColsWidth(layout, viewport);
  const start = Math.max(viewport.colStart, fc);
  for (let c = start; c < col; c += 1) x += colWidth(layout, c, viewport);
  return x;
}

export function rowY(layout: LayoutSlice, viewport: ViewportSlice, row: number): number {
  const fr = layout.freezeRows;
  if (row < fr) {
    let y = 0;
    for (let r = 0; r < row; r += 1) y += rowHeight(layout, r, viewport);
    return y;
  }
  let y = frozenRowsHeight(layout, viewport);
  const start = Math.max(viewport.rowStart, fr);
  for (let r = start; r < row; r += 1) y += rowHeight(layout, r, viewport);
  return y;
}

export function cellRect(
  layout: LayoutSlice,
  viewport: ViewportSlice,
  row: number,
  col: number,
): Rect {
  const x = gridOriginX(layout) + colX(layout, viewport, col);
  const y = gridOriginY(layout) + rowY(layout, viewport, row);
  return { x, y, w: colWidth(layout, col, viewport), h: rowHeight(layout, row, viewport) };
}

/** Hit-test a pointer position against the data area. Returns { row, col }
 *  inside the visible viewport, or null if the point is in a header / outside.
 *  Freeze-aware: a click in the frozen band resolves to a frozen row/col. */
export function hitTest(
  layout: LayoutSlice,
  viewport: ViewportSlice,
  x: number,
  y: number,
): { row: number; col: number } | null {
  const ox = gridOriginX(layout);
  const oy = gridOriginY(layout);
  if (x < ox || y < oy) return null;
  const fc = layout.freezeCols;
  const fr = layout.freezeRows;
  const fcw = frozenColsWidth(layout, viewport);
  const frh = frozenRowsHeight(layout, viewport);

  let col: number;
  let cx = ox;
  if (fc > 0 && x < ox + fcw) {
    col = 0;
    while (col < fc) {
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
  layout: LayoutSlice,
  viewport: ViewportSlice,
  x: number,
): { col: number; rightEdge: number; leftEdge: number } | null {
  const fc = layout.freezeCols;
  const fcw = frozenColsWidth(layout, viewport);
  const ox = gridOriginX(layout);
  if (fc > 0 && x < ox + fcw) {
    let cx = ox;
    for (let col = 0; col < fc; col += 1) {
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
    const w = colWidth(layout, col, viewport);
    if (x < cx + w) return { col, leftEdge: cx, rightEdge: cx + w };
    cx += w;
    col += 1;
  }
  return null;
}

function rowAtY(
  layout: LayoutSlice,
  viewport: ViewportSlice,
  y: number,
): { row: number; bottomEdge: number; topEdge: number } | null {
  const fr = layout.freezeRows;
  const frh = frozenRowsHeight(layout, viewport);
  const oy = gridOriginY(layout);
  if (fr > 0 && y < oy + frh) {
    let cy = oy;
    for (let row = 0; row < fr; row += 1) {
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
  layout: LayoutSlice,
  viewport: ViewportSlice,
  x: number,
  y: number,
  filterRange?: Range | null,
): HitZone | null {
  // Outline gutters sit outboard of the row/col header strips. Treat them as
  // header zones for now — the pointer layer routes outline-toggle clicks
  // through a dedicated hit-test before reaching this fall-through.
  const inHeaderCols = x < gridOriginX(layout);
  const inHeaderRows = y < gridOriginY(layout);

  if (inHeaderCols && inHeaderRows) return { kind: 'corner' };

  if (inHeaderRows) {
    const found = colAtX(layout, viewport, x);
    if (!found) return null;
    if (found.rightEdge - x <= RESIZE_SLACK) return { kind: 'col-resize', col: found.col };
    if (x - found.leftEdge <= RESIZE_SLACK && isColVisible(layout, viewport, found.col - 1)) {
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
    if (found.bottomEdge - y <= RESIZE_SLACK) return { kind: 'row-resize', row: found.row };
    if (y - found.topEdge <= RESIZE_SLACK && isRowVisible(layout, viewport, found.row - 1)) {
      return { kind: 'row-resize', row: found.row - 1 };
    }
    return { kind: 'row-header', row: found.row };
  }

  const cell = hitTest(layout, viewport, x, y);
  if (!cell) return null;
  return { kind: 'cell', row: cell.row, col: cell.col };
}

/** Return the absolute x of a column's left edge (header-inclusive coords). */
export function colLeftEdge(layout: LayoutSlice, viewport: ViewportSlice, col: number): number {
  return gridOriginX(layout) + colX(layout, viewport, col);
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
}

export function buildColLayout(layout: LayoutSlice, viewport: ViewportSlice): AxisLayout {
  const visible: number[] = [];
  const positionAt = new Map<number, number>();
  const sizeAt = new Map<number, number>();

  let x = 0;
  for (let c = 0; c < layout.freezeCols; c += 1) {
    const w = colWidth(layout, c, viewport);
    if (w > 0) {
      visible.push(c);
      positionAt.set(c, x);
      sizeAt.set(c, w);
    }
    x += w;
  }
  const frozenTotal = x;

  const start = Math.max(viewport.colStart, layout.freezeCols);
  const end = viewport.colStart + viewport.colCount;
  for (let c = start; c < end; c += 1) {
    const w = colWidth(layout, c, viewport);
    if (w > 0) {
      visible.push(c);
      positionAt.set(c, x);
      sizeAt.set(c, w);
    }
    x += w;
  }

  return { visible, positionAt, sizeAt, frozenTotal };
}

export function buildRowLayout(layout: LayoutSlice, viewport: ViewportSlice): AxisLayout {
  const visible: number[] = [];
  const positionAt = new Map<number, number>();
  const sizeAt = new Map<number, number>();

  let y = 0;
  for (let r = 0; r < layout.freezeRows; r += 1) {
    const h = rowHeight(layout, r, viewport);
    if (h > 0) {
      visible.push(r);
      positionAt.set(r, y);
      sizeAt.set(r, h);
    }
    y += h;
  }
  const frozenTotal = y;

  const start = Math.max(viewport.rowStart, layout.freezeRows);
  const end = viewport.rowStart + viewport.rowCount;
  for (let r = start; r < end; r += 1) {
    const h = rowHeight(layout, r, viewport);
    if (h > 0) {
      visible.push(r);
      positionAt.set(r, y);
      sizeAt.set(r, h);
    }
    y += h;
  }

  return { visible, positionAt, sizeAt, frozenTotal };
}

/** Constant-time cellRect using precomputed AxisLayouts. Caller guarantees
 *  the (row, col) pair is in `cols.visible` × `rows.visible`; otherwise the
 *  rect is anchored at the data origin with the cell's nominal size. */
export function cellRectIn(
  layout: LayoutSlice,
  cols: AxisLayout,
  rows: AxisLayout,
  row: number,
  col: number,
): Rect {
  return {
    x: gridOriginX(layout) + (cols.positionAt.get(col) ?? 0),
    y: gridOriginY(layout) + (rows.positionAt.get(row) ?? 0),
    w: cols.sizeAt.get(col) ?? colWidth(layout, col),
    h: rows.sizeAt.get(row) ?? rowHeight(layout, row),
  };
}

/** Up to four rectangles covering the visible portion of `range`, one per
 *  freeze quadrant the range overlaps. Returns an empty list if the range
 *  is entirely outside the visible area. */
export function rangeRects(
  layout: LayoutSlice,
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
      rects.push({ x: tl.x, y: tl.y, w: br.x + br.w - tl.x, h: br.y + br.h - tl.y });
    }
  }
  return rects;
}
