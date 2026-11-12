import { headerFilterRanges } from '../commands/filter-ranges.js';
import { isMacHost } from '../extensions/ui-options.js';
import { hitZone, layoutForView, type Rect, rangeRects } from '../render/geometry.js';
import {
  PAGE_BREAK_GRAB,
  type PageBandHit,
  type PageBreakHandle,
  RULER_GRAB,
  type RulerHandle,
} from '../render/grid/page-view.js';
import { type GridHitState, gridHitState } from '../render/grid.js';
import { isWholeColumnRange, isWholeRowRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { MAC_GRID_CURSORS, resizeCursor } from './grid-cursors.js';
import { resizableTableGripAt } from './table-resize.js';

/** Pixels the fill-handle hit zone extends past the painted handle. */
const FILL_HANDLE_HIT_PAD = 3;

/** The page boundary under the pointer, if Page Break Preview is showing one
 *  there. A manual break wins over an automatic one at the same position so
 *  a stack of coincident lines still drags the one the user can move. */
export const pageBreakHandleAt = (
  s: State,
  hits: GridHitState,
  x: number,
  y: number,
): PageBreakHandle | null => {
  if (s.ui.workbookView !== 'pageBreakPreview') return null;
  let best: PageBreakHandle | null = null;
  for (const handle of hits.pageBreakHandles) {
    const distance =
      handle.axis === 'row' ? Math.abs(y - handle.position) : Math.abs(x - handle.position);
    if (distance > PAGE_BREAK_GRAB) continue;
    if (!best || (handle.kind === 'break' && best.kind === 'printArea')) best = handle;
  }
  return best;
};

/** The ruler margin boundary under the pointer in Page Layout view. Only
 *  the ruler bands themselves are live; the same coordinate further into the
 *  sheet belongs to a cell. */
export const rulerHandleAt = (
  s: State,
  hits: GridHitState,
  x: number,
  y: number,
): RulerHandle | null => {
  if (s.ui.workbookView !== 'pageLayout') return null;
  for (const handle of hits.rulerHandles) {
    const horizontal = handle.side === 'left' || handle.side === 'right';
    const across = horizontal ? y : x;
    if (across < handle.bandStart || across > handle.bandEnd) continue;
    const distance = horizontal ? Math.abs(x - handle.position) : Math.abs(y - handle.position);
    if (distance <= RULER_GRAB) return handle;
  }
  return null;
};

/** Margin width, in inches, implied by dropping `handle` at (x, y). The
 *  ruler's paper origin is the page edge, so the distance from there to the
 *  pointer is the margin — mirrored on a right-to-left sheet, where the
 *  leading edge is physically on the right. */
export const marginInchesFor = (handle: RulerHandle, x: number, y: number): number => {
  const horizontal = handle.side === 'left' || handle.side === 'right';
  const at = horizontal ? x : y;
  const raw = Math.abs(at - handle.paperOrigin) / Math.max(1, handle.pxPerInch);
  return Math.max(0, Math.round(raw * 100) / 100);
};

/** The header / footer slot under the pointer in Page Layout view. */
export const pageBandAt = (
  s: State,
  hits: GridHitState,
  x: number,
  y: number,
): PageBandHit | null => {
  if (s.ui.workbookView !== 'pageLayout') return null;
  for (const band of hits.pageBands) {
    if (x < band.rect.x || x > band.rect.x + band.rect.w) continue;
    if (y < band.rect.y || y > band.rect.y + band.rect.h) continue;
    return band;
  }
  return null;
};

/** Row / column the pointer is over, ignoring which zone it lands in — a
 *  break dragged across the header rail still has a target. */
export const indexAt = (s: State, x: number, y: number): { row: number; col: number } | null => {
  const layout = layoutForView(s);
  const zone = hitZone(layout, s.viewport, x, y, [], { resizeHandles: false });
  if (!zone) return null;
  if (zone.kind === 'cell') return { row: zone.row, col: zone.col };
  if (zone.kind === 'row-header') return { row: zone.row, col: s.selection.active.col };
  if (zone.kind === 'col-header') return { row: s.selection.active.row, col: zone.col };
  return null;
};

/** Pixels either side of the selection frame that grab it for a drag-move. */
const SELECTION_BORDER_GRAB = 3;

/** Whether (x, y) sits on the frame of a single, bounded selection, where a
 *  drag moves the block. Disjoint, whole-row/column and object selections and
 *  an open editor have no movable frame. */
export const isSelectionBorderHit = (s: State, x: number, y: number): boolean => {
  const range = s.selection.range;
  if ((s.selection.extraRanges?.length ?? 0) > 0) return false;
  if (s.ui.selectedObject || s.ui.editor.kind !== 'idle') return false;
  if (isWholeRowRange(range) || isWholeColumnRange(range)) return false;
  const grab = SELECTION_BORDER_GRAB;
  return rangeRects(layoutForView(s), s.viewport, range).some((rect) => {
    const inX = x >= rect.x - grab && x <= rect.x + rect.w + grab;
    const inY = y >= rect.y - grab && y <= rect.y + rect.h + grab;
    if (!inX || !inY) return false;
    const nearVertical = Math.abs(x - rect.x) <= grab || Math.abs(x - (rect.x + rect.w)) <= grab;
    const nearHorizontal = Math.abs(y - rect.y) <= grab || Math.abs(y - (rect.y + rect.h)) <= grab;
    return nearVertical || nearHorizontal;
  });
};

export const isFillHandleHit = (hits: GridHitState, x: number, y: number): boolean => {
  const rect = hits.fillHandle;
  if (!rect) return false;
  // Pad so the handle is comfortable to grab.
  const pad = FILL_HANDLE_HIT_PAD;
  return (
    x >= rect.x - pad &&
    x <= rect.x + rect.w + pad &&
    y >= rect.y - pad &&
    y <= rect.y + rect.h + pad
  );
};

export function updateCursor(
  host: HTMLElement,
  store: SpreadsheetStore,
  x: number,
  y: number,
): void {
  const s0 = store.getState();
  const hits = gridHitState(host);
  const breakHandle = pageBreakHandleAt(s0, hits, x, y);
  if (breakHandle) {
    host.style.cursor = breakHandle.axis === 'row' ? 'row-resize' : 'col-resize';
    return;
  }
  const ruler = rulerHandleAt(s0, hits, x, y);
  if (ruler) {
    host.style.cursor =
      ruler.side === 'left' || ruler.side === 'right' ? 'col-resize' : 'row-resize';
    return;
  }
  if (pageBandAt(s0, hits, x, y)) {
    host.style.cursor = 'text';
    return;
  }
  const refBox = refBoxHitAt(s0, x, y);
  if (refBox) {
    host.style.cursor =
      refBox.part === 'edge'
        ? 'default'
        : refBox.corner === 'nw' || refBox.corner === 'se'
          ? 'nwse-resize'
          : 'nesw-resize';
    return;
  }
  if (isFillHandleHit(hits, x, y)) {
    host.style.cursor = 'crosshair';
    return;
  }
  if (resizableTableGripAt(store, x, y)) {
    host.style.cursor = 'nwse-resize';
    return;
  }
  if (isSelectionBorderHit(s0, x, y)) {
    host.style.cursor = 'move';
    return;
  }
  const s = store.getState();
  const zone = hitZone(layoutForView(s), s.viewport, x, y, headerFilterRanges(s));
  if (!zone) {
    host.style.cursor = '';
    return;
  }
  const mac = isMacHost(host);
  if (zone.kind === 'col-resize') host.style.cursor = resizeCursor(host, 'col');
  else if (zone.kind === 'row-resize') host.style.cursor = resizeCursor(host, 'row');
  else if (zone.kind === 'col-filter-btn') host.style.cursor = 'pointer';
  else if (mac && zone.kind === 'col-header') host.style.cursor = MAC_GRID_CURSORS.colHeader;
  else if (mac && zone.kind === 'row-header') host.style.cursor = MAC_GRID_CURSORS.rowHeader;
  else if (mac && zone.kind === 'cell') host.style.cursor = MAC_GRID_CURSORS.cell;
  else host.style.cursor = '';
}

/** Pixels around a reference box corner that grab its resize handle. */
const REF_BOX_CORNER_GRAB = 5;
/** Pixels either side of a reference box edge that grab it for a move. */
const REF_BOX_EDGE_GRAB = 3;

export type RefBoxCorner = 'nw' | 'ne' | 'sw' | 'se';

export interface RefBoxHit {
  /** Index into `ui.editorRefs`. */
  index: number;
  part: 'edge' | 'corner';
  corner?: RefBoxCorner;
}

const refBoxRect = (s: State, index: number): Rect | null => {
  const ref = s.ui.editorRefs[index];
  if (!ref) return null;
  const rects = rangeRects(layoutForView(s), s.viewport, {
    r0: ref.r0,
    c0: ref.c0,
    r1: ref.r1,
    c1: ref.c1,
  });
  if (rects.length === 0) return null;
  const x0 = Math.min(...rects.map((r) => r.x));
  const y0 = Math.min(...rects.map((r) => r.y));
  const x1 = Math.max(...rects.map((r) => r.x + r.w));
  const y1 = Math.max(...rects.map((r) => r.y + r.h));
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
};

/** The formula-edit reference box part under (x, y): a corner handle first,
 *  else an edge. The cells inside a box belong to the grid as usual. */
export const refBoxHitAt = (s: State, x: number, y: number): RefBoxHit | null => {
  if (s.ui.editor.kind === 'idle' || s.ui.editorRefs.length === 0) return null;
  const rects = s.ui.editorRefs.map((_, i) => refBoxRect(s, i));
  for (const [index, r] of rects.entries()) {
    if (!r) continue;
    const corners: [RefBoxCorner, number, number][] = [
      ['nw', r.x, r.y],
      ['ne', r.x + r.w, r.y],
      ['sw', r.x, r.y + r.h],
      ['se', r.x + r.w, r.y + r.h],
    ];
    for (const [corner, cx, cy] of corners) {
      if (Math.abs(x - cx) <= REF_BOX_CORNER_GRAB && Math.abs(y - cy) <= REF_BOX_CORNER_GRAB) {
        return { index, part: 'corner', corner };
      }
    }
  }
  const grab = REF_BOX_EDGE_GRAB;
  for (const [index, r] of rects.entries()) {
    if (!r) continue;
    if (x < r.x - grab || x > r.x + r.w + grab || y < r.y - grab || y > r.y + r.h + grab) continue;
    const nearVertical = Math.abs(x - r.x) <= grab || Math.abs(x - (r.x + r.w)) <= grab;
    const nearHorizontal = Math.abs(y - r.y) <= grab || Math.abs(y - (r.y + r.h)) <= grab;
    if (nearVertical || nearHorizontal) return { index, part: 'edge' };
  }
  return null;
};
