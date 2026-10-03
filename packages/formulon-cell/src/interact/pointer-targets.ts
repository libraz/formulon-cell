import { hitZone, layoutForView } from '../render/geometry.js';
import {
  getPageBandHits,
  getPageBreakHandles,
  getRulerHandles,
  PAGE_BREAK_GRAB,
  type PageBandHit,
  type PageBreakHandle,
  RULER_GRAB,
  type RulerHandle,
} from '../render/grid/page-view.js';
import { getFillHandleRect } from '../render/grid.js';
import type { SpreadsheetStore, State } from '../store/store.js';

/** Pixels the fill-handle hit zone extends past the painted handle. */
const FILL_HANDLE_HIT_PAD = 3;

/** The page boundary under the pointer, if Page Break Preview is showing one
 *  there. A manual break wins over an automatic one at the same position so
 *  a stack of coincident lines still drags the one the user can move. */
export const pageBreakHandleAt = (s: State, x: number, y: number): PageBreakHandle | null => {
  if (s.ui.workbookView !== 'pageBreakPreview') return null;
  let best: PageBreakHandle | null = null;
  for (const handle of getPageBreakHandles()) {
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
export const rulerHandleAt = (s: State, x: number, y: number): RulerHandle | null => {
  if (s.ui.workbookView !== 'pageLayout') return null;
  for (const handle of getRulerHandles()) {
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
export const pageBandAt = (s: State, x: number, y: number): PageBandHit | null => {
  if (s.ui.workbookView !== 'pageLayout') return null;
  for (const band of getPageBandHits()) {
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
  const zone = hitZone(layout, s.viewport, x, y, null, { resizeHandles: false });
  if (!zone) return null;
  if (zone.kind === 'cell') return { row: zone.row, col: zone.col };
  if (zone.kind === 'row-header') return { row: zone.row, col: s.selection.active.col };
  if (zone.kind === 'col-header') return { row: s.selection.active.row, col: zone.col };
  return null;
};

export const isFillHandleHit = (x: number, y: number): boolean => {
  const rect = getFillHandleRect();
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
  const breakHandle = pageBreakHandleAt(s0, x, y);
  if (breakHandle) {
    host.style.cursor = breakHandle.axis === 'row' ? 'row-resize' : 'col-resize';
    return;
  }
  const ruler = rulerHandleAt(s0, x, y);
  if (ruler) {
    host.style.cursor =
      ruler.side === 'left' || ruler.side === 'right' ? 'col-resize' : 'row-resize';
    return;
  }
  if (pageBandAt(s0, x, y)) {
    host.style.cursor = 'text';
    return;
  }
  if (isFillHandleHit(x, y)) {
    host.style.cursor = 'crosshair';
    return;
  }
  const s = store.getState();
  const zone = hitZone(layoutForView(s), s.viewport, x, y, s.ui.filterRange);
  if (!zone) {
    host.style.cursor = '';
    return;
  }
  if (zone.kind === 'col-resize') host.style.cursor = 'col-resize';
  else if (zone.kind === 'row-resize') host.style.cursor = 'row-resize';
  else if (zone.kind === 'col-filter-btn') host.style.cursor = 'pointer';
  else host.style.cursor = '';
}
