// View > Focus Cell: the active cell's row and column are banded across the sheet.

import { MAX_COL, MAX_ROW } from '../../engine/address.js';
import { type Rect, rangeRects, type ViewState } from '../geometry.js';
import type { ChromePaintContext } from './chrome-context.js';

/** The band, measured on Excel for Mac: #D3EED8 over a white cell and gridlines
 *  showing through, which a 20% green wash over both reproduces. */
export const FOCUS_CELL_BAND = 'rgba(35, 170, 60, 0.2)';

/** Whether the band is drawn: Focus Cell is on, or Show Auto-Highlight is on and the selection is
 *  still the cell a Find hit selected. Leaving that cell or closing the Find box ends it. */
export function focusBandShown(state: Pick<ViewState, 'ui' | 'selection'>): boolean {
  if (state.ui.focusCell) return true;
  const hit = state.ui.findHighlight;
  if (!state.ui.autoHighlight || !hit) return false;
  const { active, range } = state.selection;
  return (
    active.sheet === hit.sheet &&
    active.row === hit.row &&
    active.col === hit.col &&
    range.r0 === hit.row &&
    range.r1 === hit.row &&
    range.c0 === hit.col &&
    range.c1 === hit.col &&
    (state.selection.extraRanges ?? []).length === 0
  );
}

/** Where the band is drawn: the active cell's whole row and column, and the
 *  selection it leaves out (the active cell is always inside it). */
export function focusCellBand(state: Pick<ViewState, 'layout' | 'viewport' | 'selection'>): {
  band: Rect[];
  selection: Rect[];
} {
  const { layout, viewport, selection } = state;
  const { row, col } = selection.active;
  const band = [
    ...rangeRects(layout, viewport, { r0: row, r1: row, c0: 0, c1: MAX_COL }),
    ...rangeRects(layout, viewport, { r0: 0, r1: MAX_ROW, c0: col, c1: col }),
  ];
  const ranges = [selection.range, ...(selection.extraRanges ?? [])];
  return { band, selection: ranges.flatMap((r) => rangeRects(layout, viewport, r)) };
}

/** Paint the band under the cells: a filled cell covers it with its own fill. */
export function paintFocusCellBand(
  pc: ChromePaintContext,
  state: Pick<ViewState, 'layout' | 'viewport' | 'selection'>,
): void {
  const { ctx, cssWidth, cssHeight } = pc;
  const { band, selection } = focusCellBand(state);
  if (band.length === 0) return;
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, 0, cssWidth, cssHeight);
  for (const r of selection) ctx.rect(r.x, r.y, r.w, r.h);
  ctx.clip('evenodd');
  ctx.fillStyle = FOCUS_CELL_BAND;
  // The row and the column cross only at the active cell, which the clip leaves out.
  for (const r of band) ctx.fillRect(r.x, r.y, r.w, r.h);
  ctx.restore();
}
