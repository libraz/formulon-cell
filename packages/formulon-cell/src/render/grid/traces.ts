// Trace-arrow overlay. Iterates the active traces in state and paints one dot
// + arrow per same-sheet entry. Off-screen endpoints and cross-sheet entries
// are silently skipped; partial visibility is not handled in v1 (matches the
// spreadsheet convention of clipping against the freeze divider).

import type { State } from '../../store/store.js';
import { type AxisLayout, cellRectIn, isColVisible, isRowVisible } from '../geometry.js';
import {
  paintTraceArrow,
  paintTraceDot,
  TRACE_DEPENDENT_COLOR,
  TRACE_PRECEDENT_COLOR,
} from '../painters/trace.js';

export function paintTraces(
  ctx: CanvasRenderingContext2D,
  state: State,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const items = state.traces.items;
  if (items.length === 0) return;
  const sheet = state.data.sheetIndex;
  const { layout, viewport } = state;
  for (const item of items) {
    if (item.from.sheet !== sheet || item.to.sheet !== sheet) continue;
    if (!isRowVisible(layout, viewport, item.from.row)) continue;
    if (!isColVisible(layout, viewport, item.from.col)) continue;
    if (!isRowVisible(layout, viewport, item.to.row)) continue;
    if (!isColVisible(layout, viewport, item.to.col)) continue;
    const fromRect = cellRectIn(layout, cols, rows, item.from.row, item.from.col);
    const toRect = cellRectIn(layout, cols, rows, item.to.row, item.to.col);
    const color = item.kind === 'precedent' ? TRACE_PRECEDENT_COLOR : TRACE_DEPENDENT_COLOR;
    paintTraceDot(ctx, fromRect, color);
    paintTraceArrow(ctx, fromRect, toRect, color);
  }
}
