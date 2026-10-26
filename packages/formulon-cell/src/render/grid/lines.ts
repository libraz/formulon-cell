// Gridline + freeze-divider painters. Both run after cell painting so they
// sit above the cell rectangles but below the active selection outline.

import type { ResolvedTheme } from '../../theme/resolve.js';
import {
  type AxisLayout,
  cellRectIn,
  gridOriginX,
  gridOriginY,
  mirrorX,
  type ViewState,
} from '../geometry.js';
import { type ChromePaintContext, ruleWidth } from './chrome-context.js';

export function paintGridLines(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { ctx, dpr, cssWidth, cssHeight } = pc;
  const { layout } = state;
  const width = ruleWidth(dpr);
  const align = width / 2;

  ctx.strokeStyle = theme.rule;
  ctx.lineWidth = width;
  ctx.beginPath();

  const firstRow = rows.visible[0] ?? 0;
  const firstCol = cols.visible[0] ?? 0;

  for (const c of cols.visible) {
    const rect = cellRectIn(layout, cols, rows, firstRow, c);
    const xx = Math.round(rect.x) + align;
    ctx.moveTo(xx, 0);
    ctx.lineTo(xx, cssHeight);
  }
  const lastCol = cols.visible[cols.visible.length - 1];
  if (lastCol !== undefined) {
    const rect = cellRectIn(layout, cols, rows, firstRow, lastCol);
    const xx = Math.round(rect.x + rect.w) + align;
    ctx.moveTo(xx, 0);
    ctx.lineTo(xx, cssHeight);
  }

  for (const r of rows.visible) {
    const rect = cellRectIn(layout, cols, rows, r, firstCol);
    const yy = Math.round(rect.y) + align;
    ctx.moveTo(0, yy);
    ctx.lineTo(cssWidth, yy);
  }
  const lastRow = rows.visible[rows.visible.length - 1];
  if (lastRow !== undefined) {
    const rect = cellRectIn(layout, cols, rows, lastRow, firstCol);
    const yy = Math.round(rect.y + rect.h) + align;
    ctx.moveTo(0, yy);
    ctx.lineTo(cssWidth, yy);
  }

  ctx.stroke();
}

export function paintFreezeDividers(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { layout } = state;
  if (layout.freezeRows === 0 && layout.freezeCols === 0) return;
  if (layout.splitPanes === true) {
    paintSplitBars(pc, state, theme, cols, rows);
    return;
  }
  const { ctx, dpr, cssWidth, cssHeight } = pc;
  const width = ruleWidth(dpr);
  ctx.strokeStyle = theme.freezeRule;
  ctx.lineWidth = width;
  const align = width / 2;
  ctx.beginPath();
  if (layout.freezeRows > 0) {
    const yy = Math.round(gridOriginY(layout) + rows.frozenTotal) + align;
    ctx.moveTo(0, yy);
    ctx.lineTo(cssWidth, yy);
  }
  if (layout.freezeCols > 0) {
    const xx = Math.round(mirrorX(layout, gridOriginX(layout) + cols.frozenTotal)) + align;
    ctx.moveTo(xx, 0);
    ctx.lineTo(xx, cssHeight);
  }
  ctx.stroke();
}

/** Thickness of a split bar in CSS px; also the grab tolerance for dragging it. */
export const SPLIT_BAR_PX = 5;

/** Screen position of the split bars' centre lines; null on an axis without one. */
export function splitBarPositions(
  state: ViewState,
  cols: Pick<AxisLayout, 'frozenTotal'>,
  rows: Pick<AxisLayout, 'frozenTotal'>,
): { x: number | null; y: number | null } {
  const { layout } = state;
  return {
    x: layout.freezeCols > 0 ? mirrorX(layout, gridOriginX(layout) + cols.frozenTotal) : null,
    y: layout.freezeRows > 0 ? gridOriginY(layout) + rows.frozenTotal : null,
  };
}

function paintSplitBars(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { ctx, cssWidth, cssHeight } = pc;
  const { x, y } = splitBarPositions(state, cols, rows);
  const half = SPLIT_BAR_PX / 2;
  ctx.save();
  ctx.fillStyle = theme.headerBg;
  ctx.strokeStyle = theme.freezeRule;
  ctx.lineWidth = 1;
  if (y !== null) {
    const top = Math.round(y - half);
    ctx.fillRect(0, top, cssWidth, SPLIT_BAR_PX);
    ctx.strokeRect(-1, top + 0.5, cssWidth + 2, SPLIT_BAR_PX - 1);
  }
  if (x !== null) {
    const left = Math.round(x - half);
    ctx.fillRect(left, 0, SPLIT_BAR_PX, cssHeight);
    ctx.strokeRect(left + 0.5, -1, SPLIT_BAR_PX - 1, cssHeight + 2);
  }
  ctx.restore();
}
