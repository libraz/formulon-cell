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
