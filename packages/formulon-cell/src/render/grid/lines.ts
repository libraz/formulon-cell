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
import { type ChromePaintContext, ruleCenter, ruleWidth } from './chrome-context.js';

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
  const rtl = layout.rtl === true;

  ctx.strokeStyle = theme.rule;
  ctx.lineWidth = width;
  ctx.beginPath();

  const firstRow = rows.visible[0] ?? 0;
  const firstCol = cols.visible[0] ?? 0;

  // Each cell owns the rule on its trailing edges; the leading edge of the
  // first visible column / row carries the rule of the one before it.
  const xs = new Set<number>();
  for (const c of cols.visible) {
    const rect = cellRectIn(layout, cols, rows, firstRow, c);
    xs.add(rect.x);
    xs.add(rect.x + rect.w);
  }
  for (const x of xs) {
    const xx = ruleCenter(x, dpr, rtl);
    ctx.moveTo(xx, 0);
    ctx.lineTo(xx, cssHeight);
  }
  const ys = new Set<number>();
  for (const r of rows.visible) {
    const rect = cellRectIn(layout, cols, rows, r, firstCol);
    ys.add(rect.y);
    ys.add(rect.y + rect.h);
  }
  for (const y of ys) {
    const yy = ruleCenter(y, dpr);
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
  ctx.beginPath();
  if (layout.freezeRows > 0) {
    const yy = ruleCenter(gridOriginY(layout) + rows.frozenTotal, dpr);
    ctx.moveTo(0, yy);
    ctx.lineTo(cssWidth, yy);
  }
  if (layout.freezeCols > 0) {
    const xx = ruleCenter(
      mirrorX(layout, gridOriginX(layout) + cols.frozenTotal),
      dpr,
      layout.rtl === true,
    );
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
