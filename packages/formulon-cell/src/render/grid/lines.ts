// Gridline + freeze-divider painters. Both run after cell painting so they
// sit above the cell rectangles but below the active selection outline.

import type { Range } from '../../engine/types.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import {
  type AxisLayout,
  cellRectIn,
  gridOriginX,
  gridOriginY,
  mirrorX,
  pointPx,
  type Rect,
  rangeRects,
  type ViewState,
} from '../geometry.js';
import { wholePointAdvances } from '../text-metrics.js';
import { type ChromePaintContext, ruleCenter, ruleWidth } from './chrome-context.js';

/** Excel for Mac drops gridlines below this zoom (none at 33%, drawn at 40%). */
const MAC_GRIDLINE_MIN_ZOOM = 0.4;

/** Whether the sheet's gridlines are painted at the current zoom. */
export function gridLinesShown(state: {
  ui: Pick<ViewState['ui'], 'showGridLines'>;
  layout: Pick<ViewState['layout'], 'screenScale'>;
  viewport: Pick<ViewState['viewport'], 'zoom'>;
}): boolean {
  if (state.ui.showGridLines === false) return false;
  return !wholePointAdvances(state.layout) || state.viewport.zoom >= MAC_GRIDLINE_MIN_ZOOM;
}

/** Draw `from..to` minus the `gaps` intervals along one rule. */
const NO_GAPS: readonly [number, number][] = [];

function segments(
  ctx: CanvasRenderingContext2D,
  from: number,
  to: number,
  gaps: readonly [number, number][],
  at: (along: number) => [number, number],
): void {
  let start = from;
  for (const [a, b] of [...gaps].sort((p, q) => p[0] - q[0])) {
    if (a > start) {
      ctx.moveTo(...at(start));
      ctx.lineTo(...at(Math.min(a, to)));
    }
    start = Math.max(start, b);
  }
  if (start < to) {
    ctx.moveTo(...at(start));
    ctx.lineTo(...at(to));
  }
}

/** Paint the sheet's gridlines. Cells of a `bare` range (a PivotTable) drop
 *  the rules they own, except the one on the range's trailing edge. */
export function paintGridLines(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
  bare: readonly Range[] = [],
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
  const bareRects: Rect[] = bare
    .filter((range) => range.sheet === state.data.sheetIndex)
    .flatMap((range) => rangeRects(layout, state.viewport, range));
  for (const x of xs) {
    const xx = ruleCenter(x, dpr, rtl);
    const gaps =
      bareRects.length === 0
        ? NO_GAPS
        : bareRects
            .filter((r) => x > r.x && x < r.x + r.w)
            .map((r): [number, number] => [r.y, r.y + r.h]);
    segments(ctx, 0, cssHeight, gaps, (y) => [xx, y]);
  }
  const ys = new Set<number>();
  for (const r of rows.visible) {
    const rect = cellRectIn(layout, cols, rows, r, firstCol);
    ys.add(rect.y);
    ys.add(rect.y + rect.h);
  }
  for (const y of ys) {
    const yy = ruleCenter(y, dpr);
    const gaps =
      bareRects.length === 0
        ? NO_GAPS
        : bareRects
            .filter((r) => y > r.y && y <= r.y + r.h)
            .map((r): [number, number] => [r.x, r.x + r.w]);
    segments(ctx, 0, cssWidth, gaps, (x) => [x, yy]);
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

const MAC_SPLIT_LINE = 'rgb(212,212,212)';
const MAC_SPLIT_BAND = 'rgb(211,211,211)';

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
  if (wholePointAdvances(state.layout)) {
    paintMacSplitBars(ctx, state, x, y, cssWidth, cssHeight);
    return;
  }
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

/** Mac bar, in points: a 1 pt line (212) then a 4 pt flat band (211). */
function paintMacSplitBars(
  ctx: CanvasRenderingContext2D,
  state: ViewState,
  x: number | null,
  y: number | null,
  cssWidth: number,
  cssHeight: number,
): void {
  const pt = pointPx(state.layout, state.viewport);
  const half = (5 * pt) / 2;
  ctx.save();
  if (y !== null) {
    const top = y - half;
    ctx.fillStyle = MAC_SPLIT_LINE;
    ctx.fillRect(0, top, cssWidth, pt);
    ctx.fillStyle = MAC_SPLIT_BAND;
    ctx.fillRect(0, top + pt, cssWidth, 4 * pt);
  }
  if (x !== null) {
    const left = x - half;
    ctx.fillStyle = MAC_SPLIT_LINE;
    ctx.fillRect(left, 0, pt, cssHeight);
    ctx.fillStyle = MAC_SPLIT_BAND;
    ctx.fillRect(left + pt, 0, 4 * pt, cssHeight);
  }
  ctx.restore();
}
