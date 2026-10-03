import { addrKey } from '../../engine/address.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import { type AxisLayout, cellRectIn, rangeRects, type ViewState } from '../geometry.js';
import { paintCellBorders } from '../painters.js';
import { mergedDiagonalBorders, mergeRangeAt } from './merged-cells.js';
import { connectedRects } from './range-rects.js';

export function paintBorders(
  ctx: CanvasRenderingContext2D,
  state: ViewState,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { layout, data, format } = state;
  if (format.formats.size === 0) return;
  for (const r of rows.visible) {
    for (const c of cols.visible) {
      const key = addrKey({ sheet: data.sheetIndex, row: r, col: c });
      const f = format.formats.get(key);
      if (!f?.borders) continue;
      const merge = mergeRangeAt(state, { sheet: data.sheetIndex, row: r, col: c });
      let borders = f.borders;
      if (merge) {
        // A loaded workbook can retain per-cell border sides even after a
        // merge. Keep only the perimeter sides; otherwise every interior
        // gridline is redrawn above the merge surface.
        borders = { ...f.borders };
        if (r > merge.r0) borders.top = undefined;
        if (r < merge.r1) borders.bottom = undefined;
        if (layout.rtl ? c < merge.c1 : c > merge.c0) borders.left = undefined;
        if (layout.rtl ? c > merge.c0 : c < merge.c1) borders.right = undefined;
        // Diagonals belong to the merged surface, not to each component
        // cell. They are painted once below after the perimeter pass.
        borders.diagonalDown = undefined;
        borders.diagonalUp = undefined;
        if (Object.values(borders).every((side) => !side)) continue;
      }
      const bounds = cellRectIn(layout, cols, rows, r, c);
      paintCellBorders({
        ctx,
        theme,
        bounds,
        value: { kind: 'blank' },
        formula: null,
        isActive: false,
        isInRange: false,
        format: borders === f.borders ? f : { ...f, borders },
      });
    }
  }

  // A uniform diagonal border survives a merge in Excel, but is rendered as
  // one line over the visible merged rectangle. A per-cell pass would draw
  // one diagonal per body cell and leave crossings in the merged surface;
  // rangeRects keeps the same freeze/RTL projection as the selection path.
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== data.sheetIndex) continue;
    const borders = mergedDiagonalBorders(format.formats, merge);
    if (!borders) continue;
    const rects = rangeRects(layout, state.viewport, merge);
    const clips = connectedRects(rects);
    if (clips.length === 0) continue;
    const globalBounds = {
      x: Math.min(...rects.map((rect) => rect.x)),
      y: Math.min(...rects.map((rect) => rect.y)),
      w:
        Math.max(...rects.map((rect) => rect.x + rect.w)) -
        Math.min(...rects.map((rect) => rect.x)),
      h:
        Math.max(...rects.map((rect) => rect.y + rect.h)) -
        Math.min(...rects.map((rect) => rect.y)),
    };
    for (const clip of clips) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(clip.x, clip.y, clip.w, clip.h);
      ctx.clip();
      paintCellBorders({
        ctx,
        theme,
        bounds: globalBounds,
        value: { kind: 'blank' },
        formula: null,
        isActive: false,
        isInRange: false,
        format: { borders },
      });
      ctx.restore();
    }
  }
}
