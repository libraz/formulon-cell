// Row + column header bar, autofilter chevron, and outline gutters. Drawn in
// a single pass after cell painting so headers sit above the gridlines.

import type { State } from '../../store/store.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import { type AxisLayout, cellRectIn, colLabel, gridOriginX, gridOriginY } from '../geometry.js';
import { paintOutlineGutters } from '../painters/controls.js';
import type { ChromePaintContext } from './chrome-context.js';
import { setOutlineToggles } from './hit-state.js';

export function paintHeaders(
  pc: ChromePaintContext,
  state: State,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { ctx, dpr, cssWidth, cssHeight } = pc;
  const { layout, selection } = state;
  const active = selection.active;

  const ox = gridOriginX(layout);
  const oy = gridOriginY(layout);
  const labelTopY = layout.outlineColGutter;
  const labelLeftX = layout.outlineRowGutter;

  ctx.fillStyle = theme.bgRail;
  ctx.fillRect(0, 0, ox, oy);
  ctx.save();
  ctx.fillStyle = theme.headerFg;
  ctx.globalAlpha = 0.34;
  ctx.beginPath();
  ctx.moveTo(labelLeftX + 12, labelTopY + layout.headerRowHeight - 6);
  ctx.lineTo(ox - 7, labelTopY + 8);
  ctx.lineTo(ox - 7, labelTopY + layout.headerRowHeight - 6);
  ctx.closePath();
  ctx.fill();
  ctx.restore();

  ctx.fillStyle = theme.bgRail;
  ctx.fillRect(ox, 0, cssWidth - ox, oy);
  ctx.fillRect(0, oy, ox, cssHeight - oy);

  ctx.strokeStyle = theme.ruleStrong;
  ctx.lineWidth = 1 / dpr;
  const align = 0.5 / dpr;
  ctx.beginPath();
  ctx.moveTo(0, Math.round(oy) + align);
  ctx.lineTo(cssWidth, Math.round(oy) + align);
  ctx.moveTo(Math.round(ox) + align, 0);
  ctx.lineTo(Math.round(ox) + align, cssHeight);
  ctx.stroke();

  const firstRow = rows.visible[0] ?? 0;
  const firstCol = cols.visible[0] ?? 0;

  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  const r1c1 = state.ui.r1c1 === true;
  const fr = state.ui.filterRange;
  const selectedRanges = [selection.range, ...(selection.extraRanges ?? [])];
  const colSelected = (col: number): boolean =>
    selectedRanges.some((r) => col >= r.c0 && col <= r.c1);
  const rowSelected = (row: number): boolean =>
    selectedRanges.some((r) => row >= r.r0 && row <= r.r1);
  for (const c of cols.visible) {
    const rect = cellRectIn(layout, cols, rows, firstRow, c);
    const w = cols.sizeAt.get(c) ?? 0;
    const isActiveCol = c === active.col;
    const isSelectedCol = colSelected(c);
    if (isActiveCol) {
      ctx.fillStyle = theme.bgHeader;
      ctx.fillRect(rect.x, labelTopY, w, layout.headerRowHeight);
    } else if (isSelectedCol) {
      ctx.fillStyle = theme.bgHeader;
      ctx.fillRect(rect.x, labelTopY, w, layout.headerRowHeight);
    }
    ctx.strokeStyle = theme.rule;
    ctx.lineWidth = 1 / dpr;
    ctx.beginPath();
    ctx.moveTo(Math.round(rect.x + w) + align, labelTopY);
    ctx.lineTo(Math.round(rect.x + w) + align, oy);
    ctx.stroke();
    ctx.fillStyle = isActiveCol
      ? theme.headerFgActive
      : isSelectedCol
        ? theme.headerFgActive
        : theme.headerFg;
    ctx.font = `${isActiveCol || isSelectedCol ? 600 : 400} ${theme.textHeader}px ${theme.fontUi}`;
    const label = r1c1 ? `C${c + 1}` : colLabel(c);
    ctx.fillText(label, rect.x + w / 2, labelTopY + layout.headerRowHeight / 2 + 0.5);
    if (isActiveCol) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(rect.x, oy - Math.max(2, 2 / dpr), w, Math.max(2, 2 / dpr));
    }

    // Autofilter chevron — small ▼ in the right edge of the header for any
    // column inside the active filter range.
    if (fr && c >= fr.c0 && c <= fr.c1 && w >= 28) {
      const btnRight = rect.x + w - 4;
      const btnLeft = btnRight - 14;
      const cy = labelTopY + layout.headerRowHeight / 2;
      const filterActive = state.layout.hiddenRows.size > 0;
      ctx.save();
      ctx.fillStyle = filterActive ? theme.accent : theme.bgHeader;
      ctx.strokeStyle = filterActive ? theme.accent : theme.ruleStrong;
      ctx.lineWidth = 1 / dpr;
      const radius = 2;
      const bx = btnLeft;
      const by = cy - 7;
      const bw = btnRight - btnLeft;
      const bh = 14;
      // Subtle background pill.
      ctx.beginPath();
      ctx.moveTo(bx + radius, by);
      ctx.lineTo(bx + bw - radius, by);
      ctx.quadraticCurveTo(bx + bw, by, bx + bw, by + radius);
      ctx.lineTo(bx + bw, by + bh - radius);
      ctx.quadraticCurveTo(bx + bw, by + bh, bx + bw - radius, by + bh);
      ctx.lineTo(bx + radius, by + bh);
      ctx.quadraticCurveTo(bx, by + bh, bx, by + bh - radius);
      ctx.lineTo(bx, by + radius);
      ctx.quadraticCurveTo(bx, by, bx + radius, by);
      ctx.closePath();
      if (filterActive) ctx.fill();
      else ctx.stroke();
      // Chevron triangle.
      ctx.fillStyle = filterActive ? theme.bg : theme.headerFg;
      ctx.beginPath();
      const tx = bx + bw / 2;
      const ty = cy + 1;
      ctx.moveTo(tx - 3.5, ty - 2);
      ctx.lineTo(tx + 3.5, ty - 2);
      ctx.lineTo(tx, ty + 2);
      ctx.closePath();
      ctx.fill();
      ctx.restore();
    }
  }

  ctx.textAlign = 'right';
  for (const r of rows.visible) {
    const rect = cellRectIn(layout, cols, rows, r, firstCol);
    const h = rows.sizeAt.get(r) ?? 0;
    const isActiveRow = r === active.row;
    const isSelectedRow = rowSelected(r);
    if (isActiveRow) {
      ctx.fillStyle = theme.bgHeader;
      ctx.fillRect(labelLeftX, rect.y, layout.headerColWidth, h);
    } else if (isSelectedRow) {
      ctx.fillStyle = theme.bgHeader;
      ctx.fillRect(labelLeftX, rect.y, layout.headerColWidth, h);
    }
    ctx.strokeStyle = theme.rule;
    ctx.lineWidth = 1 / dpr;
    ctx.beginPath();
    ctx.moveTo(labelLeftX, Math.round(rect.y + h) + align);
    ctx.lineTo(ox, Math.round(rect.y + h) + align);
    ctx.stroke();
    ctx.fillStyle = isActiveRow
      ? theme.headerFgActive
      : isSelectedRow
        ? theme.headerFgActive
        : theme.headerFg;
    ctx.font = `${isActiveRow || isSelectedRow ? 600 : 400} ${theme.textHeader}px ${theme.fontUi}`;
    const rowLabel = r1c1 ? `R${r + 1}` : String(r + 1);
    ctx.fillText(rowLabel, ox - 8, rect.y + h / 2 + 0.5);
    if (isActiveRow) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(ox - Math.max(2, 2 / dpr), rect.y, Math.max(2, 2 / dpr), h);
    }
  }

  if (cols.positionAt.has(active.col)) {
    const aRect = cellRectIn(layout, cols, rows, firstRow, active.col);
    const w = cols.sizeAt.get(active.col) ?? 0;
    ctx.strokeStyle = theme.accent;
    ctx.lineWidth = 1.5 / dpr;
    ctx.beginPath();
    ctx.moveTo(aRect.x, oy - 0.5);
    ctx.lineTo(aRect.x + w, oy - 0.5);
    ctx.stroke();
  }
  if (rows.positionAt.has(active.row)) {
    const aRect = cellRectIn(layout, cols, rows, active.row, firstCol);
    const h = rows.sizeAt.get(active.row) ?? 0;
    ctx.strokeStyle = theme.accent;
    ctx.lineWidth = 1.5 / dpr;
    ctx.beginPath();
    ctx.moveTo(ox - 0.5, aRect.y);
    ctx.lineTo(ox - 0.5, aRect.y + h);
    ctx.stroke();
  }

  // Bracket gutters for outline groups.
  setOutlineToggles(paintOutlineGutters(ctx, state, theme, cols, rows, cssWidth, cssHeight));
}
