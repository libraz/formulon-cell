// Row + column header bar, autofilter chevron, and outline gutters. Drawn in
// a single pass after cell painting so headers sit above the gridlines.

import { colLetter } from '../../engine/address.js';
import { sameRange } from '../../store/selection-geometry.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import {
  type AxisLayout,
  cellRectIn,
  gridOriginX,
  gridOriginY,
  rulerLeft,
  rulerTop,
  type ViewState,
} from '../geometry.js';
import { paintOutlineGutters } from '../painters/controls.js';
import type { OutlineToggleHit } from '../painters.js';
import { type ChromePaintContext, ruleWidth } from './chrome-context.js';

export function paintHeaders(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
): OutlineToggleHit[] {
  const { ctx, dpr, cssWidth, cssHeight } = pc;
  const { layout, selection } = state;
  const active = selection.active;

  const ox = gridOriginX(layout);
  const oy = gridOriginY(layout);
  const rtl = layout.rtl;
  // The ruler bands, when Page Layout is showing them, sit outboard of the
  // outline gutters, so every rail offset starts past them.
  const labelTopY = rulerTop(layout) + layout.outlineColGutter;
  // The whole row rail — outline gutter plus number strip — swaps to the right
  // edge on a right-to-left sheet. `railLeftX` is where that rail starts,
  // `labelLeftX` where the number strip inside it starts, and `innerX` is the
  // rail's boundary with the data area, which is the edge every row-header
  // affordance hugs.
  const railLeftX = rtl ? cssWidth - ox : rulerLeft(layout);
  const labelLeftX = rtl ? cssWidth - ox : rulerLeft(layout) + layout.outlineRowGutter;
  const innerX = rtl ? cssWidth - ox : ox;

  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(railLeftX, 0, ox, oy);
  ctx.save();
  ctx.fillStyle = theme.headerFg;
  ctx.globalAlpha = 0.34;
  const wedgeOuter = rtl ? railLeftX + ox - 12 : labelLeftX + 12;
  const wedgeInner = rtl ? innerX + 7 : ox - 7;
  ctx.beginPath();
  ctx.moveTo(wedgeOuter, labelTopY + layout.headerRowHeight - 6);
  ctx.lineTo(wedgeInner, labelTopY + 8);
  ctx.lineTo(wedgeInner, labelTopY + layout.headerRowHeight - 6);
  ctx.closePath();
  ctx.fill();
  ctx.restore();

  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(rtl ? 0 : ox, 0, cssWidth - ox, oy);
  ctx.fillRect(railLeftX, oy, ox, cssHeight - oy);

  ctx.strokeStyle = theme.ruleStrong;
  const width = ruleWidth(dpr);
  ctx.lineWidth = width;
  const align = width / 2;
  ctx.beginPath();
  ctx.moveTo(0, Math.round(oy) + align);
  ctx.lineTo(cssWidth, Math.round(oy) + align);
  ctx.moveTo(Math.round(innerX) + align, 0);
  ctx.lineTo(Math.round(innerX) + align, cssHeight);
  ctx.stroke();

  const firstRow = rows.visible[0] ?? 0;
  const firstCol = cols.visible[0] ?? 0;

  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  const r1c1 = state.ui.r1c1 === true;
  const filterRanges = state.ui.filterRanges.filter((r) => r.sheet === state.data.sheetIndex);
  const filterSorts = (state.ui.filterSorts ?? []).filter(
    (spec) =>
      spec.range.sheet === state.data.sheetIndex &&
      filterRanges.some((range) => sameRange(range, spec.range)),
  );
  // A selected chart or picture takes the highlight off the headers, as in Excel.
  const objectSelected = state.ui.selectedObject !== null;
  const selectedRanges = objectSelected ? [] : [selection.range, ...(selection.extraRanges ?? [])];
  const colSelected = (col: number): boolean =>
    selectedRanges.some((r) => col >= r.c0 && col <= r.c1);
  const rowSelected = (row: number): boolean =>
    selectedRanges.some((r) => row >= r.r0 && row <= r.r1);
  for (const c of cols.visible) {
    const rect = cellRectIn(layout, cols, rows, firstRow, c);
    const w = cols.sizeAt.get(c) ?? 0;
    const isActiveCol = !objectSelected && c === active.col;
    const isSelectedCol = colSelected(c);
    if (isActiveCol) {
      ctx.fillStyle = theme.bgHeader;
      ctx.fillRect(rect.x, labelTopY, w, layout.headerRowHeight);
    } else if (isSelectedCol) {
      ctx.fillStyle = theme.bgHeader;
      ctx.fillRect(rect.x, labelTopY, w, layout.headerRowHeight);
    }
    ctx.strokeStyle = theme.headerRule;
    ctx.lineWidth = width;
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
    const label = r1c1 ? `C${c + 1}` : colLetter(c);
    ctx.fillText(label, rect.x + w / 2, labelTopY + layout.headerRowHeight / 2 + 0.5);
    // The accent bar covers the rule under the header and one pixel above it.
    if (isActiveCol || isSelectedCol) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(rect.x, Math.round(oy) - 1, w, 2);
    }

    // Autofilter chevron — small ▼ in the right edge of the header for any
    // column inside a filter range.
    if (w >= 28 && filterRanges.some((fr) => c >= fr.c0 && c <= fr.c1)) {
      // The chevron hugs the column's trailing edge, which the mirror puts on
      // the physical left of the header cell for a right-to-left sheet.
      const btnLeft = rtl ? rect.x + 4 : rect.x + w - 4 - 14;
      const btnRight = btnLeft + 14;
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
      ctx.fillStyle = filterActive ? theme.bg : theme.headerFg;
      ctx.beginPath();
      const tx = bx + bw / 2;
      const ty = cy + 1;
      const sorted = filterSorts
        .flatMap((spec) => spec.keys)
        .find((key) => key.byCol === c)?.direction;
      if (sorted) {
        // Sorted column: a stemmed arrow, pointing up for ascending.
        const up = sorted === 'asc';
        const tip = up ? ty - 4 : ty + 4;
        const tail = up ? ty + 4 : ty - 4;
        const wing = up ? ty - 0.5 : ty + 0.5;
        ctx.strokeStyle = ctx.fillStyle;
        ctx.lineWidth = 1.25;
        ctx.moveTo(tx, tail);
        ctx.lineTo(tx, tip);
        ctx.moveTo(tx - 3, wing);
        ctx.lineTo(tx, tip);
        ctx.lineTo(tx + 3, wing);
        ctx.stroke();
      } else {
        // Chevron triangle.
        ctx.moveTo(tx - 3.5, ty - 2);
        ctx.lineTo(tx + 3.5, ty - 2);
        ctx.lineTo(tx, ty + 2);
        ctx.closePath();
        ctx.fill();
      }
      ctx.restore();
    }
  }

  ctx.textAlign = rtl ? 'left' : 'right';
  for (const r of rows.visible) {
    const rect = cellRectIn(layout, cols, rows, r, firstCol);
    const h = rows.sizeAt.get(r) ?? 0;
    const isActiveRow = !objectSelected && r === active.row;
    const isSelectedRow = rowSelected(r);
    if (isActiveRow) {
      ctx.fillStyle = theme.bgHeader;
      ctx.fillRect(labelLeftX, rect.y, layout.headerColWidth, h);
    } else if (isSelectedRow) {
      ctx.fillStyle = theme.bgHeader;
      ctx.fillRect(labelLeftX, rect.y, layout.headerColWidth, h);
    }
    ctx.strokeStyle = theme.headerRule;
    ctx.lineWidth = width;
    ctx.beginPath();
    ctx.moveTo(labelLeftX, Math.round(rect.y + h) + align);
    ctx.lineTo(labelLeftX + layout.headerColWidth, Math.round(rect.y + h) + align);
    ctx.stroke();
    ctx.fillStyle = isActiveRow
      ? theme.headerFgActive
      : isSelectedRow
        ? theme.headerFgActive
        : theme.headerFg;
    ctx.font = `${isActiveRow || isSelectedRow ? 600 : 400} ${theme.textHeader}px ${theme.fontUi}`;
    const rowLabel = r1c1 ? `R${r + 1}` : String(r + 1);
    ctx.fillText(rowLabel, rtl ? innerX + 8 : innerX - 8, rect.y + h / 2 + 0.5);
    if (isActiveRow || isSelectedRow) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(Math.round(innerX) - 1, rect.y, 2, h);
    }
  }

  // Bracket gutters for outline groups.
  return paintOutlineGutters(ctx, state, theme, cols, rows, cssWidth, cssHeight);
}
