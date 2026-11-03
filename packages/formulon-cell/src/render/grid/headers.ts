// Row + column header bar and outline gutters. Drawn in
// a single pass after cell painting so headers sit above the gridlines.

import { colLetter } from '../../engine/address.js';
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
import { fontMetricsBox, textBaselineY } from '../painters/text.js';
import type { OutlineToggleHit } from '../painters.js';
import { type ChromePaintContext, ruleCenter, ruleWidth } from './chrome-context.js';

/** Row numbers are bottom-aligned like cell text: the baseline sits the
 *  label font's descent plus one px above the row's bottom edge (6 device px
 *  at DPR 2 for the default font, 8 for Courier New 11, 14 for Courier 20). */
const ROW_LABEL_PAD_Y = 1;

export function paintHeaders(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
  labelFamily: string = theme.fontUi,
  labelSize: number = theme.textHeader,
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
  // Header separators sit on the same trailing rule as the gridlines, so the
  // header band ends with its own rule rather than covering the first row.
  const headerRuleY = ruleCenter(oy, dpr);
  const headerRuleX = ruleCenter(innerX, dpr, rtl);
  ctx.beginPath();
  ctx.moveTo(0, headerRuleY);
  ctx.lineTo(cssWidth, headerRuleY);
  ctx.moveTo(headerRuleX, 0);
  ctx.lineTo(headerRuleX, cssHeight);
  ctx.stroke();

  const firstRow = rows.visible[0] ?? 0;
  const firstCol = cols.visible[0] ?? 0;

  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  const r1c1 = state.ui.r1c1 === true;
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
    const sepX = ruleCenter(rtl ? rect.x : rect.x + w, dpr, rtl);
    ctx.moveTo(sepX, labelTopY);
    ctx.lineTo(sepX, oy);
    ctx.stroke();
    ctx.fillStyle = isActiveCol
      ? theme.headerFgActive
      : isSelectedCol
        ? theme.headerFgActive
        : theme.headerFg;
    ctx.font = `${isActiveCol || isSelectedCol ? 600 : 400} ${labelSize}px ${labelFamily}`;
    const label = r1c1 ? `C${c + 1}` : colLetter(c);
    // Letters centre on the header cell less its trailing separator.
    const labelX = rtl ? rect.x + (w + width) / 2 : rect.x + (w - width) / 2;
    ctx.fillText(label, labelX, labelTopY + layout.headerRowHeight / 2 + 0.5);
    // The accent bar covers the rule under the header and one pixel above it.
    if (isActiveCol || isSelectedCol) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(rect.x, Math.round(oy) - 2, w, 2);
    }
  }

  // Row numbers centre on the strip between the rail edge and its rule.
  ctx.textAlign = 'center';
  ctx.textBaseline = 'alphabetic';
  const rowLabelX = rtl
    ? innerX + (layout.headerColWidth + width) / 2
    : innerX - (layout.headerColWidth + width) / 2;
  ctx.font = `400 ${labelSize}px ${labelFamily}`;
  const rowLabelBox = fontMetricsBox(ctx, labelSize);
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
    const sepY = ruleCenter(rect.y + h, dpr);
    ctx.moveTo(labelLeftX, sepY);
    ctx.lineTo(labelLeftX + layout.headerColWidth, sepY);
    ctx.stroke();
    ctx.fillStyle = isActiveRow
      ? theme.headerFgActive
      : isSelectedRow
        ? theme.headerFgActive
        : theme.headerFg;
    ctx.font = `${isActiveRow || isSelectedRow ? 600 : 400} ${labelSize}px ${labelFamily}`;
    const rowLabel = r1c1 ? `R${r + 1}` : String(r + 1);
    ctx.fillText(
      rowLabel,
      rowLabelX,
      textBaselineY({ ...rect, h }, rowLabelBox, 'bottom', ROW_LABEL_PAD_Y),
    );
    if (isActiveRow || isSelectedRow) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(rtl ? Math.round(innerX) : Math.round(innerX) - 2, rect.y, 2, h);
    }
  }

  // A hidden span leaves a bar across the strip: 4 device px before its
  // boundary and 2 after (colour measured in the dark appearance only).
  const barBefore = 4 / dpr;
  const barLength = 6 / dpr;
  ctx.fillStyle = theme.accent;
  for (const c of cols.visible) {
    if (!layout.hiddenCols.has(c + 1)) continue;
    const rect = cellRectIn(layout, cols, rows, firstRow, c);
    const edge = rtl ? rect.x : rect.x + (cols.sizeAt.get(c) ?? 0);
    ctx.fillRect(
      rtl ? edge - (barLength - barBefore) : edge - barBefore,
      labelTopY,
      barLength,
      oy - labelTopY,
    );
  }
  for (const r of rows.visible) {
    if (!layout.hiddenRows.has(r + 1)) continue;
    const rect = cellRectIn(layout, cols, rows, r, firstCol);
    ctx.fillRect(
      labelLeftX,
      rect.y + (rows.sizeAt.get(r) ?? 0) - barBefore,
      ox - (labelLeftX - 0),
      barLength,
    );
  }

  // Bracket gutters for outline groups.
  return paintOutlineGutters(ctx, state, theme, cols, rows, cssWidth, cssHeight);
}
