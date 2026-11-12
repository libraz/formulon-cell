// Row + column header bar and outline gutters. Drawn in
// a single pass after cell painting so headers sit above the gridlines.

import { activeFilterRanges } from '../../commands/filter-state.js';
import { colLetter } from '../../engine/address.js';
import { headerDigitPitch } from '../../store/header-strip.js';
import { renderScale } from '../../store/render-scale.js';
import { selectionRangeList } from '../../store/selection-geometry.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import {
  type AxisLayout,
  cellRectIn,
  colGap,
  gridOriginX,
  gridOriginY,
  rulerLeft,
  rulerTop,
  type ViewState,
} from '../geometry.js';
import { linearShade, paintOutlineGutters } from '../painters/controls.js';
import { fontMetricsBox, textBaselineY } from '../painters/text.js';
import type { OutlineToggleHit } from '../painters.js';
import { wholePointAdvances } from '../text-metrics.js';
import { type ChromePaintContext, ruleCenter, ruleWidth } from './chrome-context.js';

/** Row numbers are bottom-aligned like cell text: the baseline sits the
 *  label font's descent plus one px above the row's bottom edge (6 device px
 *  at DPR 2 for the default font, 8 for Courier New 11, 14 for Courier 20). */
const ROW_LABEL_PAD_Y = 1;

/** Screen extent of each page's visible columns in Page Layout view, or null
 *  in the views whose column strip runs unbroken. */
function pageColumnRuns(
  layout: ViewState['layout'],
  cols: AxisLayout,
  rows: AxisLayout,
  firstRow: number,
): { start: number; end: number }[] | null {
  if (!layout.pageGapCols || layout.pageGapCols.size === 0) return null;
  const runs: { start: number; end: number }[] = [];
  for (const c of cols.visible) {
    const rect = cellRectIn(layout, cols, rows, firstRow, c);
    const left = rect.x;
    const right = rect.x + (cols.sizeAt.get(c) ?? 0);
    const run = runs[runs.length - 1];
    if (!run || colGap(layout, c) > 0) runs.push({ start: left, end: right });
    else {
      run.start = Math.min(run.start, left);
      run.end = Math.max(run.end, right);
    }
  }
  return runs;
}

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

  // Each strip shades from its outer edge towards the grid.
  const rowStripOuterX = rtl ? labelLeftX + layout.headerColWidth : labelLeftX;
  ctx.fillStyle = linearShade(
    ctx,
    theme.headerStripOuter,
    theme.headerStripInner,
    0,
    labelTopY,
    0,
    oy,
  );
  ctx.fillRect(rtl ? 0 : ox, 0, cssWidth - ox, oy);
  ctx.fillStyle = linearShade(
    ctx,
    theme.headerStripOuter,
    theme.headerStripInner,
    rowStripOuterX,
    0,
    innerX,
    0,
  );
  ctx.fillRect(railLeftX, oy, ox, cssHeight - oy);

  const firstRow = rows.visible[0] ?? 0;
  const firstCol = cols.visible[0] ?? 0;
  // A selected chart or picture takes the highlight off the headers, as in Excel.
  const objectSelected = state.ui.selectedObject !== null;
  const selectedRanges = objectSelected ? [] : selectionRangeList(selection);
  const colSelected = (col: number): boolean =>
    selectedRanges.some((r) => col >= r.c0 && col <= r.c1);
  const rowSelected = (row: number): boolean =>
    selectedRanges.some((r) => row >= r.r0 && row <= r.r1);
  const isLitCol = (c: number): boolean => (!objectSelected && c === active.col) || colSelected(c);
  const isLitRow = (r: number): boolean => (!objectSelected && r === active.row) || rowSelected(r);

  // Selected headers refill before the strip border, which stays on top of them.
  ctx.fillStyle = theme.bgHeader;
  for (const c of cols.visible) {
    if (!isLitCol(c)) continue;
    const rect = cellRectIn(layout, cols, rows, firstRow, c);
    ctx.fillRect(rect.x, labelTopY, cols.sizeAt.get(c) ?? 0, layout.headerRowHeight);
  }
  for (const r of rows.visible) {
    if (!isLitRow(r)) continue;
    const rect = cellRectIn(layout, cols, rows, r, firstCol);
    ctx.fillRect(labelLeftX, rect.y, layout.headerColWidth, rows.sizeAt.get(r) ?? 0);
  }

  ctx.strokeStyle = theme.ruleStrong;
  const width = ruleWidth(dpr);
  ctx.lineWidth = width;
  // Header separators sit on the same trailing rule as the gridlines, so the
  // header band ends with its own rule rather than covering the first row.
  const headerRuleY = ruleCenter(oy, dpr);
  const headerRuleX = ruleCenter(innerX, dpr, rtl);
  const pageRuns = pageColumnRuns(layout, cols, rows, firstRow);
  ctx.beginPath();
  if (pageRuns) {
    // Page Layout boxes each page's letters: the rule stops at its columns.
    for (const run of pageRuns) {
      ctx.moveTo(run.start, headerRuleY);
      ctx.lineTo(run.end, headerRuleY);
    }
  } else {
    ctx.moveTo(0, headerRuleY);
    ctx.lineTo(cssWidth, headerRuleY);
  }
  ctx.moveTo(headerRuleX, 0);
  ctx.lineTo(headerRuleX, cssHeight);
  ctx.stroke();

  // A transparent bar token means the platform draws no selection bar.
  const activeBar = theme.headerActiveBar !== 'transparent';

  ctx.textAlign = 'center';
  const plainFont = `400 ${labelSize}px ${labelFamily}`;
  const litFont = `${theme.headerFgActiveWeight} ${labelSize}px ${labelFamily}`;
  ctx.font = plainFont;
  let currentFont = plainFont;
  const setFont = (font: string): void => {
    if (font === currentFont) return;
    currentFont = font;
    ctx.font = font;
  };
  const labelBox = fontMetricsBox(ctx, labelSize);
  // Mac sets column letters like a row number in a default-height row centred
  // on whole points in the strip: Excel's baseline sits 16 / 14 / 14 device px
  // above the strip bottom for Courier New 11 / Courier New 20 / the default font.
  const macLetters = wholePointAdvances(layout);
  const letterRowH = Math.min(
    layout.headerRowHeight,
    layout.defaultRowHeight * renderScale(layout, state.viewport),
  );
  const letterRow = {
    x: 0,
    y: labelTopY + Math.floor((layout.headerRowHeight - letterRowH) / 2),
    w: 0,
    h: letterRowH,
  };
  const letterY = macLetters
    ? textBaselineY(letterRow, labelBox, 'bottom', ROW_LABEL_PAD_Y)
    : labelTopY + layout.headerRowHeight / 2 + 0.5;
  ctx.textBaseline = macLetters ? 'alphabetic' : 'middle';
  const r1c1 = state.ui.r1c1 === true;
  // Separators shade along their own length and stop short of the strip's border rule.
  const colSepStyle = linearShade(
    ctx,
    theme.headerRuleOuter,
    theme.headerRule,
    0,
    labelTopY,
    0,
    oy - width,
  );
  // Separators go in one stroke; nothing painted per label reaches them.
  ctx.strokeStyle = colSepStyle;
  ctx.lineWidth = width;
  ctx.beginPath();
  for (const c of cols.visible) {
    const rect = cellRectIn(layout, cols, rows, firstRow, c);
    const w = cols.sizeAt.get(c) ?? 0;
    const sepX = ruleCenter(rtl ? rect.x : rect.x + w, dpr, rtl);
    ctx.moveTo(sepX, labelTopY);
    ctx.lineTo(sepX, oy - width);
    if (pageRuns && colGap(layout, c) > 0) {
      // A page's first column closes its letter box on the leading side too.
      const leadX = ruleCenter(rtl ? rect.x + w : rect.x, dpr, rtl);
      ctx.moveTo(leadX, labelTopY);
      ctx.lineTo(leadX, oy - width);
    }
  }
  ctx.stroke();
  for (const c of cols.visible) {
    const rect = cellRectIn(layout, cols, rows, firstRow, c);
    const w = cols.sizeAt.get(c) ?? 0;
    const lit = isLitCol(c);
    ctx.fillStyle = lit ? theme.headerFgActive : theme.headerFg;
    setFont(lit ? litFont : plainFont);
    const label = r1c1 ? `C${c + 1}` : colLetter(c);
    // Letters centre on the header cell less its trailing separator.
    const labelX = rtl ? rect.x + (w + width) / 2 : rect.x + (w - width) / 2;
    ctx.fillText(label, labelX, letterY);
    // The accent bar covers the rule under the header and one pixel above it.
    if (lit && activeBar) {
      ctx.fillStyle = theme.headerActiveBar;
      ctx.fillRect(rect.x, Math.round(oy) - 2, w, 2);
    }
  }

  // Row numbers centre on the strip between the rail edge and its rule.
  ctx.textAlign = 'center';
  ctx.textBaseline = 'alphabetic';
  const rowLabelX = rtl
    ? innerX + (layout.headerColWidth + width) / 2
    : innerX - (layout.headerColWidth + width) / 2;
  const digitPitch = headerDigitPitch(layout, state.viewport.zoom);
  const rowSepX0 = rtl ? innerX + width : labelLeftX;
  const rowSepX1 = rtl ? innerX + layout.headerColWidth : innerX - width;
  const rowSepStyle = linearShade(
    ctx,
    theme.headerRuleOuter,
    theme.headerRule,
    rtl ? rowSepX1 : rowSepX0,
    0,
    rtl ? rowSepX0 : rowSepX1,
    0,
  );
  ctx.strokeStyle = rowSepStyle;
  ctx.lineWidth = width;
  ctx.beginPath();
  for (const r of rows.visible) {
    const rect = cellRectIn(layout, cols, rows, r, firstCol);
    const sepY = ruleCenter(rect.y + (rows.sizeAt.get(r) ?? 0), dpr);
    ctx.moveTo(rowSepX0, sepY);
    ctx.lineTo(rowSepX1, sepY);
  }
  ctx.stroke();
  const filters = activeFilterRanges(state);
  for (const r of rows.visible) {
    const rect = cellRectIn(layout, cols, rows, r, firstCol);
    const h = rows.sizeAt.get(r) ?? 0;
    const lit = isLitRow(r);
    const filtered = filters.some(({ range }) => r > range.r0 && r <= range.r1);
    ctx.fillStyle = filtered ? theme.headerFgFiltered : lit ? theme.headerFgActive : theme.headerFg;
    setFont(lit ? litFont : plainFont);
    const rowLabel = r1c1 ? `R${r + 1}` : String(r + 1);
    const rowLabelY = textBaselineY({ ...rect, h }, labelBox, 'bottom', ROW_LABEL_PAD_Y);
    if (digitPitch !== undefined && !r1c1) {
      // Digits step by the strip's digit width; the run centres on whole
      // points from the strip's outer edge (Excel Courier 11 row 10: -40 dev px).
      ctx.textAlign = 'left';
      const inset = Math.floor((layout.headerColWidth - rowLabel.length * digitPitch) / 2);
      const left = rtl
        ? innerX + layout.headerColWidth - inset - rowLabel.length * digitPitch
        : labelLeftX + inset;
      for (let i = 0; i < rowLabel.length; i += 1) {
        ctx.fillText(rowLabel.charAt(i), left + i * digitPitch, rowLabelY);
      }
      ctx.textAlign = 'center';
    } else {
      ctx.fillText(rowLabel, rowLabelX, rowLabelY);
    }
    if (lit && activeBar) {
      ctx.fillStyle = theme.headerActiveBar;
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
