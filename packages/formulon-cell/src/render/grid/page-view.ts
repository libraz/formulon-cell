// Page Layout and Page Break Preview chrome.
//
// Both views draw the same page grid, computed once by `commands/pagination`
// and folded into the axis as leading gutters (Page Layout) or read straight
// off the bands (Page Break Preview). Nothing here changes what a cell paints;
// the passes only add the paper, the break lines, and the page numbering that
// tell the user where the printed sheet ends.

import { bandIndexAt, pageNumberOf, type SheetPagination } from '../../commands/pagination.js';
import type { PageSetup } from '../../store/store.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import {
  type AxisLayout,
  gridOriginX,
  gridOriginY,
  mirrorX,
  type Rect,
  rulerTop,
  type ViewState,
} from '../geometry.js';
import type { ChromePaintContext } from './chrome-context.js';

/** A stretch of one axis that is visible on screen and belongs to one page. */
interface BandRun {
  /** Index into `pagination.rowBands` / `colBands`. -1 before the first page. */
  band: number;
  /** Screen position of the run's first cell edge, in laid-out coordinates. */
  start: number;
  /** Screen position past the run's last cell edge. */
  end: number;
  /** The run begins at the band's own first index, so the page's leading
   *  margin is on screen above / left of it. */
  opens: boolean;
  /** The run reaches the band's last index, so the trailing margin follows. */
  closes: boolean;
}

type SegmentKind = 'content' | 'paper' | 'gutter';

interface Segment {
  start: number;
  end: number;
  kind: SegmentKind;
}

function runsFor(axis: AxisLayout, bands: SheetPagination['rowBands'], origin: number): BandRun[] {
  const runs: BandRun[] = [];
  for (const index of axis.visible) {
    const pos = axis.positionAt.get(index);
    const size = axis.sizeAt.get(index);
    if (pos === undefined || size === undefined) continue;
    const band = bandIndexAt(bands, index);
    const last = runs[runs.length - 1];
    if (last && last.band === band) {
      last.end = origin + pos + size;
      last.closes = band >= 0 && index === bands[band]?.end;
      continue;
    }
    runs.push({
      band,
      start: origin + pos,
      end: origin + pos + size,
      opens: band >= 0 && index === bands[band]?.start,
      closes: band >= 0 && index === bands[band]?.end,
    });
  }
  return runs;
}

/** Tile the visible extent of one axis into content / margin / gutter spans.
 *  The two axes' segment lists cross-multiply into the page chrome: a cell of
 *  the product is paper unless either side is a gutter. */
function segmentsFor(runs: readonly BandRun[], leadMargin: number, trailMargin: number): Segment[] {
  const segments: Segment[] = [];
  let cursor = Number.NEGATIVE_INFINITY;
  for (const run of runs) {
    const paperStart = run.opens ? run.start - leadMargin : run.start;
    const paperEnd = run.closes ? run.end + trailMargin : run.end;
    if (cursor > Number.NEGATIVE_INFINITY && paperStart > cursor) {
      segments.push({ start: cursor, end: paperStart, kind: 'gutter' });
    }
    if (paperStart < run.start) segments.push({ start: paperStart, end: run.start, kind: 'paper' });
    segments.push({ start: run.start, end: run.end, kind: 'content' });
    if (paperEnd > run.end) segments.push({ start: run.end, end: paperEnd, kind: 'paper' });
    cursor = paperEnd;
  }
  return segments;
}

/** Paper rectangles for the visible pages, in laid-out coordinates. */
function paperRects(
  rowRuns: readonly BandRun[],
  colRuns: readonly BandRun[],
  margins: { top: number; right: number; bottom: number; left: number },
): Rect[] {
  const rects: Rect[] = [];
  for (const rowRun of rowRuns) {
    const top = rowRun.opens ? rowRun.start - margins.top : rowRun.start;
    const bottom = rowRun.closes ? rowRun.end + margins.bottom : rowRun.end;
    for (const colRun of colRuns) {
      const left = colRun.opens ? colRun.start - margins.left : colRun.start;
      const right = colRun.closes ? colRun.end + margins.right : colRun.end;
      rects.push({ x: left, y: top, w: right - left, h: bottom - top });
    }
  }
  return rects;
}

interface PageAxes {
  rowRuns: BandRun[];
  colRuns: BandRun[];
  /** Margins in screen pixels — pagination stores them unscaled. */
  margins: { top: number; right: number; bottom: number; left: number };
  originX: number;
  originY: number;
}

function pageAxes(
  state: ViewState,
  pagination: SheetPagination,
  cols: AxisLayout,
  rows: AxisLayout,
): PageAxes {
  const zoom = state.viewport.zoom > 0 ? state.viewport.zoom : 1;
  const m = pagination.marginPx;
  return {
    rowRuns: runsFor(rows, pagination.rowBands, gridOriginY(state.layout)),
    colRuns: runsFor(cols, pagination.colBands, gridOriginX(state.layout)),
    margins: {
      top: m.top * zoom,
      right: m.right * zoom,
      bottom: m.bottom * zoom,
      left: m.left * zoom,
    },
    originX: gridOriginX(state.layout),
    originY: gridOriginY(state.layout),
  };
}

/** Restrict painting to the data area so page chrome never bleeds into the
 *  row / column header rails. */
function clipToData(pc: ChromePaintContext, state: ViewState): void {
  const { ctx, cssWidth, cssHeight } = pc;
  const ox = gridOriginX(state.layout);
  const oy = gridOriginY(state.layout);
  const left = state.layout.rtl ? 0 : ox;
  const width = Math.max(0, cssWidth - ox);
  ctx.beginPath();
  ctx.rect(left, oy, width, Math.max(0, cssHeight - oy));
  ctx.clip();
}

/**
 * Background pass for Page Layout view: the desk the pages sit on, then the
 * paper itself. Runs before the cells so a cell with no fill of its own reads
 * as printed on the page rather than as a hole in it.
 */
export function paintPageLayoutBackground(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  pagination: SheetPagination,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { ctx, cssWidth, cssHeight } = pc;
  const axes = pageAxes(state, pagination, cols, rows);
  ctx.save();
  clipToData(pc, state);
  ctx.fillStyle = theme.pageBackdrop;
  ctx.fillRect(0, 0, cssWidth, cssHeight);
  ctx.fillStyle = theme.pagePaper;
  for (const rect of paperRects(axes.rowRuns, axes.colRuns, axes.margins)) {
    const x = mirrorX(state.layout, rect.x, rect.w);
    ctx.fillRect(x, rect.y, rect.w, rect.h);
  }
  ctx.restore();
}

/** Rect of one header / footer slot, plus what it edits. */
export interface PageBandHit {
  rect: Rect;
  slot: 'left' | 'center' | 'right';
  kind: 'header' | 'footer';
}

let pageBandHits: PageBandHit[] = [];

/** Header / footer slots drawn by the last Page Layout paint, in screen
 *  coordinates. The pointer layer reads this to open the inline editor. */
export function getPageBandHits(): readonly PageBandHit[] {
  return pageBandHits;
}

const BAND_SLOTS: PageBandHit['slot'][] = ['left', 'center', 'right'];

function bandText(setup: PageSetup, kind: PageBandHit['kind'], slot: PageBandHit['slot']): string {
  if (kind === 'header') {
    if (slot === 'left') return setup.headerLeft ?? '';
    if (slot === 'center') return setup.headerCenter ?? '';
    return setup.headerRight ?? '';
  }
  if (slot === 'left') return setup.footerLeft ?? '';
  if (slot === 'center') return setup.footerCenter ?? '';
  return setup.footerRight ?? '';
}

export interface PageBandStrings {
  addHeader: string;
  addFooter: string;
}

/**
 * Foreground pass for Page Layout view. Repaints everything the cells are not
 * allowed to cover — the gutters between sheets and the margin bands inside
 * them — then frames each page and lays the header / footer slots into the top
 * and bottom margins.
 */
export function paintPageLayoutChrome(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  pagination: SheetPagination,
  setup: PageSetup,
  strings: PageBandStrings,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { ctx, dpr } = pc;
  const axes = pageAxes(state, pagination, cols, rows);
  const rowSegs = segmentsFor(axes.rowRuns, axes.margins.top, axes.margins.bottom);
  const colSegs = segmentsFor(axes.colRuns, axes.margins.left, axes.margins.right);

  ctx.save();
  clipToData(pc, state);

  // Everything except the cell content: a product cell is paper when neither
  // axis is in a gutter, and desk otherwise.
  for (const rowSeg of rowSegs) {
    for (const colSeg of colSegs) {
      if (rowSeg.kind === 'content' && colSeg.kind === 'content') continue;
      const gutter = rowSeg.kind === 'gutter' || colSeg.kind === 'gutter';
      ctx.fillStyle = gutter ? theme.pageBackdrop : theme.pagePaper;
      const w = colSeg.end - colSeg.start;
      ctx.fillRect(
        mirrorX(state.layout, colSeg.start, w),
        rowSeg.start,
        w,
        rowSeg.end - rowSeg.start,
      );
    }
  }

  const rects = paperRects(axes.rowRuns, axes.colRuns, axes.margins);
  ctx.strokeStyle = theme.pageEdge;
  ctx.lineWidth = 1 / dpr;
  const align = 0.5 / dpr;
  for (const rect of rects) {
    const x = mirrorX(state.layout, rect.x, rect.w);
    ctx.strokeRect(Math.round(x) + align, Math.round(rect.y) + align, rect.w, rect.h);
  }

  pageBandHits = paintPageBands(pc, state, theme, axes, setup, strings);
  ctx.restore();
}

/** Height of a header / footer slot inside its margin. */
const BAND_HEIGHT = 22;

function paintPageBands(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  axes: PageAxes,
  setup: PageSetup,
  strings: PageBandStrings,
): PageBandHit[] {
  const { ctx } = pc;
  const hits: PageBandHit[] = [];
  // A band needs its whole page width, so only pages whose column run spans
  // the full band get one — a page sliced by the viewport edge would otherwise
  // centre its text against the visible fragment.
  ctx.font = `12px ${theme.fontUi}`;
  ctx.textBaseline = 'middle';
  for (const rowRun of axes.rowRuns) {
    for (const colRun of axes.colRuns) {
      if (!colRun.opens || !colRun.closes) continue;
      const left = colRun.start;
      const width = colRun.end - colRun.start;
      if (width < 90) continue;
      const bands: { kind: PageBandHit['kind']; y: number }[] = [];
      if (rowRun.opens && axes.margins.top >= BAND_HEIGHT) {
        bands.push({ kind: 'header', y: rowRun.start - axes.margins.top + 4 });
      }
      if (rowRun.closes && axes.margins.bottom >= BAND_HEIGHT) {
        bands.push({ kind: 'footer', y: rowRun.end + axes.margins.bottom - BAND_HEIGHT - 4 });
      }
      for (const band of bands) {
        const slotW = width / 3;
        for (let i = 0; i < BAND_SLOTS.length; i += 1) {
          const slot = BAND_SLOTS[i] as PageBandHit['slot'];
          const laidOutX = left + slotW * i;
          const rect: Rect = {
            x: mirrorX(state.layout, laidOutX, slotW),
            y: band.y,
            w: slotW,
            h: BAND_HEIGHT,
          };
          const text = bandText(setup, band.kind, slot);
          const placeholder = band.kind === 'header' ? strings.addHeader : strings.addFooter;
          // Only the centre slot advertises itself; three "Add header" prompts
          // per page would drown the sheet.
          const label = text || (slot === 'center' ? placeholder : '');
          if (label) {
            ctx.fillStyle = text ? theme.fg : theme.pageBandFg;
            ctx.textAlign = slot === 'left' ? 'left' : slot === 'right' ? 'right' : 'center';
            const tx =
              slot === 'left'
                ? rect.x + 6
                : slot === 'right'
                  ? rect.x + rect.w - 6
                  : rect.x + rect.w / 2;
            ctx.fillText(label, tx, rect.y + rect.h / 2, rect.w - 12);
          }
          hits.push({ rect, slot, kind: band.kind });
        }
      }
    }
  }
  ctx.textAlign = 'left';
  return hits;
}

/** A margin boundary on one of the Page Layout rulers. */
export interface RulerHandle {
  side: 'left' | 'right' | 'top' | 'bottom';
  /** Screen coordinate of the boundary on its own axis. */
  position: number;
  /** Screen extent of the ruler band on the *other* axis. Carried on the
   *  handle because a right-to-left sheet puts the vertical ruler on the
   *  right, and the pointer layer should not have to work that out again. */
  bandStart: number;
  bandEnd: number;
  /** Screen coordinate the page's paper starts at on that axis, so a drag can
   *  turn a dropped position back into a margin width. */
  paperOrigin: number;
  /** Screen pixels per printed inch, zoom applied. */
  pxPerInch: number;
}

let rulerHandles: RulerHandle[] = [];

/** Margin boundaries drawn by the last Page Layout paint. */
export function getRulerHandles(): readonly RulerHandle[] {
  return rulerHandles;
}

/** Pointer slack around a ruler margin boundary, in pixels. */
export const RULER_GRAB = 4;

const CM_PER_INCH = 2.54;

/**
 * The Page Layout rulers.
 *
 * Each ruler runs the length of the page it belongs to: the margins read as
 * desk, the printable width as paper, and the boundary between them is the
 * handle that sets the margin. Ticks are numbered from the printable origin,
 * which is where the desktop app zeroes them too, and in the locale's unit —
 * centimetres almost everywhere, inches in the US.
 */
export function paintPageRulers(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  pagination: SheetPagination,
  unit: 'in' | 'cm',
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { ctx, dpr, cssWidth, cssHeight } = pc;
  const layout = state.layout;
  const band = rulerTop(layout);
  if (band <= 0) return;
  const axes = pageAxes(state, pagination, cols, rows);
  const zoom = state.viewport.zoom > 0 ? state.viewport.zoom : 1;
  const pxPerInch = pagination.pxPerInch * zoom;
  const pxPerUnit = unit === 'cm' ? pxPerInch / CM_PER_INCH : pxPerInch;
  const handles: RulerHandle[] = [];
  const align = 0.5 / dpr;

  const ox = gridOriginX(layout);
  const oy = gridOriginY(layout);
  // The row rail — and with it the vertical ruler — swaps to the right edge on
  // a right-to-left sheet.
  const vBandX = layout.rtl ? cssWidth - band : 0;
  ctx.save();
  ctx.fillStyle = theme.bgRail;
  ctx.fillRect(0, 0, cssWidth, band);
  ctx.fillRect(vBandX, 0, band, cssHeight);
  ctx.font = `9px ${theme.fontUi}`;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';

  /** One ruler segment: desk for the margins, paper for the printable span.
   *  Each band is clipped to the data area on its own axis so neither ruler
   *  runs across the header rails or the corner they share. */
  const paintBand = (
    runs: readonly BandRun[],
    lead: number,
    trail: number,
    horizontal: boolean,
  ): void => {
    ctx.save();
    ctx.beginPath();
    if (horizontal) {
      const left = layout.rtl ? 0 : ox;
      ctx.rect(left, 0, Math.max(0, cssWidth - ox), band);
    } else {
      ctx.rect(vBandX, oy, band, Math.max(0, cssHeight - oy));
    }
    ctx.clip();
    for (const run of runs) {
      const paperStart = run.opens ? run.start - lead : run.start;
      const paperEnd = run.closes ? run.end + trail : run.end;
      const fill = (from: number, to: number, colour: string): void => {
        if (to <= from) return;
        ctx.fillStyle = colour;
        if (horizontal) {
          const w = to - from;
          ctx.fillRect(mirrorX(layout, from, w), 2, w, band - 4);
        } else {
          ctx.fillRect(vBandX + 2, from, band - 4, to - from);
        }
      };
      fill(paperStart, run.start, theme.pageBackdrop);
      fill(run.start, run.end, theme.pagePaper);
      fill(run.end, paperEnd, theme.pageBackdrop);

      // Ticks, numbered outward from the printable origin.
      ctx.fillStyle = theme.headerFg;
      const first = Math.ceil((paperStart - run.start) / pxPerUnit);
      const last = Math.floor((paperEnd - run.start) / pxPerUnit);
      for (let step = first; step <= last; step += 1) {
        const at = run.start + step * pxPerUnit;
        if (step === 0) continue;
        const label = String(Math.abs(step));
        if (horizontal) ctx.fillText(label, mirrorX(layout, at), band / 2);
        else ctx.fillText(label, vBandX + band / 2, at);
      }

      const bandStart = horizontal ? 0 : vBandX;
      const bandEnd = bandStart + band;
      if (run.opens) {
        handles.push({
          side: horizontal ? 'left' : 'top',
          position: horizontal ? mirrorX(layout, run.start) : run.start,
          paperOrigin: horizontal ? mirrorX(layout, paperStart) : paperStart,
          bandStart,
          bandEnd,
          pxPerInch,
        });
      }
      if (run.closes) {
        handles.push({
          side: horizontal ? 'right' : 'bottom',
          position: horizontal ? mirrorX(layout, run.end) : run.end,
          paperOrigin: horizontal ? mirrorX(layout, paperEnd) : paperEnd,
          bandStart,
          bandEnd,
          pxPerInch,
        });
      }
    }

    // Margin grips, drawn last so they sit over the segments they divide.
    ctx.strokeStyle = theme.pageEdge;
    ctx.lineWidth = 1 / dpr;
    ctx.beginPath();
    for (const handle of handles) {
      if (horizontal !== (handle.side === 'left' || handle.side === 'right')) continue;
      if (horizontal) {
        const x = Math.round(handle.position) + align;
        ctx.moveTo(x, 2);
        ctx.lineTo(x, band - 2);
      } else {
        const y = Math.round(handle.position) + align;
        ctx.moveTo(vBandX + 2, y);
        ctx.lineTo(vBandX + band - 2, y);
      }
    }
    ctx.stroke();
    ctx.restore();
  };

  paintBand(axes.colRuns, axes.margins.left, axes.margins.right, true);
  paintBand(axes.rowRuns, axes.margins.top, axes.margins.bottom, false);
  ctx.textAlign = 'left';
  ctx.restore();

  rulerHandles = handles;
}

/** A line the user can drag in Page Break Preview. */
export interface PageBreakHandle {
  axis: 'row' | 'col';
  /** A page boundary, or the trailing edge of the printed area — dragging the
   *  latter resizes the print area instead of moving a break. */
  kind: 'break' | 'printArea';
  /** Index the page starts at — the break sits immediately before it. For a
   *  print-area edge this is the last printed index. */
  index: number;
  /** Screen coordinate of the line on its own axis. */
  position: number;
  manual: boolean;
}

let pageBreakHandles: PageBreakHandle[] = [];

/** Break lines drawn by the last preview paint. The pointer layer hit-tests
 *  against these to start a drag. */
export function getPageBreakHandles(): readonly PageBreakHandle[] {
  return pageBreakHandles;
}

/** Pointer slack around a break line, in pixels. */
export const PAGE_BREAK_GRAB = 4;

/**
 * Page Break Preview: wash out everything that will not print, frame the print
 * area, draw every page boundary, and number the pages. Runs after the cells so
 * the lines sit on top, and before the selection so the active outline still
 * reads clearly.
 */
export function paintPageBreakPreview(
  pc: ChromePaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  pagination: SheetPagination,
  setup: PageSetup,
  pageLabel: (page: number) => string,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { ctx, dpr, cssWidth, cssHeight } = pc;
  const layout = state.layout;
  const ox = gridOriginX(layout);
  const oy = gridOriginY(layout);
  const rowRuns = runsFor(rows, pagination.rowBands, oy);
  const colRuns = runsFor(cols, pagination.colBands, ox);
  const handles: PageBreakHandle[] = [];

  ctx.save();
  clipToData(pc, state);

  // Cells past the printed extent are still editable but never print, so they
  // get the same wash the desktop preview uses. The frame traces the printed
  // range itself, not the page it ends on — a print area of A1:F5 is framed at
  // F5 even though page 1 reaches much further.
  const liveRows = rowRuns.filter((run) => isLiveBand(pagination, run.band, 'row'));
  const liveCols = colRuns.filter((run) => isLiveBand(pagination, run.band, 'col'));
  const printBottom = trailingEdge(rows, pagination.content.row, oy, cssHeight);
  const printTrail = trailingEdge(cols, pagination.content.col, ox, cssWidth);
  const printTop = leadingEdge(rows, pagination.origin.row, oy, oy);
  const printLead = leadingEdge(cols, pagination.origin.col, ox, ox);
  ctx.fillStyle = theme.pageOutside;
  ctx.fillRect(0, oy, cssWidth, Math.max(0, printTop - oy));
  ctx.fillRect(0, printBottom, cssWidth, Math.max(0, cssHeight - printBottom));
  const bandH = Math.max(0, printBottom - printTop);
  const leadW = Math.max(0, printLead - ox);
  ctx.fillRect(mirrorX(layout, ox, leadW), printTop, leadW, bandH);
  const trailW = Math.max(0, cssWidth - printTrail);
  ctx.fillRect(mirrorX(layout, printTrail, trailW), printTop, trailW, bandH);

  // Page numbers, centred on the page they belong to.
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillStyle = theme.pageNumberFg;
  for (const rowRun of liveRows) {
    // The last page is usually only part-filled, so the number centres on the
    // printed part of it rather than on the empty paper below.
    const bottom = Math.min(rowRun.end, printBottom);
    for (const colRun of liveCols) {
      const page = pageNumberOf(pagination, rowRun.band, colRun.band, setup.pageOrder);
      if (page <= 0) continue;
      const trail = Math.min(colRun.end, printTrail);
      const w = trail - colRun.start;
      const h = bottom - rowRun.start;
      // Below this the number would be unreadable and would only obscure the
      // cells it sits on, so the page goes unlabelled.
      if (w < 48 || h < 24) continue;
      const size = Math.min(72, Math.max(20, Math.min(w, h) / 3));
      ctx.font = `700 ${size}px ${theme.fontUi}`;
      ctx.fillText(
        pageLabel(page),
        mirrorX(layout, colRun.start + w / 2),
        rowRun.start + h / 2,
        w - 16,
      );
    }
  }

  // Break lines. The frame around the print area is the same weight as a
  // manual break, which is what the desktop preview does.
  const align = 0.5 / dpr;
  const strokeLine = (
    axis: 'row' | 'col',
    position: number,
    manual: boolean,
    index: number,
  ): void => {
    ctx.strokeStyle = manual ? theme.pageBreakManual : theme.pageBreakAuto;
    ctx.lineWidth = (manual ? 2.5 : 1.5) / dpr;
    ctx.setLineDash(manual ? [] : [6 / dpr, 4 / dpr]);
    ctx.beginPath();
    if (axis === 'row') {
      const y = Math.round(position) + align;
      ctx.moveTo(ox, y);
      ctx.lineTo(cssWidth, y);
    } else {
      const x = Math.round(position) + align;
      ctx.moveTo(x, oy);
      ctx.lineTo(x, cssHeight);
    }
    ctx.stroke();
    handles.push({ axis, kind: 'break', index, position, manual });
  };

  // Only the printed pages get boundaries; the blank paper past the data has
  // no breaks to show and the desktop preview leaves it plain grey.
  for (const run of rowRuns) {
    const band = pagination.rowBands[run.band];
    if (!band || !run.opens || run.band === 0) continue;
    if (!isLiveBand(pagination, run.band, 'row')) continue;
    strokeLine('row', run.start, band.manual, band.start);
  }
  for (const run of colRuns) {
    const band = pagination.colBands[run.band];
    if (!band || !run.opens || run.band === 0) continue;
    if (!isLiveBand(pagination, run.band, 'col')) continue;
    strokeLine('col', mirrorX(layout, run.start), band.manual, band.start);
  }

  // Outer frame — the printed extent's trailing edges.
  ctx.setLineDash([]);
  ctx.strokeStyle = theme.pageBreakManual;
  ctx.lineWidth = 2.5 / dpr;
  ctx.beginPath();
  const frameBottom = Math.round(printBottom) + align;
  ctx.moveTo(ox, frameBottom);
  ctx.lineTo(cssWidth, frameBottom);
  const frameTrail = Math.round(mirrorX(layout, printTrail)) + align;
  ctx.moveTo(frameTrail, oy);
  ctx.lineTo(frameTrail, cssHeight);
  // A print area that does not start at A1 has leading edges of its own; when
  // it does start there these land on the grid origin and read as the rail.
  const frameTop = Math.round(leadingEdge(rows, pagination.origin.row, oy, oy)) + align;
  ctx.moveTo(ox, frameTop);
  ctx.lineTo(cssWidth, frameTop);
  const frameLead =
    Math.round(mirrorX(layout, leadingEdge(cols, pagination.origin.col, ox, ox))) + align;
  ctx.moveTo(frameLead, oy);
  ctx.lineTo(frameLead, cssHeight);
  ctx.stroke();
  // Dragging the frame grows or shrinks what prints, so it is a handle too.
  handles.push({
    axis: 'row',
    kind: 'printArea',
    index: pagination.content.row,
    position: printBottom,
    manual: true,
  });
  handles.push({
    axis: 'col',
    kind: 'printArea',
    index: pagination.content.col,
    position: mirrorX(layout, printTrail),
    manual: true,
  });

  // The line following the pointer mid-drag. The pages only reflow on drop,
  // so until then this is the only thing that moves.
  const dragPreview = state.ui.pageBreakDrag;
  if (dragPreview) {
    ctx.strokeStyle = theme.pageBreakManual;
    ctx.lineWidth = 2.5 / dpr;
    ctx.setLineDash([4 / dpr, 3 / dpr]);
    ctx.beginPath();
    if (dragPreview.axis === 'row') {
      const y = Math.round(dragPreview.position) + align;
      ctx.moveTo(ox, y);
      ctx.lineTo(cssWidth, y);
    } else {
      const x = Math.round(dragPreview.position) + align;
      ctx.moveTo(x, oy);
      ctx.lineTo(x, cssHeight);
    }
    ctx.stroke();
  }

  ctx.setLineDash([]);
  ctx.textAlign = 'left';
  ctx.restore();

  pageBreakHandles = handles;
}

/** Screen coordinate just past `index` on its axis. When the index has
 *  scrolled out of view the answer collapses to whichever end of the viewport
 *  it left by, so the wash still covers the right side of the boundary. */
function trailingEdge(axis: AxisLayout, index: number, origin: number, limit: number): number {
  const pos = axis.positionAt.get(index);
  if (pos !== undefined) return origin + pos + (axis.sizeAt.get(index) ?? 0);
  const first = axis.visible[0];
  if (first !== undefined && index < first) return origin;
  return limit;
}

/** Screen coordinate of `index`'s leading edge, clamped to the viewport when
 *  the index has scrolled out of view. */
function leadingEdge(axis: AxisLayout, index: number, origin: number, fallback: number): number {
  const pos = axis.positionAt.get(index);
  return pos === undefined ? fallback : origin + pos;
}

function isLiveBand(pagination: SheetPagination, band: number, axis: 'row' | 'col'): boolean {
  if (band < 0) return false;
  const bands = axis === 'row' ? pagination.rowBands : pagination.colBands;
  const limit = axis === 'row' ? pagination.content.row : pagination.content.col;
  const entry = bands[band];
  return !!entry && entry.start <= limit;
}
