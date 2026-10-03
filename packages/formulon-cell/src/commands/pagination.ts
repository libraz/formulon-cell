import { parseAddrKey } from '../engine/address.js';
// Page pagination for the Page Layout and Page Break Preview workbook views.
//
// `print.ts` owns the same split for the printed document; this module reuses
// its band splitter so the blue break lines the user drags in Page Break
// Preview land on exactly the page boundaries the print pipeline produces.
// The difference is reach: printing stops at the last populated cell, whereas
// the views keep laying out empty pages past it so the grid never runs out of
// paper as the user scrolls.
import type { LayoutSlice, PageSetup, State } from '../store/store.js';
import { getPageSetup } from '../store/store.js';
import {
  type AxisBand,
  computeFitToPagesScale,
  effectivePrintMargins,
  PRINT_PX_PER_INCH,
  paperInches,
  parsePrintArea,
  printablePagePixels,
  splitAxisIntoBands,
} from './print.js';

export type { AxisBand } from './print.js';

/** Page geometry for one sheet, in unscaled sheet pixels. */
export interface SheetPagination {
  sheet: number;
  /** Row bands top-to-bottom; every rendered row falls in exactly one. */
  rowBands: AxisBand[];
  /** Column bands left-to-right, in sheet order (never mirrored for RTL). */
  colBands: AxisBand[];
  /** First row / column the pagination is anchored at — the print area's
   *  top-left when one is set, otherwise A1. */
  origin: { row: number; col: number };
  /** Last row / column that carries printable content. Bands past this are
   *  the blank pages the views keep drawing; the preview greys them out. */
  content: { row: number; col: number };
  /** Print scale in force, fit-to-pages already resolved. */
  scale: number;
  /** Printable box of one page. */
  pagePx: { width: number; height: number };
  /** Page margins, converted from inches at the current scale. */
  marginPx: { top: number; right: number; bottom: number; left: number };
  /** Sheet pixels per printed inch at the current scale. Multiply by the
   *  viewport zoom to get screen pixels — the rulers measure with this. */
  pxPerInch: number;
  /** Number of pages that actually carry content. */
  pageCount: number;
}

/** Zero-based band index containing `index`, or -1 when it sits before the
 *  first band. Bands are contiguous, so this is a plain scan. */
export function bandIndexAt(bands: readonly AxisBand[], index: number): number {
  for (let i = 0; i < bands.length; i += 1) {
    const band = bands[i];
    if (!band) continue;
    if (index >= band.start && index <= band.end) return i;
  }
  return -1;
}

/**
 * 1-based page number for a band pair. Pages past the content extent are not
 * printed and get 0 — the preview uses that to skip the page-number watermark
 * on blank paper.
 */
export function pageNumberOf(
  pagination: SheetPagination,
  rowBand: number,
  colBand: number,
  pageOrder: PageSetup['pageOrder'] = 'downThenOver',
): number {
  const rows = pagination.rowBands;
  const cols = pagination.colBands;
  const liveRows = rows.filter((b) => b.start <= pagination.content.row).length;
  const liveCols = cols.filter((b) => b.start <= pagination.content.col).length;
  if (rowBand < 0 || colBand < 0 || rowBand >= liveRows || colBand >= liveCols) return 0;
  return pageOrder === 'overThenDown'
    ? rowBand * liveCols + colBand + 1
    : colBand * liveRows + rowBand + 1;
}

/** Last populated row / column on `sheet`, from the cells and formats the
 *  store holds. Mirrors what the print builder walks. */
function usedExtent(state: State, sheet: number): { row: number; col: number } {
  let row = -1;
  let col = -1;
  const prefix = `${sheet}:`;
  for (const key of state.data.cells.keys()) {
    if (!key.startsWith(prefix)) continue;
    const addr = parseAddrKey(key);
    if (!addr) continue;
    if (addr.row > row) row = addr.row;
    if (addr.col > col) col = addr.col;
  }
  for (const key of state.format.formats.keys()) {
    if (!key.startsWith(prefix)) continue;
    const addr = parseAddrKey(key);
    if (!addr) continue;
    if (addr.row > row) row = addr.row;
    if (addr.col > col) col = addr.col;
  }
  for (const [, range] of state.merges.byAnchor) {
    if (range.sheet !== sheet) continue;
    if (range.r1 > row) row = range.r1;
    if (range.c1 > col) col = range.c1;
  }
  return { row, col };
}

export interface PaginateOptions {
  /** Paginate at least this far so the views can draw pages past the data.
   *  Rounded up internally, so scrolling one row does not invalidate the cache. */
  throughRow?: number;
  throughCol?: number;
}

/** Reach is rounded to this many indices so a one-row scroll reuses the cache. */
const REACH_CHUNK = 64;

function reach(through: number | undefined, content: number): number {
  const wanted = Math.max(through ?? 0, content, 0);
  return (Math.floor(wanted / REACH_CHUNK) + 1) * REACH_CHUNK;
}

function paginate(state: State, sheet: number, opts: PaginateOptions): SheetPagination {
  const setup = getPageSetup(state, sheet);
  const layout: LayoutSlice = state.layout;
  const used = usedExtent(state, sheet);
  const area = parsePrintArea(setup.printArea);
  const origin = { row: area?.row0 ?? 0, col: area?.col0 ?? 0 };
  const content = {
    row: area ? area.row1 : Math.max(used.row, origin.row),
    col: area ? area.col1 : Math.max(used.col, origin.col),
  };

  const region = { row0: origin.row, row1: content.row, col0: origin.col, col1: content.col };
  const scale = computeFitToPagesScale(
    setup,
    [region],
    layout,
    layout.hiddenRows,
    layout.hiddenCols,
  );
  const pagePx = printablePagePixels(setup, scale);

  const rowBands = splitAxisIntoBands({
    from: origin.row,
    to: reach(opts.throughRow, content.row),
    budget: Math.max(layout.defaultRowHeight, pagePx.height),
    manualBreaks: setup.manualPageBreakRows,
    sizeOf: (row) =>
      layout.hiddenRows.has(row) ? 0 : (layout.rowHeights.get(row) ?? layout.defaultRowHeight),
  });
  const colBands = splitAxisIntoBands({
    from: origin.col,
    to: reach(opts.throughCol, content.col),
    budget: Math.max(layout.defaultColWidth, pagePx.width),
    manualBreaks: setup.manualPageBreakCols,
    sizeOf: (col) =>
      layout.hiddenCols.has(col) ? 0 : (layout.colWidths.get(col) ?? layout.defaultColWidth),
  });

  const margins = effectivePrintMargins(setup);
  const perInch = PRINT_PX_PER_INCH / Math.max(scale, 0.1);
  const liveRows = rowBands.filter((b) => b.start <= content.row).length;
  const liveCols = colBands.filter((b) => b.start <= content.col).length;

  return {
    sheet,
    rowBands,
    colBands,
    origin,
    content,
    scale,
    pagePx,
    marginPx: {
      top: margins.top * perInch,
      right: margins.right * perInch,
      bottom: margins.bottom * perInch,
      left: margins.left * perInch,
    },
    pxPerInch: perInch,
    pageCount: liveRows * liveCols,
  };
}

/** Physical page box in unscaled sheet pixels — printable area plus margins. */
export function pageBoxPx(
  pagination: SheetPagination,
  setup: PageSetup,
): {
  width: number;
  height: number;
} {
  const inches = paperInches(setup);
  const perInch = PRINT_PX_PER_INCH / Math.max(pagination.scale, 0.1);
  return { width: inches.w * perInch, height: inches.h * perInch };
}

interface CacheEntry {
  key: unknown[];
  value: SheetPagination;
}

let cache: CacheEntry | null = null;

const sameKey = (a: unknown[], b: unknown[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

/**
 * Pagination for `sheet`, memoised against the state slices it reads. Every
 * hit-test and paint pass in Page Layout view goes through here, so a pointer
 * move must not re-walk the workbook; the store replaces slice objects on
 * write, which makes identity a sound cache key.
 */
export function paginationFor(
  state: State,
  sheet: number,
  opts: PaginateOptions = {},
): SheetPagination {
  const key = [
    sheet,
    state.layout,
    state.data.cells,
    state.format.formats,
    state.merges.byAnchor,
    state.pageSetup,
    reach(opts.throughRow, 0),
    reach(opts.throughCol, 0),
  ];
  if (cache && sameKey(cache.key, key)) return cache.value;
  const value = paginate(state, sheet, opts);
  cache = { key, value };
  return value;
}

/** Drop the memoised pagination. Tests use this to isolate cases; production
 *  code never needs it because the cache key tracks every input. */
export function resetPaginationCache(): void {
  cache = null;
}

/** Grey gutter between two sheets of paper in Page Layout view, in screen
 *  pixels. Fixed rather than zoom-scaled: it is chrome separating pages, not
 *  part of either page. */
export const PAGE_GUTTER_PX = 12;

export interface PageLayoutGaps {
  pageGapRows: Map<number, number>;
  pageGapCols: Map<number, number>;
}

/**
 * Leading gutters that turn the continuous grid into a stack of pages, keyed
 * by the row / column that opens each page and already multiplied by the
 * viewport zoom. The first page contributes only its own leading margin; every
 * page after it also carries the previous page's trailing margin and the grey
 * strip between the two.
 */
export function pageLayoutGaps(
  state: State,
  sheet: number,
  opts: PaginateOptions = {},
): PageLayoutGaps {
  const pagination = paginationFor(state, sheet, opts);
  const zoom = state.viewport.zoom > 0 ? state.viewport.zoom : 1;
  const m = pagination.marginPx;
  const pageGapRows = new Map<number, number>();
  const pageGapCols = new Map<number, number>();
  pagination.rowBands.forEach((band, i) => {
    const lead = i === 0 ? m.top : m.bottom + m.top;
    pageGapRows.set(band.start, lead * zoom + (i === 0 ? 0 : PAGE_GUTTER_PX));
  });
  pagination.colBands.forEach((band, i) => {
    const lead = i === 0 ? m.left : m.right + m.left;
    pageGapCols.set(band.start, lead * zoom + (i === 0 ? 0 : PAGE_GUTTER_PX));
  });
  return { pageGapRows, pageGapCols };
}
