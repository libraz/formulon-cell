import type { PaginationResult } from '@libraz/formulon';
import { parseAddrKey } from '../engine/address.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { renderScale } from '../store/render-scale.js';
import { hiddenRowsOf, isRowHidden } from '../store/row-visibility.js';
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
   *  render scale to get screen pixels — the rulers measure with this. */
  pxPerInch: number;
  /** Number of pages that actually carry content. */
  pageCount: number;
  /** Centering is kept beside native geometry because `PaginationResult`
   *  intentionally reports the uncentered cell-block origin. Page Layout
   *  adds the setting when placing the physical paper. */
  centerHorizontally?: boolean;
  centerVertically?: boolean;
  /** Detailed native page geometry when formulon exposes `paginate`. Preview
   *  leaves this unset for multiple print areas or overlapping repeated-title
   *  spans because the interactive grid can only draw a rectangular band
   *  product; print requests still consume every native page directly. The
   *  fallback fields above remain populated for those cases and older engines. */
  native?: PaginationResult;
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
  const row = rows[rowBand];
  const col = cols[colBand];
  if (pagination.native && row && col) {
    const nativeIndex = pagination.native.pages.findIndex(
      (page) => page.firstRow === row.start && page.firstCol === col.start,
    );
    if (nativeIndex >= 0) return nativeIndex + 1;
    return 0;
  }
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
  /** Workbook handle used to opt into formulon's detailed print geometry. */
  engine?: WorkbookHandle;
  /** Print consumes every native page directly; preview requires a
   *  rectangular Cartesian band model and falls back for layouts it cannot
   *  represent without inventing pages. */
  purpose?: 'preview' | 'print';
}

/** Reach is rounded to this many indices so a one-row scroll reuses the cache. */
const REACH_CHUNK = 64;

const paginationEngines = new WeakMap<object, WorkbookHandle>();

/** Associate the immutable store state object with its live workbook handle.
 * StoreMirror calls this for each replacement state so page-view callers can
 * opt into native geometry without threading the handle through every paint
 * helper. */
export function associatePaginationEngine(state: State, engine: WorkbookHandle | null): void {
  if (engine) paginationEngines.set(state, engine);
  else paginationEngines.delete(state);
  resetPaginationCache();
}

function reach(through: number | undefined, content: number): number {
  const wanted = Math.max(through ?? 0, content, 0);
  return (Math.floor(wanted / REACH_CHUNK) + 1) * REACH_CHUNK;
}

const CSS_PX_PER_POINT = 96 / 72;

function finitePositive(value: number, fallback: number): number {
  return Number.isFinite(value) && value > 0 ? value : fallback;
}

function nativeBands(native: PaginationResult, axis: 'row' | 'col'): AxisBand[] {
  const spans = new Map<string, { start: number; end: number }>();
  for (const page of native.pages) {
    const start = axis === 'row' ? page.firstRow : page.firstCol;
    const end = axis === 'row' ? page.lastRow : page.lastCol;
    if (!Number.isInteger(start) || !Number.isInteger(end) || end < start) continue;
    spans.set(`${start}:${end}`, { start, end });
  }
  const manualBreaks = new Map<number, boolean>();
  const breaks = axis === 'row' ? native.horizontalBreaks : native.verticalBreaks;
  const manual = axis === 'row' ? native.horizontalBreakManual : native.verticalBreakManual;
  for (let i = 0; i < breaks.length; i += 1) {
    const at = breaks[i];
    if (typeof at === 'number') manualBreaks.set(at, manual[i] === true);
  }
  return [...spans.values()]
    .sort((a, b) => a.start - b.start || a.end - b.end)
    .map(({ start, end }) => ({
      start,
      end,
      manual: manualBreaks.get(start) === true,
    }));
}

function hasOverlappingNativeSpans(native: PaginationResult, axis: 'row' | 'col'): boolean {
  const spans = new Map<string, [number, number]>();
  for (const page of native.pages) {
    const first = axis === 'row' ? page.firstRow : page.firstCol;
    const last = axis === 'row' ? page.lastRow : page.lastCol;
    spans.set(`${first}:${last}`, [first, last]);
  }
  const sorted = [...spans.values()].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  for (let i = 1; i < sorted.length; i += 1) {
    const previous = sorted[i - 1];
    const current = sorted[i];
    if (previous && current && current[0] <= previous[1]) return true;
  }
  return false;
}

function nativePreviewCompatible(native: PaginationResult): boolean {
  return (
    native.printArea.length <= 1 &&
    !hasOverlappingNativeSpans(native, 'row') &&
    !hasOverlappingNativeSpans(native, 'col')
  );
}

function extendNativeBands(
  bands: AxisBand[],
  axis: 'row' | 'col',
  state: State,
  setup: PageSetup,
  content: number,
  through: number | undefined,
  pageBudget: number,
  lastPageRoom: number,
): AxisBand[] {
  const target = reach(through, content);
  if (bands.length === 0) return bands;
  const last = bands[bands.length - 1];
  if (!last || last.end >= target) return bands;
  const hidden = axis === 'row' ? hiddenRowsOf(state.layout) : state.layout.hiddenCols;
  const defaultSize = axis === 'row' ? state.layout.defaultRowHeight : state.layout.defaultColWidth;
  const sizes = axis === 'row' ? state.layout.rowHeights : state.layout.colWidths;
  const manual = axis === 'row' ? setup.manualPageBreakRows : setup.manualPageBreakCols;
  const sizeOf = (index: number): number =>
    hidden.has(index) ? 0 : (sizes.get(index) ?? defaultSize);
  // Blank tracks past the content fill the last printed page before a new
  // one starts, as the desktop Page Layout view lays them out. A print area
  // bounds what prints, so its last page keeps its own extent.
  let end = last.end;
  if (!setup.printArea) {
    const breaks = new Set(manual ?? []);
    let room = lastPageRoom;
    while (end < target && !breaks.has(end + 1) && sizeOf(end + 1) <= room) {
      room -= sizeOf(end + 1);
      end += 1;
    }
  }
  const grown = end === last.end ? bands : [...bands.slice(0, -1), { ...last, end }];
  if (end >= target) return grown;
  const tail = splitAxisIntoBands({
    from: end + 1,
    to: target,
    budget: Math.max(defaultSize, pageBudget),
    manualBreaks: manual,
    sizeOf,
  });
  return grown.concat(tail);
}

/** Unscaled px left on the printable area after the native page that opens
 *  at `start`, print titles included in what the page already holds. */
function nativePageRoom(
  native: PaginationResult,
  start: number,
  axis: 'row' | 'col',
  scale: number,
): number {
  const page = native.pages.find((candidate) =>
    axis === 'row' ? candidate.firstRow === start : candidate.firstCol === start,
  );
  if (!page) return 0;
  const printableOrigin = axis === 'row' ? native.printable.y : native.printable.x;
  const printableExtent = axis === 'row' ? native.printable.height : native.printable.width;
  const origin = axis === 'row' ? page.originYPt : page.originXPt;
  const extent = axis === 'row' ? page.heightPt : page.widthPt;
  const titles = Math.max(0, origin - printableOrigin);
  return (Math.max(0, printableExtent - titles - extent) * CSS_PX_PER_POINT) / scale;
}

function nativePagination(
  state: State,
  sheet: number,
  setup: PageSetup,
  opts: PaginateOptions,
  native: PaginationResult,
): SheetPagination | null {
  const rowBands = nativeBands(native, 'row');
  const colBands = nativeBands(native, 'col');
  if (rowBands.length === 0 || colBands.length === 0) return null;

  const firstPage = native.pages[0];
  const origin = {
    row: firstPage?.firstRow ?? native.printArea[0]?.firstRow ?? 0,
    col: firstPage?.firstCol ?? native.printArea[0]?.firstCol ?? 0,
  };
  const content = {
    row: Math.max(...rowBands.map((band) => band.end)),
    col: Math.max(...colBands.map((band) => band.end)),
  };
  const scale = finitePositive(native.scale, 1);
  const pagePx = {
    width: finitePositive((native.printable.width * CSS_PX_PER_POINT) / scale, 1),
    height: finitePositive((native.printable.height * CSS_PX_PER_POINT) / scale, 1),
  };
  const marginPx = {
    top: Math.max(0, (native.margins.top * CSS_PX_PER_POINT) / scale),
    right: Math.max(0, (native.margins.right * CSS_PX_PER_POINT) / scale),
    bottom: Math.max(0, (native.margins.bottom * CSS_PX_PER_POINT) / scale),
    left: Math.max(0, (native.margins.left * CSS_PX_PER_POINT) / scale),
  };
  const rows = extendNativeBands(
    rowBands,
    'row',
    state,
    setup,
    content.row,
    opts.throughRow,
    pagePx.height,
    nativePageRoom(native, rowBands.at(-1)?.start ?? origin.row, 'row', scale),
  );
  const cols = extendNativeBands(
    colBands,
    'col',
    state,
    setup,
    content.col,
    opts.throughCol,
    pagePx.width,
    nativePageRoom(native, colBands.at(-1)?.start ?? origin.col, 'col', scale),
  );
  return {
    sheet,
    rowBands: rows,
    colBands: cols,
    origin,
    content,
    scale,
    pagePx,
    marginPx,
    pxPerInch: PRINT_PX_PER_INCH / scale,
    pageCount: native.pageCount,
    centerHorizontally: setup.centerHorizontally === true,
    centerVertically: setup.centerVertically === true,
    native,
  };
}

/** The engine whose native page breaks apply. Those follow the Windows print
 *  column model; a Mac sheet breaks on its own point widths, which the
 *  store's px already carry. */
function nativePaginationEngine(state: State, opts: PaginateOptions): WorkbookHandle | null {
  const engine = opts.engine ?? paginationEngines.get(state);
  if (engine?.capabilities.detailedPagination !== true) return null;
  return engine.geometryPlatform === 'mac' ? null : engine;
}

function paginate(state: State, sheet: number, opts: PaginateOptions): SheetPagination {
  const setup = getPageSetup(state, sheet);
  const engine = nativePaginationEngine(state, opts);
  if (engine) {
    const native = engine.paginateSheet(sheet);
    if (native) {
      const detailed =
        (opts.purpose === 'print' || nativePreviewCompatible(native)) &&
        nativePagination(state, sheet, setup, opts, native);
      if (detailed) return detailed;
    }
  }
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
    hiddenRowsOf(layout),
    layout.hiddenCols,
  );
  const pagePx = printablePagePixels(setup, scale);

  const rowBands = splitAxisIntoBands({
    from: origin.row,
    to: reach(opts.throughRow, content.row),
    budget: Math.max(layout.defaultRowHeight, pagePx.height),
    manualBreaks: setup.manualPageBreakRows,
    sizeOf: (row) =>
      isRowHidden(layout, row) ? 0 : (layout.rowHeights.get(row) ?? layout.defaultRowHeight),
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
    centerHorizontally: setup.centerHorizontally === true,
    centerVertically: setup.centerVertically === true,
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
  if (pagination.native) {
    const paper = pagination.native.paper;
    const scale = Math.max(pagination.scale, 0.01);
    return {
      width: (paper.widthPt * CSS_PX_PER_POINT) / scale,
      height: (paper.heightPt * CSS_PX_PER_POINT) / scale,
    };
  }
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
  // The native workbook can change independently of the store (for example
  // a host may write through the handle while calculation is manual). There
  // is no public native revision counter, so do not memoise this path and
  // risk reusing stale page placements.
  if (nativePaginationEngine(state, opts)) return paginate(state, sheet, opts);
  const engine = opts.engine ?? paginationEngines.get(state);
  const key = [
    sheet,
    state.layout,
    state.data.cells,
    state.format.formats,
    state.merges.byAnchor,
    engine,
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
 * render scale. The first page contributes only its own leading margin; every
 * page after it also carries the previous page's trailing margin and the grey
 * strip between the two.
 */
export function pageLayoutGaps(
  state: State,
  sheet: number,
  opts: PaginateOptions = {},
): PageLayoutGaps {
  const pagination = paginationFor(state, sheet, opts);
  const setup = getPageSetup(state, sheet);
  const zoom = renderScale(state.layout, state.viewport);
  const m = pagination.marginPx;
  const pageGapRows = new Map<number, number>();
  const pageGapCols = new Map<number, number>();
  const nativePaper = pagination.native ? pageBoxPx(pagination, setup) : null;
  const centeredNativeOrigin = (
    page: NonNullable<SheetPagination['native']>['pages'][number],
    axis: 'row' | 'col',
  ): number => {
    const native = pagination.native;
    if (!native) return axis === 'row' ? page.originYPt : page.originXPt;
    const origin = axis === 'row' ? page.originYPt : page.originXPt;
    const printableOrigin = axis === 'row' ? native.printable.y : native.printable.x;
    const pageExtent = axis === 'row' ? page.heightPt : page.widthPt;
    const printableExtent = axis === 'row' ? native.printable.height : native.printable.width;
    const centered = axis === 'row' ? pagination.centerVertically : pagination.centerHorizontally;
    if (!centered) return origin;
    const titleOffset = Math.max(0, origin - printableOrigin);
    return origin + Math.max(0, (printableExtent - titleOffset - pageExtent) / 2);
  };
  const nativeOrigin = (axis: 'row' | 'col', band: AxisBand): number | null => {
    if (!pagination.native) return null;
    const page = pagination.native.pages.find((candidate) =>
      axis === 'row' ? candidate.firstRow === band.start : candidate.firstCol === band.start,
    );
    const printableOrigin =
      axis === 'row' ? pagination.native.printable.y : pagination.native.printable.x;
    if (page) return centeredNativeOrigin(page, axis);
    // A native pagination result only covers the printed content. Preview
    // still extends bands to keep scrolling continuous; place those blank
    // bands after the last physical page by carrying forward that page's
    // effective cell-block origin. Using the printable origin here loses the
    // prior page's centering offset and makes the blank paper jump backward.
    const previous = pagination.native.pages
      .filter((candidate) =>
        axis === 'row' ? candidate.lastRow < band.start : candidate.lastCol < band.start,
      )
      .sort((a, b) => (axis === 'row' ? b.lastRow - a.lastRow : b.lastCol - a.lastCol))[0];
    return previous ? centeredNativeOrigin(previous, axis) : printableOrigin;
  };
  const axisExtent = (axis: 'row' | 'col', band: AxisBand): number => {
    const hidden = axis === 'row' ? hiddenRowsOf(state.layout) : state.layout.hiddenCols;
    const sizes = axis === 'row' ? state.layout.rowHeights : state.layout.colWidths;
    const fallback = axis === 'row' ? state.layout.defaultRowHeight : state.layout.defaultColWidth;
    let extent = 0;
    for (let index = band.start; index <= band.end; index += 1) {
      if (!hidden.has(index)) extent += sizes.get(index) ?? fallback;
    }
    return extent;
  };
  pagination.rowBands.forEach((band, i) => {
    const previous = pagination.rowBands[i - 1];
    const previousOrigin = previous ? nativeOrigin('row', previous) : null;
    const currentOrigin = nativeOrigin('row', band);
    const lead =
      previous === undefined
        ? m.top +
          ((currentOrigin === null ? 0 : currentOrigin - (pagination.native?.printable.y ?? 0)) *
            (96 / 72)) /
            pagination.scale
        : nativePaper
          ? Math.max(
              0,
              nativePaper.height +
                (previousOrigin !== null && currentOrigin !== null
                  ? ((currentOrigin - previousOrigin) * (96 / 72)) / pagination.scale
                  : 0) -
                axisExtent('row', previous),
            )
          : Math.max(
              m.bottom + m.top,
              pagination.pagePx.height + m.top + m.bottom - axisExtent('row', previous),
            );
    pageGapRows.set(band.start, lead * zoom + (i === 0 ? 0 : PAGE_GUTTER_PX));
  });
  pagination.colBands.forEach((band, i) => {
    const previous = pagination.colBands[i - 1];
    const previousOrigin = previous ? nativeOrigin('col', previous) : null;
    const currentOrigin = nativeOrigin('col', band);
    const lead =
      previous === undefined
        ? m.left +
          ((currentOrigin === null ? 0 : currentOrigin - (pagination.native?.printable.x ?? 0)) *
            (96 / 72)) /
            pagination.scale
        : nativePaper
          ? Math.max(
              0,
              nativePaper.width +
                (previousOrigin !== null && currentOrigin !== null
                  ? ((currentOrigin - previousOrigin) * (96 / 72)) / pagination.scale
                  : 0) -
                axisExtent('col', previous),
            )
          : Math.max(
              m.right + m.left,
              pagination.pagePx.width + m.left + m.right - axisExtent('col', previous),
            );
    pageGapCols.set(band.start, lead * zoom + (i === 0 ? 0 : PAGE_GUTTER_PX));
  });
  return { pageGapRows, pageGapCols };
}
