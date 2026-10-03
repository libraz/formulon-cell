/** Physical page geometry for printing and the page-layout views: paper sizes, margins, per-page axis bands, region tiling, and fit-to-pages scale. */
import type { PageMargins, PageSetup } from '../store/store.js';
import { type PrintAreaBounds, parsePrintTitleCols } from './print-ranges.js';

export function printColumnsForRegion(
  region: PrintAreaBounds,
  titleColRange: [number, number] | null,
): number[] {
  const cols: number[] = [];
  const seen = new Set<number>();
  const add = (col: number): void => {
    if (seen.has(col)) return;
    seen.add(col);
    cols.push(col);
  };
  if (titleColRange) {
    for (let c = titleColRange[0]; c <= titleColRange[1]; c += 1) add(c);
  }
  for (let c = region.col0; c <= region.col1; c += 1) add(c);
  return cols;
}

export const PAPER_DIMENSIONS: Record<string, string> = {
  A3: 'A3',
  A4: 'A4',
  A5: 'A5',
  letter: 'letter',
  legal: 'legal',
  tabloid: 'tabloid',
};

/** Physical paper dimensions in inches (portrait). Used to derive a
 *  fit-to-pages scale — the `@page size` keyword tells the browser the sheet
 *  but gives us no numbers to scale content against. */
const PAPER_INCHES: Record<string, { w: number; h: number }> = {
  A3: { w: 11.69, h: 16.54 },
  A4: { w: 8.27, h: 11.69 },
  A5: { w: 5.83, h: 8.27 },
  letter: { w: 8.5, h: 11 },
  legal: { w: 8.5, h: 14 },
  tabloid: { w: 11, h: 17 },
};

/** CSS reference pixel density — column widths / row heights are stored in px. */
export const PRINT_PX_PER_INCH = 96;

/** One page's worth of a single axis, as produced by {@link splitAxisIntoBands}. */
export interface AxisBand {
  /** First index on the page. */
  start: number;
  /** Last index on the page, inclusive. */
  end: number;
  /** The band opens at a user-inserted break rather than an automatic one.
   *  False for the first band, which starts because the content does. */
  manual: boolean;
}

export interface SplitAxisOptions {
  /** Inclusive index range to paginate. */
  from: number;
  to: number;
  /** Size of one index in unscaled sheet pixels; 0 for hidden. */
  sizeOf: (index: number) => number;
  /** Pixels of content one page can hold. */
  budget: number;
  /** Indices a page must start at. Values outside `[from, to]` are ignored. */
  manualBreaks?: readonly number[];
}

/**
 * Split one axis into per-page bands. A band closes when the next index would
 * overflow the page budget, or when a manual break forces a new page. A single
 * index wider than a whole page still gets its own band rather than looping
 * forever — the print CSS lets it overflow, matching the desktop app.
 */
export function splitAxisIntoBands(opts: SplitAxisOptions): AxisBand[] {
  const { from, to, sizeOf, budget } = opts;
  const breaks = new Set((opts.manualBreaks ?? []).filter((i) => i > from && i <= to));
  const bands: AxisBand[] = [];
  let index = from;
  while (index <= to) {
    const start = index;
    let used = 0;
    let end = index;
    while (end <= to) {
      if (end > start && breaks.has(end)) break;
      const next = sizeOf(end);
      if (end > start && used + next > budget) break;
      used += next;
      end += 1;
    }
    bands.push({ start, end: Math.max(start, end - 1), manual: breaks.has(start) });
    index = Math.max(start + 1, end);
  }
  return bands;
}

interface PrintLayoutMetrics {
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  defaultColWidth: number;
  defaultRowHeight: number;
}

export function effectivePrintMargins(setup: PageSetup): PageMargins {
  const printable = setup.printableBounds;
  if (!printable) return { ...setup.margins };
  return {
    top: Math.max(setup.margins.top, printable.top),
    right: Math.max(setup.margins.right, printable.right),
    bottom: Math.max(setup.margins.bottom, printable.bottom),
    left: Math.max(setup.margins.left, printable.left),
  };
}

export interface PrintableMarginAdjustment {
  side: keyof PageMargins;
  margin: number;
  minimum: number;
  effective: number;
}

export function printableMarginAdjustments(setup: PageSetup): PrintableMarginAdjustment[] {
  const printable = setup.printableBounds;
  if (!printable) return [];
  const sides: (keyof PageMargins)[] = ['top', 'right', 'bottom', 'left'];
  return sides
    .map((side) => ({
      side,
      margin: setup.margins[side],
      minimum: printable[side],
      effective: Math.max(setup.margins[side], printable[side]),
    }))
    .filter((item) => item.minimum > item.margin);
}

/** Natural size of the print regions in inches, from column widths / row
 *  heights. Width is the widest single region (regions break onto their own
 *  page); height is the stacked total. */
function printContentInches(
  regions: readonly PrintAreaBounds[],
  layout: PrintLayoutMetrics,
  hiddenRows: ReadonlySet<number>,
  hiddenCols: ReadonlySet<number>,
  titleColRange: [number, number] | null = null,
): { width: number; height: number } {
  let width = 0;
  let height = 0;
  for (const region of regions) {
    let regionW = 0;
    for (const c of printColumnsForRegion(region, titleColRange)) {
      if (hiddenCols.has(c)) continue;
      regionW += layout.colWidths.get(c) ?? layout.defaultColWidth;
    }
    let regionH = 0;
    for (let r = region.row0; r <= region.row1; r += 1) {
      if (hiddenRows.has(r)) continue;
      regionH += layout.rowHeights.get(r) ?? layout.defaultRowHeight;
    }
    width = Math.max(width, regionW);
    height += regionH;
  }
  return { width: width / PRINT_PX_PER_INCH, height: height / PRINT_PX_PER_INCH };
}

/** Physical page box in inches, orientation applied. */
export function paperInches(setup: PageSetup): { w: number; h: number } {
  const portrait = PAPER_INCHES[setup.paperSize] ?? { w: 8.27, h: 11.69 };
  return setup.orientation === 'landscape'
    ? { w: portrait.h, h: portrait.w }
    : { w: portrait.w, h: portrait.h };
}

/** Printable area of one page expressed in layout pixels — the budget a page
 *  of content has to fit into. Dividing by `scale` converts the physical area
 *  into unscaled sheet pixels, which is the space the splitter measures
 *  column widths and row heights against. */
export function printablePagePixels(
  setup: PageSetup,
  scale: number,
): { width: number; height: number } {
  const page = paperInches(setup);
  const m = effectivePrintMargins(setup);
  const factor = Math.max(scale, 0.1);
  return {
    width: (Math.max(0.1, page.w - m.left - m.right) * PRINT_PX_PER_INCH) / factor,
    height: (Math.max(0.1, page.h - m.top - m.bottom) * PRINT_PX_PER_INCH) / factor,
  };
}

function measuredColumnWidth(
  col: number,
  layout: PrintLayoutMetrics,
  hiddenCols: ReadonlySet<number>,
): number {
  return hiddenCols.has(col) ? 0 : (layout.colWidths.get(col) ?? layout.defaultColWidth);
}

function measuredRowHeight(
  row: number,
  layout: PrintLayoutMetrics,
  hiddenRows: ReadonlySet<number>,
): number {
  return hiddenRows.has(row) ? 0 : (layout.rowHeights.get(row) ?? layout.defaultRowHeight);
}

export function splitPrintRegionIntoTiles(
  region: PrintAreaBounds,
  setup: PageSetup,
  layout: PrintLayoutMetrics,
  hiddenRows: ReadonlySet<number>,
  hiddenCols: ReadonlySet<number>,
  titleRowRange: [number, number] | null,
  titleColRange: [number, number] | null,
  scale: number,
): PrintAreaBounds[] {
  const page = printablePagePixels(setup, scale);
  const titleWidth = titleColRange
    ? printColumnsForRegion(
        { row0: region.row0, row1: region.row0, col0: region.col0, col1: region.col0 - 1 },
        titleColRange,
      ).reduce((sum, col) => sum + measuredColumnWidth(col, layout, hiddenCols), 0)
    : 0;
  const titleHeight = titleRowRange
    ? Array.from(
        { length: titleRowRange[1] - titleRowRange[0] + 1 },
        (_, i) => titleRowRange[0] + i,
      )
        .filter((row) => row < region.row0 || row > region.row1)
        .reduce((sum, row) => sum + measuredRowHeight(row, layout, hiddenRows), 0)
    : 0;
  const colBudget = Math.max(layout.defaultColWidth, page.width - titleWidth);
  const rowBudget = Math.max(layout.defaultRowHeight, page.height - titleHeight);

  const colChunks = splitAxisIntoBands({
    from: region.col0,
    to: region.col1,
    budget: colBudget,
    manualBreaks: setup.manualPageBreakCols,
    sizeOf: (col) => measuredColumnWidth(col, layout, hiddenCols),
  });
  const rowChunks = splitAxisIntoBands({
    from: region.row0,
    to: region.row1,
    budget: rowBudget,
    manualBreaks: setup.manualPageBreakRows,
    // Print titles repeat on every page, so they cost nothing against the
    // budget of the band they happen to fall inside.
    sizeOf: (row) =>
      titleRowRange && row >= titleRowRange[0] && row <= titleRowRange[1]
        ? 0
        : measuredRowHeight(row, layout, hiddenRows),
  });

  const tiles: PrintAreaBounds[] = [];
  if (setup.pageOrder === 'overThenDown') {
    for (const rowBand of rowChunks) {
      for (const colBand of colChunks) {
        tiles.push({
          row0: rowBand.start,
          row1: rowBand.end,
          col0: colBand.start,
          col1: colBand.end,
        });
      }
    }
  } else {
    for (const colBand of colChunks) {
      for (const rowBand of rowChunks) {
        tiles.push({
          row0: rowBand.start,
          row1: rowBand.end,
          col0: colBand.start,
          col1: colBand.end,
        });
      }
    }
  }
  return tiles;
}

/**
 * Resolve the document scale. When Fit-to-pages is set (`fitWidth`/`fitHeight`)
 * the requested page count is honoured by estimating the content's natural size
 * against the printable page area and deriving the largest scale that fits —
 * Excel never scales *up*, so the result is capped at 100% and floored to whole
 * percent. With no fit constraint the explicit `scale` (default 100%) is used.
 */
export function computeFitToPagesScale(
  setup: PageSetup,
  regions: readonly PrintAreaBounds[],
  layout: PrintLayoutMetrics,
  hiddenRows: ReadonlySet<number>,
  hiddenCols: ReadonlySet<number>,
): number {
  const fitWidth = setup.fitWidth ?? 0;
  const fitHeight = setup.fitHeight ?? 0;
  if (fitWidth <= 0 && fitHeight <= 0) {
    return setup.scale && setup.scale > 0 ? setup.scale : 1;
  }
  const page = paperInches(setup);
  const m = effectivePrintMargins(setup);
  const printableW = Math.max(0.1, page.w - m.left - m.right);
  const printableH = Math.max(0.1, page.h - m.top - m.bottom);
  const content = printContentInches(
    regions,
    layout,
    hiddenRows,
    hiddenCols,
    parsePrintTitleCols(setup.printTitleCols),
  );

  const candidates: number[] = [];
  if (fitWidth > 0 && content.width > 0) candidates.push((fitWidth * printableW) / content.width);
  if (fitHeight > 0 && content.height > 0) {
    candidates.push((fitHeight * printableH) / content.height);
  }
  if (candidates.length === 0) return 1;
  const raw = Math.min(1, ...candidates);
  // Excel's minimum print scale is 10%; floor to whole percent so the content
  // never overflows the requested page count by a rounding sliver.
  return Math.max(0.1, Math.floor(raw * 100) / 100);
}
