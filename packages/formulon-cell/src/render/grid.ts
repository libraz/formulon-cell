import { isHeaderRow, tableForCell } from '../commands/format-as-table.js';
import { paginationFor, type SheetPagination } from '../commands/pagination.js';
import { formatA1FormulaAsR1C1 } from '../commands/refs.js';
import { addrKey } from '../engine/address.js';
import { evaluateCfFromEngine } from '../engine/cf-sync.js';
import { makeRangeResolver, type RangeResolver } from '../engine/range-resolver.js';
import { findSpillBlockers, findSpillRanges, looksLikeArrayFormula } from '../engine/spill.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { CellBorderSide, CellFormat, State } from '../store/store.js';
import { getPageSetup } from '../store/store.js';
import type { ResolvedTheme } from '../theme/resolve.js';
import { evaluateConditional } from './conditional.js';
import {
  type AxisLayout,
  buildColLayout,
  buildRowLayout,
  cellRectIn,
  colGap,
  colWidth,
  frozenColsWidth,
  frozenRowsHeight,
  gridOriginX,
  gridOriginY,
  isColVisible,
  isRowVisible,
  layoutForView,
  type Rect,
  rangeRects,
  rowGap,
  rowHeight,
  type ViewState,
} from './geometry.js';
import type { ChromePaintContext } from './grid/chrome-context.js';
import { paintDataBar } from './grid/data-bar.js';
import { paintHeaders } from './grid/headers.js';
import {
  detectErrorKind,
  detectValidationViolation,
  ERROR_TRIANGLE_COLOR,
  type ErrorTriangleHit,
  isPlainTextOverflowCandidate,
  normalizeFormatLocale,
  setErrorTriangleHits,
  setFillHandleRect,
  setValidationChevron,
  shouldShowValidationChevron,
  VALIDATION_TRIANGLE_COLOR,
} from './grid/hit-state.js';
import { paintFreezeDividers, paintGridLines } from './grid/lines.js';
import {
  paintPageBreakPreview,
  paintPageLayoutBackground,
  paintPageLayoutChrome,
  paintPageRulers,
} from './grid/page-view.js';
import { tableCellFormat } from './grid/table-format.js';
import { paintTraces } from './grid/traces.js';
import {
  CONDITIONAL_ICON_GUTTER,
  paintActiveCellOutline,
  paintCellBackground,
  paintCellBorders,
  paintCellFill,
  paintCellText,
  paintCommentMarker,
  paintConditionalIcon,
  paintCopyMarquee,
  paintErrorTriangle,
  paintFillHandle,
  paintFillPreview,
  paintLockMarker,
  paintRefHighlight,
  paintSpillBlocker,
  paintSpillOutline,
  paintTableHeaderChevron,
  paintValidationChevron,
  paintValidationCircle,
  paintValidationTriangle,
} from './painters.js';
import { paintCellSparkline } from './sparkline.js';

const inRange = (a: Addr, r: Range): boolean =>
  a.row >= r.r0 && a.row <= r.r1 && a.col >= r.c0 && a.col <= r.c1;

const mergeRangeAt = (state: State, addr: Addr): Range | null => {
  const key = addrKey(addr);
  const anchorKey = state.merges.byCell.get(key) ?? key;
  return state.merges.byAnchor.get(anchorKey) ?? null;
};

const borderSideSignature = (side: CellBorderSide | undefined): string => {
  if (side === undefined) return 'absent';
  if (typeof side === 'boolean') return `boolean:${side}`;
  return `style:${side.style};color:${side.color ?? ''}`;
};

/** Excel stores a uniform diagonal border on every cell covered by a merge,
 *  but renders that border once across the merged surface. Mixed diagonal
 *  metadata is discarded by Excel at merge time; suppress it here as well so
 *  stale per-cell data cannot produce crossing lines inside a merge. */
const mergedDiagonalBorders = (
  formats: ReadonlyMap<string, CellFormat>,
  merge: Range,
): CellFormat['borders'] => {
  let down: CellBorderSide | undefined;
  let downSignature: string | null = null;
  let up: CellBorderSide | undefined;
  let upSignature: string | null = null;
  for (let row = merge.r0; row <= merge.r1; row += 1) {
    for (let col = merge.c0; col <= merge.c1; col += 1) {
      const borders = formats.get(addrKey({ sheet: merge.sheet, row, col }))?.borders;
      const nextDown = borders?.diagonalDown;
      const nextUp = borders?.diagonalUp;
      const nextDownSignature = borderSideSignature(nextDown);
      const nextUpSignature = borderSideSignature(nextUp);
      if (downSignature === null) {
        down = nextDown;
        downSignature = nextDownSignature;
      } else if (downSignature !== 'mixed' && downSignature !== nextDownSignature) {
        downSignature = 'mixed';
      }
      if (upSignature === null) {
        up = nextUp;
        upSignature = nextUpSignature;
      } else if (upSignature !== 'mixed' && upSignature !== nextUpSignature) {
        upSignature = 'mixed';
      }
    }
  }
  const result: NonNullable<CellFormat['borders']> = {};
  if (downSignature !== null && downSignature !== 'mixed' && downSignature !== 'absent') {
    result.diagonalDown = down;
  }
  if (upSignature !== null && upSignature !== 'mixed' && upSignature !== 'absent') {
    result.diagonalUp = up;
  }
  return Object.keys(result).length > 0 ? result : undefined;
};

const sameRect = (a: Rect, b: Rect): boolean =>
  a.x === b.x && a.y === b.y && a.w === b.w && a.h === b.h;

const sameRectList = (a: readonly Rect[], b: readonly Rect[]): boolean =>
  a.length === b.length &&
  a.every((rect, index) => {
    const other = b[index];
    return other !== undefined && sameRect(rect, other);
  });

const rectsTouch = (a: Rect, b: Rect): boolean => {
  const eps = 1e-6;
  const overlapY = a.y < b.y + b.h - eps && b.y < a.y + a.h - eps;
  const overlapX = a.x < b.x + b.w - eps && b.x < a.x + a.w - eps;
  const sideBySide =
    (Math.abs(a.x + a.w - b.x) < eps || Math.abs(b.x + b.w - a.x) < eps) && overlapY;
  const stacked = (Math.abs(a.y + a.h - b.y) < eps || Math.abs(b.y + b.h - a.y) < eps) && overlapX;
  return (overlapX && overlapY) || sideBySide || stacked;
};

/** Join visible range rects that share an edge. Freeze quadrants are separate
 *  for fills and outlines, but their screen coordinates are contiguous, so a
 *  single text clip must span their union. Page-layout gutters remain
 *  separate components and therefore do not make text cross a paper gap. */
const connectedRects = (rects: readonly Rect[]): Rect[] => {
  const pending = rects.map((rect) => ({ ...rect }));
  let changed = true;
  while (changed) {
    changed = false;
    outer: for (let i = 0; i < pending.length; i += 1) {
      const first = pending[i];
      if (!first) continue;
      for (let j = i + 1; j < pending.length; j += 1) {
        const second = pending[j];
        if (!second || !rectsTouch(first, second)) continue;
        pending[i] = {
          x: Math.min(first.x, second.x),
          y: Math.min(first.y, second.y),
          w: Math.max(first.x + first.w, second.x + second.w) - Math.min(first.x, second.x),
          h: Math.max(first.y + first.h, second.y + second.h) - Math.min(first.y, second.y),
        };
        pending.splice(j, 1);
        changed = true;
        break outer;
      }
    }
  }
  return pending;
};

/** Pick the visible rectangle at a range's logical bottom/trailing corner.
 *  `rangeRects` returns freeze quadrants in logical row/column order, but the
 *  physical trailing edge moves to the left on a right-to-left sheet. */
const trailingRect = (rects: readonly Rect[], rtl: boolean): Rect | null => {
  if (rects.length === 0) return null;
  const bottom = Math.max(...rects.map((rect) => rect.y + rect.h));
  const bottomRects = rects.filter((rect) => rect.y + rect.h === bottom);
  return bottomRects.reduce(
    (best, rect) => {
      if (!best) return rect;
      if (rtl) return rect.x < best.x ? rect : best;
      return rect.x + rect.w > best.x + best.w ? rect : best;
    },
    null as Rect | null,
  );
};

const visibleCount = (
  pixels: number,
  start: number,
  max: number,
  sizeOf: (idx: number) => number,
): number => {
  if (pixels <= 0 || start >= max) return 1;
  let used = 0;
  let count = 0;
  const cap = Math.min(max - start, 2_000);
  while (count < cap && used < pixels) {
    used += Math.max(0, sizeOf(start + count));
    count += 1;
  }
  return Math.max(1, count + 1);
};

export interface RendererDeps {
  host: HTMLElement;
  canvas: HTMLCanvasElement;
  getState: () => State;
  getTheme: () => ResolvedTheme;
  onViewportSize?: (rowCount: number, colCount: number, widthPx: number) => void;
  /** Optional accessor for the active workbook. When supplied and the engine
   *  exposes `evaluateCfRange`, conditional-format rules loaded from the
   *  .xlsx are evaluated alongside the JS-side rule set and overlaid on top
   *  of the rendered cells. */
  getWb?: () => WorkbookHandle | null;
  /** UI/data-format locale. Short app ids like `ja` are accepted by Intl,
   *  but callers may return full tags like `ja-JP`. */
  getLocale?: () => string;
  /** UI strings. Only the page views paint text of their own — the header /
   *  footer placeholders and the page-number watermark. */
  getStrings?: () => Strings;
  /** Optional formatter pipeline — `inst.cells.resolveDisplay`. Returns
   *  the displayed string for matching cells, or null to fall through
   *  to the default text. */
  getDisplay?: (
    addr: { sheet: number; row: number; col: number },
    value: CellValue,
    formula: string | null,
    format: CellFormat | undefined,
  ) => string | null;
}

/**
 * Owns the Canvas. Schedules paints on next animation frame and coalesces
 * multiple `invalidate()` calls into one. The store and the engine never
 * touch the canvas directly — they call `invalidate()` and let this paint.
 */
export class GridRenderer {
  private readonly host: HTMLElement;

  private readonly canvas: HTMLCanvasElement;

  private readonly ctx: CanvasRenderingContext2D;

  private readonly getState: () => State;

  private readonly getTheme: () => ResolvedTheme;

  private readonly getWb: () => WorkbookHandle | null;

  private readonly getLocale: () => string;

  private readonly getStrings: () => Strings;

  private readonly getDisplay: RendererDeps['getDisplay'];

  private readonly onViewportSize: RendererDeps['onViewportSize'];

  private dpr = 1;

  private cssWidth = 0;

  private cssHeight = 0;

  private rafId = 0;

  private marqueeRafId = 0;

  private readonly sheetBackgroundImages = new Map<
    string,
    { image: HTMLImageElement; status: 'loading' | 'loaded' | 'error' }
  >();

  constructor(deps: RendererDeps) {
    this.host = deps.host;
    this.canvas = deps.canvas;
    const ctx = this.canvas.getContext('2d', { alpha: false });
    if (!ctx) throw new Error('formulon-cell: 2D canvas context unavailable');
    this.ctx = ctx;
    this.getState = deps.getState;
    this.getTheme = deps.getTheme;
    this.getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
    this.getLocale = deps.getLocale ?? ((): string => 'en-US');
    this.getStrings = deps.getStrings ?? ((): Strings => defaultStrings);
    this.getDisplay = deps.getDisplay;
    this.onViewportSize = deps.onViewportSize;
  }

  resize(): void {
    const rect = this.host.getBoundingClientRect();
    this.cssWidth = Math.max(0, rect.width);
    this.cssHeight = Math.max(0, rect.height);
    this.dpr = Math.max(1, Math.min(3, window.devicePixelRatio || 1));
    this.canvas.width = Math.round(this.cssWidth * this.dpr);
    this.canvas.height = Math.round(this.cssHeight * this.dpr);
    this.canvas.style.width = `${this.cssWidth}px`;
    this.canvas.style.height = `${this.cssHeight}px`;
    this.resizeViewport();
    this.invalidate();
  }

  private resizeViewport(): void {
    if (!this.onViewportSize) return;
    const state = this.getState();
    const layout = layoutForView(state);
    const { viewport } = state;
    const bodyH = Math.max(
      0,
      this.cssHeight - gridOriginY(layout) - frozenRowsHeight(layout, viewport),
    );
    const bodyW = Math.max(
      0,
      this.cssWidth - gridOriginX(layout) - frozenColsWidth(layout, viewport),
    );
    const firstRow = Math.max(
      viewport.rowStart,
      layout.freezeRows,
      viewport.navigationRange?.r0 ?? 0,
    );
    const firstCol = Math.max(
      viewport.colStart,
      layout.freezeCols,
      viewport.navigationRange?.c0 ?? 0,
    );
    const lastRow = (viewport.navigationRange?.r1 ?? 1_048_575) + 1;
    const lastCol = (viewport.navigationRange?.c1 ?? 16_383) + 1;
    // Page gutters take screen space without holding a cell, so they count
    // against the viewport budget too — otherwise Page Layout view would
    // under-fill the canvas by a page margin per page.
    const rowCount = visibleCount(
      bodyH,
      firstRow,
      lastRow,
      (idx) => rowHeight(layout, idx, viewport) + rowGap(layout, idx),
    );
    const colCount = visibleCount(
      bodyW,
      firstCol,
      lastCol,
      (idx) => colWidth(layout, idx, viewport) + colGap(layout, idx),
    );
    this.onViewportSize(rowCount, colCount, this.cssWidth);
  }

  invalidate(): void {
    if (this.cssWidth > 0 && this.cssHeight > 0) this.resizeViewport();
    if (this.rafId) return;
    this.rafId = requestAnimationFrame(() => {
      this.rafId = 0;
      this.paint();
    });
  }

  dispose(): void {
    if (this.rafId) cancelAnimationFrame(this.rafId);
    this.rafId = 0;
    if (this.marqueeRafId) cancelAnimationFrame(this.marqueeRafId);
    this.marqueeRafId = 0;
  }

  private paint(): void {
    if (this.cssWidth === 0 || this.cssHeight === 0) return;

    const baseState = this.getState();
    const state: ViewState = { ...baseState, layout: layoutForView(baseState) };
    const theme = this.getTheme();
    const ctx = this.ctx;

    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.imageSmoothingEnabled = false;

    ctx.fillStyle = theme.bg;
    ctx.fillRect(0, 0, this.cssWidth, this.cssHeight);

    // Build per-axis position caches once per paint. Sub-passes reuse them
    // for O(1) cellRect lookups.
    const cols = buildColLayout(state.layout, state.viewport);
    const rows = buildRowLayout(state.layout, state.viewport);

    // Page views need the same pagination the print pipeline uses; it is
    // memoised, so asking for it every paint costs a cache probe.
    const pagination = this.paginationForView(state);
    if (pagination && state.ui.workbookView === 'pageLayout') {
      paintPageLayoutBackground(this.chromeCtx(), state, theme, pagination, cols, rows);
    }

    this.paintSheetBackground(state);
    if (state.ui.showGridLines !== false) this.paintGridLines(state, theme, cols, rows);
    this.paintCells(state, theme, cols, rows);
    if (state.ui.showHeaders !== false) this.paintHeaders(state, theme, cols, rows);
    this.paintFreezeDividers(state, theme, cols, rows);
    this.paintBorders(state, theme, cols, rows);
    // Page chrome covers the gutters, so it has to follow every pass that
    // paints across the whole canvas — gridlines and freeze dividers do.
    if (pagination) this.paintPageChrome(state, theme, pagination, cols, rows);
    this.paintSpills(state, theme, cols, rows);
    this.paintActive(state, theme, cols, rows);
    this.paintEditorRefs(state);
    this.paintTraces(state, cols, rows);
  }

  /** Pagination for the active sheet, or null in Normal view. Reaches one
   *  viewport past the visible slice so the page opening just off screen is
   *  already laid out when the user scrolls onto it. */
  private paginationForView(state: ViewState): SheetPagination | null {
    if (state.ui.workbookView === 'normal') return null;
    const { viewport } = state;
    return paginationFor(this.getState(), state.data.sheetIndex, {
      throughRow: viewport.rowStart + viewport.rowCount,
      throughCol: viewport.colStart + viewport.colCount,
    });
  }

  private paintPageChrome(
    state: ViewState,
    theme: ResolvedTheme,
    pagination: SheetPagination,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const setup = getPageSetup(this.getState(), state.data.sheetIndex);
    const strings = this.getStrings();
    if (state.ui.workbookView === 'pageLayout') {
      paintPageLayoutChrome(
        this.chromeCtx(),
        state,
        theme,
        pagination,
        setup,
        { addHeader: strings.pageView.addHeader, addFooter: strings.pageView.addFooter },
        cols,
        rows,
      );
      paintPageRulers(
        this.chromeCtx(),
        state,
        theme,
        pagination,
        strings.pageView.rulerUnit,
        cols,
        rows,
      );
      return;
    }
    const template = strings.pageView.pageNumber;
    paintPageBreakPreview(
      this.chromeCtx(),
      state,
      theme,
      pagination,
      setup,
      (page) => template.replace('{n}', String(page)),
      cols,
      rows,
    );
  }

  private paintSheetBackground(state: ViewState): void {
    const url = state.ui.sheetBackgroundImages.get(state.data.sheetIndex);
    if (!url) return;
    let entry = this.sheetBackgroundImages.get(url);
    if (!entry) {
      const image = new Image();
      entry = { image, status: 'loading' };
      image.onload = () => {
        const current = this.sheetBackgroundImages.get(url);
        if (!current) return;
        current.status = 'loaded';
        this.invalidate();
      };
      image.onerror = () => {
        const current = this.sheetBackgroundImages.get(url);
        if (!current) return;
        current.status = 'error';
      };
      image.src = url;
      this.sheetBackgroundImages.set(url, entry);
    }
    if (
      entry.status !== 'loaded' ||
      entry.image.naturalWidth <= 0 ||
      entry.image.naturalHeight <= 0
    ) {
      return;
    }

    const ox = gridOriginX(state.layout);
    const oy = gridOriginY(state.layout);
    const ctx = this.ctx;
    ctx.save();
    ctx.imageSmoothingEnabled = true;
    const pattern = ctx.createPattern(entry.image, 'repeat');
    if (pattern) {
      ctx.fillStyle = pattern;
      ctx.translate(state.layout.rtl ? 0 : ox, oy);
      ctx.fillRect(0, 0, this.cssWidth - ox, this.cssHeight - oy);
    }
    ctx.restore();
  }

  private paintEditorRefs(state: ViewState): void {
    const refs = state.ui.editorRefs;
    if (!refs || refs.length === 0) return;
    const sheet = state.data.sheetIndex;
    const ctx = this.ctx;
    for (const ref of refs) {
      const range: Range = {
        sheet,
        r0: ref.r0,
        c0: ref.c0,
        r1: ref.r1,
        c1: ref.c1,
      };
      const rects = rangeRects(state.layout, state.viewport, range);
      // The bounding box of the entire range — paint as one outline rather
      //  than per-cell so 2x2 selections look like a single bordered box.
      if (rects.length === 0) continue;
      let x0 = Number.POSITIVE_INFINITY;
      let y0 = Number.POSITIVE_INFINITY;
      let x1 = Number.NEGATIVE_INFINITY;
      let y1 = Number.NEGATIVE_INFINITY;
      for (const r of rects) {
        x0 = Math.min(x0, r.x);
        y0 = Math.min(y0, r.y);
        x1 = Math.max(x1, r.x + r.w);
        y1 = Math.max(y1, r.y + r.h);
      }
      paintRefHighlight(ctx, { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }, ref.colorIndex);
    }
  }

  private chromeCtx(): ChromePaintContext {
    return { ctx: this.ctx, dpr: this.dpr, cssWidth: this.cssWidth, cssHeight: this.cssHeight };
  }

  private paintGridLines(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    paintGridLines(this.chromeCtx(), state, theme, cols, rows);
  }

  private paintCells(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const ctx = this.ctx;
    const {
      layout,
      data,
      selection,
      format,
      merges,
      sparkline,
      errorIndicators,
      protection,
      tables,
    } = state;
    const rtl = layout.rtl;
    const active = selection.active;
    const conditional = evaluateConditional(state);
    const sparklines = sparkline.sparklines;
    const ignored = errorIndicators.ignoredErrors;
    const validationCircles = errorIndicators.validationCircles;
    // Lock-icon overlay only fires when the active sheet is currently
    // protected. Pre-compute the flag so the per-cell loop can skip the
    // Map lookup on every iteration.
    const sheetProtected = protection.protectedSheets.has(data.sheetIndex);
    // RangeResolver is only needed for list-validation re-resolution. We
    // build it lazily — most viewport repaints don't have a single DV cell,
    // so allocating the resolver up-front would be wasted work.
    const wbForResolver = this.getWb();
    let resolver: RangeResolver | undefined;
    const getResolver = (): RangeResolver | undefined => {
      if (resolver) return resolver;
      if (!wbForResolver) return undefined;
      resolver = makeRangeResolver(wbForResolver, data.sheetIndex);
      return resolver;
    };
    const triangleHits: ErrorTriangleHit[] = [];
    const locale = normalizeFormatLocale(this.getLocale());
    // Engine-side CF (rules loaded from .xlsx). Merged on top — engine rules
    // currently win per field for cells with overlap. Restricted to the
    // visible viewport rect so we don't pay for off-screen cells.
    const wb = this.getWb();
    if (wb?.capabilities.conditionalFormat) {
      const sheet = state.data.sheetIndex;
      const vp = state.viewport;
      const r0 = vp.rowStart;
      const r1 = Math.max(r0, vp.rowStart + vp.rowCount - 1);
      const c0 = vp.colStart;
      const c1 = Math.max(c0, vp.colStart + vp.colCount - 1);
      const engineCf = evaluateCfFromEngine(wb, sheet, r0, c0, r1, c1);
      for (const [k, v] of engineCf) {
        const merged = { ...(conditional.get(k) ?? {}), ...v };
        conditional.set(k, merged);
      }
    }

    const rng = selection.range;
    if (rng.r0 !== rng.r1 || rng.c0 !== rng.c1) {
      ctx.fillStyle = theme.accentSoft;
      const rects = rangeRects(layout, state.viewport, rng);
      for (const r of rects) ctx.fillRect(r.x, r.y, r.w, r.h);
    }
    // Disjoint multi-range selection (Ctrl/Cmd+click). Paint each extra band
    // with the same accent so the user can read all members at a glance.
    const extras = selection.extraRanges;
    if (extras && extras.length > 0) {
      ctx.fillStyle = theme.accentSoft;
      for (const er of extras) {
        if (er.r0 > er.r1 || er.c0 > er.c1) continue;
        for (const r of rangeRects(layout, state.viewport, er)) {
          ctx.fillRect(r.x, r.y, r.w, r.h);
        }
      }
    }

    // Compute merged cell bounds from the same freeze/RTL-aware projection
    // used by selection outlines. The old local width/height summation mixed
    // logical and screen coordinates when a merge crossed a frozen boundary.
    const mergeRectsAt = (row: number, col: number): Rect[] => {
      const range = merges.byAnchor.get(addrKey({ sheet: data.sheetIndex, row, col }));
      return range ? rangeRects(layout, state.viewport, range) : [];
    };
    const mergeBounds = (row: number, col: number): Rect | null =>
      connectedRects(mergeRectsAt(row, col))[0] ?? null;

    // The sheet background is painted before grid lines. Repaint it inside a
    // merge before its content so internal gridlines cannot show through. If
    // the background image is already loaded, restore the same repeat pattern
    // instead of replacing it with an opaque theme fill.
    const paintMergeBase = (bounds: Rect): void => {
      const url = state.ui.sheetBackgroundImages.get(data.sheetIndex);
      const image = url ? this.sheetBackgroundImages.get(url) : undefined;
      const pattern =
        image?.status === 'loaded' &&
        image.image.naturalWidth > 0 &&
        image.image.naturalHeight > 0 &&
        typeof ctx.createPattern === 'function'
          ? ctx.createPattern(image.image, 'repeat')
          : null;
      if (pattern) {
        const ox = layout.rtl ? 0 : gridOriginX(layout);
        const oy = gridOriginY(layout);
        ctx.save();
        ctx.translate(ox, oy);
        ctx.fillStyle = pattern;
        ctx.fillRect(bounds.x - ox, bounds.y - oy, bounds.w, bounds.h);
        ctx.restore();
        return;
      }
      ctx.fillStyle = theme.bg;
      ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
    };
    const paintMergeSurface = (
      bounds: Rect,
      isActive: boolean,
      isSelected: boolean | undefined,
    ): void => {
      paintMergeBase(bounds);
      if (isActive) {
        paintCellBackground({
          ctx,
          theme,
          bounds,
          value: { kind: 'blank' },
          formula: null,
          isActive: true,
          isInRange: false,
        });
      } else if (isSelected) {
        paintCellBackground({
          ctx,
          theme,
          bounds,
          value: { kind: 'blank' },
          formula: null,
          isActive: false,
          isInRange: true,
        });
      }
    };
    // Overflow and centre-across both spill into the following columns, which
    // the mirror puts to the left of the anchor on a right-to-left sheet — so
    // the widened rect has to move its left edge as well as grow.
    const spillRect = (base: Rect, extra: number): Rect =>
      layout.rtl
        ? { ...base, x: base.x - extra, w: base.w + extra }
        : { ...base, w: base.w + extra };
    const overflowBounds = (row: number, col: number, base: Rect): Rect => {
      let w = 0;
      for (let nextCol = col + 1; cols.positionAt.has(nextCol); nextCol += 1) {
        const nextKey = addrKey({ sheet: data.sheetIndex, row, col: nextCol });
        const nextCell = data.cells.get(nextKey);
        const nextFmt = format.formats.get(nextKey);
        const nextTable = tableForCell(tables.tables, data.sheetIndex, row, nextCol);
        if (
          merges.byCell.has(nextKey) ||
          merges.byAnchor.has(nextKey) ||
          nextTable ||
          nextFmt?.fill ||
          nextFmt?.borders ||
          (nextCell && (nextCell.formula || nextCell.value.kind !== 'blank'))
        ) {
          break;
        }
        w += cols.sizeAt.get(nextCol) ?? 0;
      }
      return w === 0 ? base : spillRect(base, w);
    };
    const centerContinuousBounds = (row: number, col: number, base: Rect): Rect => {
      let extra = 0;
      for (let nextCol = col + 1; cols.positionAt.has(nextCol); nextCol += 1) {
        const nextKey = addrKey({ sheet: data.sheetIndex, row, col: nextCol });
        const nextCell = data.cells.get(nextKey);
        const nextFmt = format.formats.get(nextKey);
        const nextTable = tableForCell(tables.tables, data.sheetIndex, row, nextCol);
        if (
          nextTable ||
          merges.byCell.has(nextKey) ||
          merges.byAnchor.has(nextKey) ||
          nextFmt?.align !== 'centerContinuous' ||
          (nextCell && (nextCell.formula || nextCell.value.kind !== 'blank'))
        ) {
          break;
        }
        extra += cols.sizeAt.get(nextCol) ?? 0;
      }
      return extra === 0 ? base : spillRect(base, extra);
    };

    // The anchor can be scrolled or hidden while a visible body segment of the
    // merge remains on screen. Clear that segment here; the normal anchor walk
    // below handles the common visible-anchor case and then paints content.
    for (const [anchorKey, merge] of merges.byAnchor) {
      if (merge.sheet !== data.sheetIndex) continue;
      const rects = rangeRects(layout, state.viewport, merge);
      const anchorVisible = rows.positionAt.has(merge.r0) && cols.positionAt.has(merge.c0);
      if (rects.length === 0 || anchorVisible) continue;
      const anchor = { sheet: merge.sheet, row: merge.r0, col: merge.c0 };
      const selected =
        inRange(anchor, rng) || extras?.some((extra) => inRange(anchor, extra)) === true;
      const isActive =
        active.sheet === merge.sheet && active.row === merge.r0 && active.col === merge.c0;
      const anchorFormat = format.formats.get(anchorKey);
      for (const rect of rects) {
        paintMergeSurface(rect, isActive, selected);
        if (anchorFormat?.fill || anchorFormat?.fillPattern) {
          paintCellFill({
            ctx,
            theme,
            bounds: rect,
            value: { kind: 'blank' },
            formula: null,
            isActive,
            isInRange: selected,
            format: anchorFormat,
          });
        }
      }
    }

    // Single visible-cells walk paints both static format fills (for blank
    // formatted cells too) and cell content. Iterating format.formats here
    // would be O(formats), which dominates on sheets with thousands of
    // formatted cells; a viewport-sized grid is bounded.
    for (const r of rows.visible) {
      for (const c of cols.visible) {
        const key = addrKey({ sheet: data.sheetIndex, row: r, col: c });
        // Skip cells hidden inside a merge — only the anchor paints.
        if (merges.byCell.has(key)) continue;
        const cell = data.cells.get(key);
        const fmt = format.formats.get(key);
        const table = tableForCell(tables.tables, data.sheetIndex, r, c);
        const tableFmt = table ? tableCellFormat(table, r, c) : undefined;
        const isMergeAnchor = merges.byAnchor.has(key);
        // Use the union of contiguous visible rectangles for content bounds,
        // but retain every visible quadrant for the merged surface and static
        // format fill. A merge crossing a frozen row/column is projected into
        // more than one rect; painting only one leaves the other quadrants
        // with a different background or exposed gridline.
        const mergedRects = isMergeAnchor ? mergeRectsAt(r, c) : [];
        const spark = sparklines.get(key);
        const isActive = r === active.row && c === active.col;
        const hasValidationCircle = validationCircles.has(key);
        // Render-worthy when there's data, a merge anchor, a static fill, or a
        // sparkline/table host — all reasons to paint into an otherwise blank cell.
        if (
          !cell &&
          !isActive &&
          !isMergeAnchor &&
          !fmt?.fill &&
          !fmt?.fillPattern &&
          !tableFmt?.fill &&
          !tableFmt?.fillPattern &&
          !spark &&
          !hasValidationCircle
        )
          continue;
        const bounds = mergeBounds(r, c) ?? cellRectIn(layout, cols, rows, r, c);
        const isInRange = inRange({ sheet: data.sheetIndex, row: r, col: c }, rng);
        const isInExtraRange = extras?.some((extra) =>
          inRange({ sheet: data.sheetIndex, row: r, col: c }, extra),
        );

        const overlay = conditional.get(key);
        const baseFmt: typeof fmt =
          tableFmt || fmt
            ? {
                ...tableFmt,
                ...fmt,
              }
            : undefined;
        const effectiveFmt: typeof fmt =
          overlay && (overlay.fill || overlay.color || overlay.bold || overlay.italic)
            ? {
                ...baseFmt,
                fill: overlay.fill ?? baseFmt?.fill,
                color: overlay.color ?? baseFmt?.color,
                bold: overlay.bold || baseFmt?.bold,
                italic: overlay.italic || baseFmt?.italic,
                underline: overlay.underline || baseFmt?.underline,
                strike: overlay.strike || baseFmt?.strike,
              }
            : baseFmt;

        if (isMergeAnchor) {
          for (const mergedRect of mergedRects) {
            paintMergeSurface(mergedRect, isActive, isInRange || isInExtraRange);
          }
        }

        const value: CellValue = cell?.value ?? { kind: 'blank' };
        const storedFormula = cell?.formula ?? null;
        const formula =
          storedFormula && state.ui.showFormulas === true && state.ui.r1c1 === true
            ? formatA1FormulaAsR1C1(storedFormula, { row: r, col: c })
            : storedFormula;
        const displayOverride =
          this.getDisplay?.(
            { sheet: data.sheetIndex, row: r, col: c },
            value,
            storedFormula,
            fmt,
          ) ?? null;
        const paintCtx = {
          ctx,
          theme,
          bounds,
          value,
          formula,
          isActive,
          isInRange,
          format: effectiveFmt,
          showFormulas: state.ui.showFormulas === true,
          showZeros: state.ui.showZeros !== false,
          displayOverride,
          locale,
          rtl,
        };

        // Static format fill OR overlay fill — both painted via paintCellFill,
        // which reads `format.fill`. The merged effectiveFmt already prefers
        // overlay over static, so a single call yields the right result.
        if (effectiveFmt?.fill || effectiveFmt?.fillPattern) {
          if (isMergeAnchor && mergedRects.length > 0) {
            for (const mergedRect of mergedRects) {
              paintCellFill({ ...paintCtx, bounds: mergedRect });
            }
          } else {
            paintCellFill(paintCtx);
          }
        }
        if (isActive && !effectiveFmt?.fill) paintCellBackground(paintCtx);
        if (overlay) paintDataBar(ctx, bounds, overlay);
        const hideConditionalValue =
          overlay?.showValue === false &&
          (overlay.bar !== undefined || (overlay.iconKind && overlay.iconSlot !== undefined));
        // Icon-set: paint the glyph in the leading gutter and inset the text by
        // the gutter width so the value reads cleanly next to the icon.
        const tableHeader = table ? isHeaderRow(table, r, c) : false;
        if (overlay?.iconKind && overlay.iconSlot !== undefined) {
          paintConditionalIcon(ctx, bounds, overlay.iconKind, overlay.iconSlot, rtl);
          if (!hideConditionalValue) {
            const insetBounds = {
              x: rtl ? bounds.x : bounds.x + CONDITIONAL_ICON_GUTTER,
              y: bounds.y,
              w: bounds.w - CONDITIONAL_ICON_GUTTER,
              h: bounds.h,
            };
            paintCellText({ ...paintCtx, bounds: insetBounds });
          }
        } else if (hideConditionalValue) {
          // Conditional formatting's "Show Bar/Icon Only" suppresses the
          // value text while still leaving fills, bars, and icons visible.
        } else if (tableHeader) {
          paintCellText({
            ...paintCtx,
            bounds: {
              ...bounds,
              x: rtl ? bounds.x + 20 : bounds.x,
              w: Math.max(0, bounds.w - 20),
            },
          });
        } else if (effectiveFmt?.align === 'centerContinuous' && !isMergeAnchor) {
          paintCellText({ ...paintCtx, bounds: centerContinuousBounds(r, c, bounds) });
        } else if (
          isPlainTextOverflowCandidate({
            value,
            formula,
            format: effectiveFmt,
            showFormulas: state.ui.showFormulas === true,
            displayOverride,
            tableHeader,
            hasIcon: false,
            isMergeAnchor,
          })
        ) {
          paintCellText({ ...paintCtx, bounds: overflowBounds(r, c, bounds) });
        } else {
          paintCellText(paintCtx);
        }
        if (tableHeader) paintTableHeaderChevron(ctx, bounds, theme, rtl);
        if (spark) paintCellSparkline(ctx, bounds, spark, state, this.getWb());
        if (fmt?.comment) paintCommentMarker(ctx, bounds, rtl);
        if (hasValidationCircle) paintValidationCircle(ctx, bounds, VALIDATION_TRIANGLE_COLOR);
        // Lock-icon overlay — only when the sheet is protected AND the cell
        // is explicitly unlocked, signalling which cells the user can still
        // type into despite the protection flag.
        if (sheetProtected && fmt?.locked === false) paintLockMarker(ctx, bounds, theme, rtl);

        // Error / validation triangles. Error wins over validation when both
        // would apply (an error-kind value already implies the data is bad —
        // the green triangle conveys that without piling on a red one). The
        // ignoredErrors set suppresses both kinds for the cell once the user
        // dismisses the popover via the "Ignore" action.
        const cellAddr = { sheet: data.sheetIndex, row: r, col: c };
        const cellKey = key;
        if (!ignored.has(cellKey)) {
          if (detectErrorKind(value)) {
            const hit = paintErrorTriangle(ctx, bounds, ERROR_TRIANGLE_COLOR, rtl);
            triangleHits.push({ rect: hit, addr: cellAddr, kind: 'error' });
          } else if (
            fmt?.validation &&
            detectValidationViolation(value, fmt.validation, getResolver())
          ) {
            const hit = paintValidationTriangle(ctx, bounds, VALIDATION_TRIANGLE_COLOR, rtl);
            triangleHits.push({ rect: hit, addr: cellAddr, kind: 'validation' });
          }
        }
      }
    }
    setErrorTriangleHits(triangleHits);
  }

  private paintBorders(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const ctx = this.ctx;
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

  private paintHeaders(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    paintHeaders(this.chromeCtx(), state, theme, cols, rows);
  }

  private paintFreezeDividers(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    paintFreezeDividers(this.chromeCtx(), state, theme, cols, rows);
  }

  private paintSpills(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const { layout, viewport, data } = state;
    // Prefer the engine's authoritative spill list when available; fall back
    // to the JS heuristic for stub mode or older engine package builds.
    const wb = this.getWb();
    const ranges = wb?.spillRanges(data.sheetIndex) ?? findSpillRanges(data.cells, data.sheetIndex);
    for (const r of ranges) {
      for (const rect of rangeRects(layout, viewport, r)) {
        paintSpillOutline(this.ctx, rect, theme);
      }
    }

    // #SPILL! obstruction outline. The engine's `spillInfo` reports the
    // attempted shape even when the result couldn't materialise; we use
    // that to figure out which cells are blocking and ring them in red.
    if (!wb?.spillInfo) return;
    for (const [key, cell] of data.cells) {
      if (!cell.formula || cell.value.kind !== 'error') continue;
      if (!looksLikeArrayFormula(cell.formula)) continue;
      const [sStr, rStr, cStr] = key.split(':');
      if (sStr === undefined || rStr === undefined || cStr === undefined) continue;
      if (Number.parseInt(sStr, 10) !== data.sheetIndex) continue;
      const row = Number.parseInt(rStr, 10);
      const col = Number.parseInt(cStr, 10);
      const info = wb.spillInfo(data.sheetIndex, row, col);
      if (!info) continue;
      const target: Range = {
        sheet: data.sheetIndex,
        r0: info.anchorRow,
        c0: info.anchorCol,
        r1: info.anchorRow + info.rows - 1,
        c1: info.anchorCol + info.cols - 1,
      };
      const blockers = findSpillBlockers(data.cells, data.sheetIndex, target);
      for (const b of blockers) {
        if (!isRowVisible(layout, viewport, b.row)) continue;
        if (!isColVisible(layout, viewport, b.col)) continue;
        const rect = cellRectIn(layout, cols, rows, b.row, b.col);
        paintSpillBlocker(this.ctx, rect);
      }
    }
  }

  private paintActive(
    state: ViewState,
    theme: ResolvedTheme,
    _cols: AxisLayout,
    _rows: AxisLayout,
  ): void {
    const { layout, viewport, selection, ui, format, data } = state;
    const a = selection.active;
    const r = selection.range;
    setFillHandleRect(null);
    setValidationChevron(null);

    const activeMerge = mergeRangeAt(state, a);
    const activeVisualRange: Range = activeMerge ?? {
      sheet: a.sheet,
      r0: a.row,
      c0: a.col,
      r1: a.row,
      c1: a.col,
    };
    const activeRects = rangeRects(layout, viewport, activeVisualRange);
    for (const bounds of activeRects) {
      paintActiveCellOutline(this.ctx, bounds, theme);
    }

    const activeAnchor = activeMerge
      ? { sheet: activeMerge.sheet, row: activeMerge.r0, col: activeMerge.c0 }
      : a;
    const fmt = format.formats.get(addrKey(activeAnchor));
    const validationBounds = trailingRect(activeRects, layout.rtl);
    if (validationBounds && shouldShowValidationChevron(fmt?.validation)) {
      const rect = paintValidationChevron(this.ctx, validationBounds, theme, layout.rtl);
      setValidationChevron({ rect, row: activeAnchor.row, col: activeAnchor.col });
    }

    if (r.r0 !== r.r1 || r.c0 !== r.c1) {
      const rects = rangeRects(layout, viewport, r);
      // A merged selection already received its full-range active outline
      // above. Painting the generic range outline over the same rectangle
      // creates a darker double border, unlike desktop spreadsheets.
      if (!sameRectList(activeRects, rects)) {
        const ctx = this.ctx;
        ctx.save();
        ctx.strokeStyle = theme.accent;
        ctx.lineWidth = 2;
        ctx.setLineDash([]);
        for (const rect of rects) {
          ctx.strokeRect(rect.x + 1, rect.y + 1, Math.max(0, rect.w - 2), Math.max(0, rect.h - 2));
        }
        ctx.restore();
      }
    }

    const preview = ui.fillPreview;
    if (preview) {
      const rects = rangeRects(layout, viewport, preview);
      for (const r of rects) paintFillPreview(this.ctx, r, theme);
    }
    const copyRanges = ui.copyRanges && ui.copyRanges.length > 0 ? ui.copyRanges : null;
    const rangesToPaint = (copyRanges ?? (ui.copyRange ? [ui.copyRange] : [])).filter(
      (copyRange) => copyRange.sheet === data.sheetIndex,
    );
    if (rangesToPaint.length > 0) {
      const phase = (performance.now() / 180) % 7;
      for (const copyRange of rangesToPaint) {
        const rects = rangeRects(layout, viewport, copyRange);
        for (const r of rects) paintCopyMarquee(this.ctx, r, phase);
      }
      // Keep the marquee animating, but track the handle so dispose() can cancel
      // it — otherwise the self-scheduling loop pins the renderer after teardown.
      if (!this.marqueeRafId) {
        this.marqueeRafId = requestAnimationFrame(() => {
          this.marqueeRafId = 0;
          this.invalidate();
        });
      }
    }

    const handleRange = activeMerge && r.r0 === r.r1 && r.c0 === r.c1 ? activeVisualRange : r;
    const handleBounds = trailingRect(rangeRects(layout, viewport, handleRange), layout.rtl);
    if (handleBounds) {
      setFillHandleRect(paintFillHandle(this.ctx, handleBounds, theme, layout.rtl));
    }
  }

  private paintTraces(state: ViewState, cols: AxisLayout, rows: AxisLayout): void {
    paintTraces(this.ctx, state, cols, rows);
  }
}

export {
  detectErrorKind,
  detectValidationViolation,
  ERROR_TRIANGLE_COLOR,
  type ErrorTriangleHit,
  type ErrorTriangleKind,
  getErrorTriangleHits,
  getFillHandleRect,
  getOutlineToggleHits,
  getValidationChevron,
  isPlainTextOverflowCandidate,
  VALIDATION_TRIANGLE_COLOR,
} from './grid/hit-state.js';
