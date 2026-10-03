import { paginationFor, type SheetPagination } from '../commands/pagination.js';
import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import { findSpillBlockers, findSpillRanges, looksLikeArrayFormula } from '../engine/spill.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { State } from '../store/store.js';
import { getPageSetup } from '../store/store.js';
import type { ResolvedTheme } from '../theme/resolve.js';
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
  rangeRects,
  rowGap,
  rowHeight,
  type ViewState,
} from './geometry.js';
import { paintBorders } from './grid/borders.js';
import { type CellDisplayResolver, type CellPaintContext, paintCells } from './grid/cells.js';
import type { ChromePaintContext } from './grid/chrome-context.js';
import { paintHeaders } from './grid/headers.js';
import {
  setFillHandleRect,
  setValidationChevron,
  shouldShowValidationChevron,
} from './grid/hit-state.js';
import { paintFreezeDividers, paintGridLines } from './grid/lines.js';
import { mergeRangeAt } from './grid/merged-cells.js';
import {
  paintPageBreakPreview,
  paintPageLayoutBackground,
  paintPageLayoutChrome,
  paintPageRulers,
} from './grid/page-view.js';
import { sameRectList, trailingRect } from './grid/range-rects.js';
import { paintTraces } from './grid/traces.js';
import {
  paintActiveCellOutline,
  paintCopyMarquee,
  paintFillHandle,
  paintFillPreview,
  paintRefHighlight,
  paintSpillBlocker,
  paintSpillOutline,
  paintValidationChevron,
} from './painters.js';

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
  getDisplay?: CellDisplayResolver;
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
    const lastRow = (viewport.navigationRange?.r1 ?? MAX_ROW) + 1;
    const lastCol = (viewport.navigationRange?.c1 ?? MAX_COL) + 1;
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

  private cellCtx(state: ViewState): CellPaintContext {
    const url = state.ui.sheetBackgroundImages.get(state.data.sheetIndex);
    const entry = url ? this.sheetBackgroundImages.get(url) : undefined;
    return {
      ctx: this.ctx,
      wb: this.getWb(),
      locale: this.getLocale(),
      getDisplay: this.getDisplay,
      sheetBackgroundImage: entry?.status === 'loaded' ? entry.image : null,
    };
  }

  private paintCells(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    paintCells(this.cellCtx(state), state, theme, cols, rows);
  }

  private paintBorders(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    paintBorders(this.ctx, state, theme, cols, rows);
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
