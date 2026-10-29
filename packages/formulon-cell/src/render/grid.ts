import { totalsCellAt } from '../commands/format-as-table.js';
import { paginationFor, type SheetPagination } from '../commands/pagination.js';
import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import { findSpillRanges, looksLikeArrayFormula } from '../engine/spill.js';
import type { CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { frozenColEnd, frozenRowEnd } from '../store/freeze-band.js';
import type { State } from '../store/store.js';
import { getPageSetup } from '../store/store.js';
import type { ResolvedTheme } from '../theme/resolve.js';
import { createConditionalCache } from './conditional.js';
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
  layoutForView,
  type Rect,
  rangeRects,
  rowGap,
  rowHeight,
  type ViewState,
} from './geometry.js';
import { paintBorders } from './grid/borders.js';
import { type CellDisplayResolver, type CellPaintContext, paintCells } from './grid/cells.js';
import type { ChromePaintContext } from './grid/chrome-context.js';
import { paintHeaders } from './grid/headers.js';
import { type GridHitState, gridHitState, shouldShowValidationChevron } from './grid/hit-state.js';
import { paintFreezeDividers, paintGridLines } from './grid/lines.js';
import { mergeRangeAt } from './grid/merged-cells.js';
import { paintVisibleNotes } from './grid/notes.js';
import {
  paintPageBreakPreview,
  paintPageLayoutBackground,
  paintPageLayoutChrome,
  paintPageRulers,
} from './grid/page-view.js';
import { sameRectList, trailingRect } from './grid/range-rects.js';
import { paintTraces } from './grid/traces.js';
import {
  paintChartSourceHighlight,
  paintCopyMarquee,
  paintFillHandle,
  paintFillPreview,
  paintRefHighlight,
  paintSelectionOutline,
  paintSpillBlocker,
  paintSpillOutline,
  paintValidationChevron,
} from './painters.js';

type SpillCandidateCells = ReadonlyMap<string, { value: CellValue; formula: string | null }>;
type SpillErrorCandidate = { row: number; col: number };

const spillErrorCandidateCache = new WeakMap<
  SpillCandidateCells,
  Map<number, readonly SpillErrorCandidate[]>
>();

function spillErrorCandidates(
  cells: SpillCandidateCells,
  sheet: number,
): readonly SpillErrorCandidate[] {
  let bySheet = spillErrorCandidateCache.get(cells);
  if (!bySheet) {
    bySheet = new Map();
    spillErrorCandidateCache.set(cells, bySheet);
  }
  const cached = bySheet.get(sheet);
  if (cached) return cached;

  const candidates: SpillErrorCandidate[] = [];
  for (const [key, cell] of cells) {
    if (cell.value.kind !== 'error' || !cell.formula || !looksLikeArrayFormula(cell.formula))
      continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    candidates.push({ row: addr.row, col: addr.col });
  }
  bySheet.set(sheet, candidates);
  return candidates;
}

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

  private readonly hits: GridHitState;

  private readonly conditionalCache = createConditionalCache();

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
    this.hits = gridHitState(deps.host);
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
      frozenRowEnd(layout),
      viewport.navigationRange?.r0 ?? 0,
    );
    const firstCol = Math.max(
      viewport.colStart,
      frozenColEnd(layout),
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
    this.hits.outlineToggles = [];
    if (state.ui.showHeaders !== false) this.paintHeaders(state, theme, cols, rows);
    this.paintFreezeDividers(state, theme, cols, rows);
    this.paintBorders(state, theme, cols, rows);
    // Page chrome covers the gutters, so it has to follow every pass that
    // paints across the whole canvas — gridlines and freeze dividers do.
    this.hits.pageBands = [];
    this.hits.rulerHandles = [];
    this.hits.pageBreakHandles = [];
    if (pagination) this.paintPageChrome(state, theme, pagination, cols, rows);
    this.paintSpills(state, theme, cols, rows);
    this.paintActive(state, theme, cols, rows);
    this.paintEditorRefs(state);
    this.paintChartSourceRefs(state);
    paintVisibleNotes(this.ctx, state, cols, rows);
    this.paintTraces(state, cols, rows);
  }

  /** Pagination for the active sheet, or null in Normal view. Reaches one
   *  viewport past the visible slice so the page opening just off screen is
   *  already laid out when the user scrolls onto it. */
  private paginationForView(state: ViewState): SheetPagination | null {
    if (state.ui.workbookView === 'normal') return null;
    const { viewport } = state;
    return paginationFor(this.getState(), state.data.sheetIndex, {
      engine: this.getWb() ?? undefined,
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
      this.hits.pageBands = paintPageLayoutChrome(
        this.chromeCtx(),
        state,
        theme,
        pagination,
        setup,
        { addHeader: strings.pageView.addHeader, addFooter: strings.pageView.addFooter },
        cols,
        rows,
      );
      this.hits.rulerHandles = paintPageRulers(
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
    this.hits.pageBreakHandles = paintPageBreakPreview(
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

  /** The screen box around the visible part of `range`, or null when none of it is visible. */
  private rangeBounds(state: ViewState, range: Range): Rect | null {
    const rects = rangeRects(state.layout, state.viewport, range);
    // One box around the whole range, so a 2x2 range reads as a single outline.
    if (rects.length === 0) return null;
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
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  }

  private paintEditorRefs(state: ViewState): void {
    const refs = state.ui.editorRefs;
    if (!refs || refs.length === 0) return;
    const sheet = state.data.sheetIndex;
    for (const ref of refs) {
      const bounds = this.rangeBounds(state, {
        sheet,
        r0: ref.r0,
        c0: ref.c0,
        r1: ref.r1,
        c1: ref.c1,
      });
      if (bounds) paintRefHighlight(this.ctx, bounds, ref.colorIndex);
    }
  }

  private paintChartSourceRefs(state: ViewState): void {
    for (const ref of state.ui.chartSourceRefs ?? []) {
      if (ref.sheet !== state.data.sheetIndex) continue;
      const bounds = this.rangeBounds(state, ref);
      if (bounds) paintChartSourceHighlight(this.ctx, bounds, ref.role);
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
      fontEffects:
        (this.host.closest?.<HTMLElement>('.fc-host') ?? this.host).dataset?.fcPlatform === 'mac',
      conditionalCache: this.conditionalCache,
    };
  }

  private paintCells(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    this.hits.errorTriangles = paintCells(this.cellCtx(state), state, theme, cols, rows);
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
    this.hits.outlineToggles = paintHeaders(this.chromeCtx(), state, theme, cols, rows);
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
    const wb = this.getWb();
    const supportsSpillInfo =
      wb?.capabilities?.spillInfo === true && typeof wb.spillInfo === 'function';
    const useVisibleSpillQueries =
      supportsSpillInfo && data.cells.size > rows.visible.length * cols.visible.length;
    const frameSpillInfo = new Map<string, ReturnType<WorkbookHandle['spillInfo']>>();
    let ranges: Range[];

    if (useVisibleSpillQueries && wb) {
      // Query every visible coordinate: phantom cells carry the anchor's
      // shape, even when the anchor itself is outside the viewport.
      const seenAnchors = new Set<string>();
      ranges = [];
      for (const row of rows.visible) {
        for (const col of cols.visible) {
          const key = addrKey({ sheet: data.sheetIndex, row, col });
          const info = wb.spillInfo(data.sheetIndex, row, col);
          frameSpillInfo.set(key, info);
          if (!info || (info.rows <= 1 && info.cols <= 1)) continue;
          const anchorKey = `${info.anchorRow}:${info.anchorCol}`;
          if (seenAnchors.has(anchorKey)) continue;
          seenAnchors.add(anchorKey);
          ranges.push({
            sheet: data.sheetIndex,
            r0: info.anchorRow,
            c0: info.anchorCol,
            r1: info.anchorRow + info.rows - 1,
            c1: info.anchorCol + info.cols - 1,
          });
        }
      }
    } else {
      // Prefer the engine's authoritative spill list when available; fall back
      // to the JS heuristic for stub mode or older engine package builds.
      ranges = wb?.spillRanges(data.sheetIndex) ?? findSpillRanges(data.cells, data.sheetIndex);
    }

    for (const r of ranges) {
      for (const rect of rangeRects(layout, viewport, r)) {
        paintSpillOutline(this.ctx, rect, theme);
      }
    }

    // #SPILL! obstruction outline. The engine's `spillInfo` reports the
    // attempted shape even when the result couldn't materialise; we use
    // that to figure out which cells are blocking and ring them in red.
    if (!wb || typeof wb.spillInfo !== 'function') return;
    for (const { row, col } of spillErrorCandidates(data.cells, data.sheetIndex)) {
      const key = addrKey({ sheet: data.sheetIndex, row, col });
      const info = frameSpillInfo.has(key)
        ? (frameSpillInfo.get(key) ?? null)
        : wb.spillInfo(data.sheetIndex, row, col);
      if (!info) continue;
      const target: Range = {
        sheet: data.sheetIndex,
        r0: info.anchorRow,
        c0: info.anchorCol,
        r1: info.anchorRow + info.rows - 1,
        c1: info.anchorCol + info.cols - 1,
      };
      for (const visibleRow of rows.visible) {
        if (visibleRow < target.r0 || visibleRow > target.r1) continue;
        for (const visibleCol of cols.visible) {
          if (visibleCol < target.c0 || visibleCol > target.c1) continue;
          if (visibleRow === info.anchorRow && visibleCol === info.anchorCol) continue;
          const blocker = data.cells.get(
            addrKey({ sheet: data.sheetIndex, row: visibleRow, col: visibleCol }),
          );
          if (!blocker) continue;
          if (blocker.formula === null && blocker.value.kind === 'blank') continue;
          const rect = cellRectIn(layout, cols, rows, visibleRow, visibleCol);
          paintSpillBlocker(this.ctx, rect);
        }
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
    this.hits.fillHandle = null;
    this.hits.validationChevron = null;
    this.hits.totalsChevron = null;
    // A selected chart or picture replaces the cell cursor, as in Excel.
    if (ui.selectedObject) return;

    const activeMerge = mergeRangeAt(state, a);
    const activeVisualRange: Range = activeMerge ?? {
      sheet: a.sheet,
      r0: a.row,
      c0: a.col,
      r1: a.row,
      c1: a.col,
    };
    const activeRects = rangeRects(layout, viewport, activeVisualRange);
    // A multi-cell range carries the one outline; the active cell inside it is
    // only left unshaded, as in Excel. A merge selected alone is its own range.
    const multiCell = r.r0 !== r.r1 || r.c0 !== r.c1;
    const rangeOutline = multiCell ? rangeRects(layout, viewport, r) : activeRects;
    const outlineRects = sameRectList(activeRects, rangeOutline) ? activeRects : rangeOutline;
    for (const bounds of outlineRects) {
      paintSelectionOutline(this.ctx, bounds, theme, outlineRects !== activeRects);
    }

    const activeAnchor = activeMerge
      ? { sheet: activeMerge.sheet, row: activeMerge.r0, col: activeMerge.c0 }
      : a;
    const validationBounds = trailingRect(activeRects, layout.rtl);
    if (validationBounds && shouldShowValidationChevron(validationAt(format, activeAnchor))) {
      const rect = paintValidationChevron(this.ctx, validationBounds, theme, layout.rtl);
      this.hits.validationChevron = { rect, row: activeAnchor.row, col: activeAnchor.col };
    } else if (validationBounds) {
      const totals = totalsCellAt(state, a.sheet, activeAnchor.row, activeAnchor.col);
      if (totals) {
        const rect = paintValidationChevron(this.ctx, validationBounds, theme, layout.rtl);
        this.hits.totalsChevron = {
          rect,
          row: activeAnchor.row,
          col: activeAnchor.col,
          tableId: totals.table.id,
          column: totals.column,
        };
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
      this.hits.fillHandle = paintFillHandle(this.ctx, handleBounds, theme, layout.rtl);
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
  type GridHitState,
  gridHitState,
  isPlainTextOverflowCandidate,
  type TotalsChevron,
  VALIDATION_TRIANGLE_COLOR,
  type ValidationChevron,
} from './grid/hit-state.js';

import { validationAt } from '../store/range-validations.js';
