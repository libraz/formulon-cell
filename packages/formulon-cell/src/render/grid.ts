import { totalsCellAt } from '../commands/format-as-table.js';
import { paginationFor, type SheetPagination } from '../commands/pagination.js';
import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import { effectiveCellFormatFromEngine } from '../engine/cell-format-projection.js';
import { pivotStyleLayer } from '../engine/pivot-style.js';
import { findSpillRanges, looksLikeArrayFormula } from '../engine/spill.js';
import { spillBlockers, spillOccupancy } from '../engine/spill-blockers.js';
import type { CellValue, Range } from '../engine/types.js';
import { SPILL_ERROR_CODE } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { frozenColEnd, frozenRowEnd } from '../store/freeze-band.js';
import { headerLabelScale } from '../store/header-strip.js';
import { validationAt } from '../store/range-validations.js';
import { isWholeColumnRange, isWholeRowRange } from '../store/selection-geometry.js';
import type { CellFormat, State } from '../store/store.js';
import { getPageSetup } from '../store/store.js';
import type { ResolvedTheme } from '../theme/resolve.js';
import { automaticFontRowLayout } from './automatic-font-row-layout.js';
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
  pointPx,
  type Rect,
  rangeRects,
  rowGap,
  rowHeight,
  type ViewState,
} from './geometry.js';
import { paintBorders } from './grid/borders.js';
import {
  type CellDisplayResolver,
  type CellPaintContext,
  paintCells,
  paintPivotPlaceholders,
} from './grid/cells.js';
import type { ChromePaintContext } from './grid/chrome-context.js';
import { paintHeaders } from './grid/headers.js';
import { type GridHitState, gridHitState, shouldShowValidationChevron } from './grid/hit-state.js';
import { gridLinesShown, paintFreezeDividers, paintGridLines } from './grid/lines.js';
import { mergeRangeAt } from './grid/merged-cells.js';
import { paintVisibleNotes } from './grid/notes.js';
import {
  paintPageBreakPreview,
  paintPageLayoutBackground,
  paintPageLayoutChrome,
  paintPageRulers,
} from './grid/page-view.js';
import { sameRectList, trailingRect } from './grid/range-rects.js';
import { styleElementFormatSource } from './grid/table-format.js';
import { paintTraces } from './grid/traces.js';
import { cellFontFamily } from './painters/text.js';
import {
  copyMarqueePhase,
  isFillShrinkPreview,
  paintActiveCellBox,
  paintChartSourceHighlight,
  paintCopyMarquee,
  paintErrorSmartTag,
  paintFillHandle,
  paintFillPreview,
  paintFillShrinkShade,
  paintMoveBand,
  paintRangeOutlineThin,
  paintRefHighlight,
  paintSelectionOutline,
  paintSpillBlocker,
  paintSpillOutline,
  paintSpillTarget,
  paintSpillWarningTag,
  paintTableGrip,
  paintValidationChevron,
} from './painters.js';

/** Header labels follow a Normal font larger than the default one at its own size. */
export const headerLabelSize = (normalSize: number | undefined, fallback: number): number =>
  normalSize !== undefined && normalSize > DEFAULT_NORMAL_SIZE ? normalSize : fallback;

const DEFAULT_NORMAL_SIZE = 11;

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
  /** Vertical rules the last cell pass's overflowing text covers. */
  private coveredRules = new Set<string>();

  private cssWidth = 0;

  private cssHeight = 0;

  private rafId = 0;

  private resolutionQueryDpr: number | undefined;

  private resolutionUnwatch: (() => void) | null = null;

  /** Platform profile resolved for the paint in flight; null outside a paint. */
  private macProfileInPaint: boolean | null = null;

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
    const pxW = Math.round(this.cssWidth * this.dpr);
    const pxH = Math.round(this.cssHeight * this.dpr);
    // Assigning either dimension clears the backing store, so only on change.
    if (this.canvas.width !== pxW) this.canvas.width = pxW;
    if (this.canvas.height !== pxH) this.canvas.height = pxH;
    this.canvas.style.width = `${this.cssWidth}px`;
    this.canvas.style.height = `${this.cssHeight}px`;
    this.watchResolution();
    this.resizeViewport();
    this.invalidate();
  }

  /** Re-runs `resize` when the display scale changes (window moved to another monitor, browser zoom). */
  private watchResolution(): void {
    if (this.resolutionQueryDpr === window.devicePixelRatio) return;
    this.unwatchResolution();
    if (typeof window.matchMedia !== 'function') return;
    const dpr = window.devicePixelRatio;
    const query = window.matchMedia(`(resolution: ${dpr}dppx)`);
    const onChange = (): void => this.resize();
    query.addEventListener?.('change', onChange);
    this.resolutionQueryDpr = dpr;
    this.resolutionUnwatch = () => query.removeEventListener?.('change', onChange);
  }

  private unwatchResolution(): void {
    this.resolutionUnwatch?.();
    this.resolutionUnwatch = null;
    this.resolutionQueryDpr = undefined;
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
    // The first invalidate of a frame sizes now; the frame pass repeats it
    // after the native format epoch settles. Later ones wait for that pass.
    if (this.rafId) return;
    if (this.cssWidth > 0 && this.cssHeight > 0) this.resizeViewport();
    this.rafId = requestAnimationFrame(() => {
      try {
        this.resizeViewport();
      } finally {
        this.rafId = 0;
      }
      this.paint();
    });
  }

  dispose(): void {
    if (this.rafId) cancelAnimationFrame(this.rafId);
    this.rafId = 0;
    if (this.marqueeRafId) cancelAnimationFrame(this.marqueeRafId);
    this.marqueeRafId = 0;
    this.unwatchResolution();
  }

  private paint(): void {
    if (this.cssWidth === 0 || this.cssHeight === 0) return;
    this.macProfileInPaint = this.macProfileLookup();
    try {
      this.paintFrame();
    } finally {
      this.macProfileInPaint = null;
    }
  }

  private paintFrame(): void {
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
    if (gridLinesShown(state)) this.paintGridLines(state, theme, cols, rows);
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
    this.paintPivotPlaceholders(state, theme);
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
    const baseState = this.getState();
    const projectedLayout = automaticFontRowLayout(baseState);
    const geometryChanged = projectedLayout !== baseState.layout;
    const paginationState = geometryChanged ? { ...baseState, layout: projectedLayout } : baseState;
    const { viewport } = state;
    const engine = geometryChanged ? null : this.getWb();
    return paginationFor(paginationState, state.data.sheetIndex, {
      ...(engine ? { engine } : {}),
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
        {
          addHeader: strings.pageView.addHeader,
          addFooter: strings.pageView.addFooter,
          addData: strings.pageView.addData,
        },
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
    const macPt = this.macProfile() ? pointPx(state.layout, state.viewport) : undefined;
    for (const ref of refs) {
      const bounds = this.rangeBounds(state, {
        sheet,
        r0: ref.r0,
        c0: ref.c0,
        r1: ref.r1,
        c1: ref.c1,
      });
      if (bounds) paintRefHighlight(this.ctx, bounds, ref.colorIndex, macPt);
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
    const bare = pivotStyleLayer(this.getWb(), state).ranges;
    paintGridLines(this.chromeCtx(), state, theme, cols, rows, bare);
  }

  /** The error-check smart tag, shown on the Mac profile for the flagged cell a mouse selection landed on. */
  private paintErrorTag(
    state: ViewState,
    anchor: { sheet: number; row: number; col: number },
    activeRects: readonly Rect[],
  ): void {
    if (!this.macProfile()) return;
    if (this.hits.errorTagCell === 'pending') this.hits.errorTagCell = { ...anchor };
    const armed = this.hits.errorTagCell;
    if (
      !armed ||
      armed.sheet !== anchor.sheet ||
      armed.row !== anchor.row ||
      armed.col !== anchor.col
    ) {
      return;
    }
    const triangle = this.hits.errorTriangles.find(
      (hit) =>
        hit.kind === 'error' &&
        hit.addr.sheet === anchor.sheet &&
        hit.addr.row === anchor.row &&
        hit.addr.col === anchor.col,
    );
    const bounds = activeRects[0];
    if (!triangle || !bounds) return;
    const ptPx = pointPx(state.layout, state.viewport);
    const rect = paintErrorSmartTag(this.ctx, bounds, ptPx, state.layout.rtl);
    this.hits.errorTag = { rect, addr: { ...anchor }, kind: 'error' };
  }

  /** Whether the host renders the Mac profile, which several painters follow. */
  private macProfile(): boolean {
    return this.macProfileInPaint ?? this.macProfileLookup();
  }

  private macProfileLookup(): boolean {
    return (
      (this.host.closest?.<HTMLElement>('.fc-host') ?? this.host).dataset?.fcPlatform === 'mac'
    );
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
      fontEffects: this.macProfile(),
      conditionalCache: this.conditionalCache,
      dpr: this.dpr,
      pivotEmptyPrompt: this.getStrings().workbookObjects.pivotEmptyPrompt,
    };
  }

  private paintCells(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const result = paintCells(this.cellCtx(state), state, theme, cols, rows);
    this.hits.errorTriangles = result.errorTriangles;
    this.coveredRules = result.coveredRules;
  }

  /** Empty-report placeholders cover the selection, clipped below the headers. */
  private paintPivotPlaceholders(state: ViewState, theme: ResolvedTheme): void {
    const { layout } = state;
    const ctx = this.ctx;
    const x = layout.rtl ? 0 : gridOriginX(layout);
    const y = gridOriginY(layout);
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, this.cssWidth - x, this.cssHeight - y);
    ctx.clip();
    paintPivotPlaceholders(this.cellCtx(state), state, theme);
    ctx.restore();
  }

  private paintBorders(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const wb = this.getWb();
    const sheet = state.data.sheetIndex;
    const effective = new Map<string, Partial<CellFormat> | null>();
    const styleFormatAt = styleElementFormatSource(state, wb);
    const engineBorders =
      wb?.capabilities.effectiveStyle && typeof wb.getEffectiveStyle === 'function'
        ? (row: number, col: number) =>
            effectiveCellFormatFromEngine(wb, sheet, row, col, effective)?.borders
        : undefined;
    const inherited = (row: number, col: number) =>
      engineBorders?.(row, col) ?? styleFormatAt(row, col)?.borders;
    paintBorders(this.ctx, this.dpr, state, theme, cols, rows, inherited, this.coveredRules);
  }

  private paintHeaders(
    state: ViewState,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    // Headings take the workbook's Normal font, as the desktop draws them.
    const normalFont = this.getWb()?.workbookDisplayFont;
    this.hits.outlineToggles = paintHeaders(
      this.chromeCtx(),
      state,
      theme,
      cols,
      rows,
      cellFontFamily(normalFont?.name, theme.fontUi),
      headerLabelSize(normalFont?.size, theme.textHeader) *
        headerLabelScale(state.layout, state.viewport.zoom),
    );
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
    let ranges: Range[];

    if (useVisibleSpillQueries && wb) {
      // Query every visible coordinate: phantom cells carry the anchor's
      // shape, even when the anchor itself is outside the viewport.
      const seenAnchors = new Set<string>();
      ranges = [];
      for (const row of rows.visible) {
        for (const col of cols.visible) {
          const info = wb.spillInfo(data.sheetIndex, row, col);
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

    const active = state.selection.active;
    const inRange = (r: Range): boolean =>
      active.sheet === r.sheet &&
      active.row >= r.r0 &&
      active.row <= r.r1 &&
      active.col >= r.c0 &&
      active.col <= r.c1;
    // The Mac desktop outlines a spill range, 1 pt wide, only while one of its
    // cells is active, and boxes a blocked formula's target only while its
    // anchor is; other profiles keep the faint always-on ring and red blockers.
    const mac = this.macProfile();
    const macPtPx = mac ? pointPx(layout, viewport) : undefined;
    for (const r of ranges) {
      if (mac && !inRange(r)) continue;
      for (const rect of rangeRects(layout, viewport, r)) {
        paintSpillOutline(this.ctx, rect, theme, macPtPx);
      }
    }
    for (const spill of pivotStyleLayer(wb, state).spills) {
      for (const rect of rangeRects(layout, viewport, spill.needed)) {
        paintSpillTarget(this.ctx, rect, macPtPx);
      }
      // The warning tag sits beside the anchor while it is the active cell.
      const { anchor } = spill;
      if (active.sheet === anchor.sheet && active.row === anchor.row && active.col === anchor.col) {
        const cell = rangeRects(layout, viewport, {
          r0: anchor.row,
          c0: anchor.col,
          r1: anchor.row,
          c1: anchor.col,
        })[0];
        if (cell) paintSpillWarningTag(this.ctx, cell, layout.rtl);
      }
    }
    // A formula held at #SPILL! carries the same tag while its anchor is active.
    const activeCell =
      mac && active.sheet === data.sheetIndex ? data.cells.get(addrKey(active)) : null;
    if (
      activeCell?.formula &&
      activeCell.value.kind === 'error' &&
      activeCell.value.code === SPILL_ERROR_CODE
    ) {
      const cell = rangeRects(layout, viewport, {
        r0: active.row,
        c0: active.col,
        r1: active.row,
        c1: active.col,
      })[0];
      if (cell) paintSpillWarningTag(this.ctx, cell, layout.rtl);
    }

    // #SPILL! obstruction outline: the engine reports the rectangle a blocked
    // formula attempted, and the cells holding it back are ringed in red.
    if (!wb || typeof wb.blockedSpillRange !== 'function') return;
    let occupied: ((row: number, col: number) => boolean) | undefined;
    for (const { row, col } of spillErrorCandidates(data.cells, data.sheetIndex)) {
      const target = wb.blockedSpillRange(data.sheetIndex, row, col);
      if (!target) continue;
      if (mac) {
        const isAnchor =
          active.sheet === data.sheetIndex && active.row === row && active.col === col;
        if (!isAnchor) continue;
        for (const rect of rangeRects(layout, viewport, target)) {
          paintSpillTarget(this.ctx, rect, macPtPx);
        }
        continue;
      }
      occupied ??= spillOccupancy(wb, data.sheetIndex, (row, col) =>
        data.cells.get(addrKey({ sheet: data.sheetIndex, row, col })),
      );
      const anchor = { sheet: data.sheetIndex, row, col };
      for (const blocker of spillBlockers(target, anchor, occupied, {
        rows: rows.visible,
        cols: cols.visible,
      })) {
        paintSpillBlocker(this.ctx, cellRectIn(layout, cols, rows, blocker.row, blocker.col));
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
    this.hits.errorTag = null;
    for (const table of state.tables.tables) {
      if (table.range.sheet !== data.sheetIndex) continue;
      const { r1, c1 } = table.range;
      const [corner] = rangeRects(layout, viewport, { r0: r1, c0: c1, r1, c1 });
      if (corner) paintTableGrip(this.ctx, corner, layout.rtl);
    }
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
    // An empty report's placeholder covers its own cursor; an outline drawn
    // around it would stick out of the box's edge.
    const underPlaceholder =
      r.r0 === r.r1 &&
      r.c0 === r.c1 &&
      pivotStyleLayer(this.getWb(), state).placeholders.some(
        ({ range }) =>
          range.r0 <= a.row && a.row <= range.r1 && range.c0 <= a.col && a.col <= range.c1,
      );
    // A multi-cell range carries the one outline; the active cell inside it is
    // only left unshaded, as in Excel. A merge selected alone is its own range.
    const multiCell = r.r0 !== r.r1 || r.c0 !== r.c1;
    const rangeOutline = multiCell ? rangeRects(layout, viewport, r) : activeRects;
    const outlineRects = sameRectList(activeRects, rangeOutline) ? activeRects : rangeOutline;
    // Disjoint ranges: every range gets a thin rule and the active cell a thin
    // box, with no thick frame and no fill handle.
    const extraRanges = selection.extraRanges ?? [];
    const multiRange = extraRanges.length > 0;
    if (multiRange) {
      for (const range of [r, ...extraRanges]) {
        for (const bounds of rangeRects(layout, viewport, range)) {
          paintRangeOutlineThin(this.ctx, bounds, theme, layout.rtl);
        }
      }
      for (const bounds of activeRects) paintActiveCellBox(this.ctx, bounds, theme, layout.rtl);
    } else if (!underPlaceholder) {
      const limit = {
        x: layout.headerColWidth === 0 && !layout.rtl ? gridOriginX(layout) : undefined,
        y: layout.headerRowHeight === 0 ? gridOriginY(layout) : undefined,
      };
      for (const bounds of outlineRects) {
        paintSelectionOutline(
          this.ctx,
          bounds,
          theme,
          outlineRects !== activeRects,
          layout.rtl,
          limit,
        );
      }
    }

    const activeAnchor = activeMerge
      ? { sheet: activeMerge.sheet, row: activeMerge.r0, col: activeMerge.c0 }
      : a;
    const validationBounds = trailingRect(activeRects, layout.rtl);
    // The Mac desktop sets its chrome in points.
    const macPt = this.macProfile() ? pointPx(layout, viewport) : undefined;
    if (validationBounds && shouldShowValidationChevron(validationAt(format, activeAnchor))) {
      const rect = paintValidationChevron(this.ctx, validationBounds, theme, layout.rtl, macPt);
      this.hits.validationChevron = {
        rect,
        cell: validationBounds,
        row: activeAnchor.row,
        col: activeAnchor.col,
      };
    } else if (validationBounds) {
      const totals = totalsCellAt(state, a.sheet, activeAnchor.row, activeAnchor.col);
      if (totals) {
        const rect = paintValidationChevron(this.ctx, validationBounds, theme, layout.rtl, macPt);
        this.hits.totalsChevron = {
          rect,
          row: activeAnchor.row,
          col: activeAnchor.col,
          tableId: totals.table.id,
          column: totals.column,
        };
      }
    }

    this.paintErrorTag(state, activeAnchor, activeRects);

    const preview = ui.fillPreview;
    if (preview && isFillShrinkPreview(r, preview)) {
      // A fill pulled back inside its source lightens the cells it will drop.
      const dropped =
        preview.r1 < r.r1 ? { ...r, r0: preview.r1 + 1 } : { ...r, c0: preview.c1 + 1 };
      for (const rect of rangeRects(layout, viewport, dropped))
        paintFillShrinkShade(this.ctx, rect);
    } else if (preview) {
      const rects = rangeRects(layout, viewport, preview);
      // The Mac desktop shows the fill target as its green drag band.
      for (const r of rects) {
        if (macPt !== undefined) paintMoveBand(this.ctx, r, macPt);
        else paintFillPreview(this.ctx, r, theme);
      }
    }
    // A drag-move's drop target takes the same mark as a fill target.
    if (ui.movePreview) {
      for (const r of rangeRects(layout, viewport, ui.movePreview)) {
        if (macPt !== undefined) paintMoveBand(this.ctx, r, macPt);
        else paintFillPreview(this.ctx, r, theme);
      }
    }
    const copyRanges = ui.copyRanges && ui.copyRanges.length > 0 ? ui.copyRanges : null;
    const rangesToPaint = (copyRanges ?? (ui.copyRange ? [ui.copyRange] : [])).filter(
      (copyRange) => copyRange.sheet === data.sheetIndex,
    );
    if (rangesToPaint.length > 0) {
      const phase = copyMarqueePhase(performance.now(), macPt !== undefined);
      for (const copyRange of rangesToPaint) {
        const rects = rangeRects(layout, viewport, copyRange);
        for (const r of rects) paintCopyMarquee(this.ctx, r, phase, macPt);
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
    const handleRects = rangeRects(layout, viewport, handleRange);
    let handleBounds = trailingRect(handleRects, layout.rtl);
    let handleRing = false;
    // A whole-column selection has no visible bottom and a whole-row one no
    // visible end: the handle sits on the first row's top / first column's edge.
    if (handleBounds && isWholeColumnRange(handleRange)) {
      handleBounds = { ...handleBounds, y: Math.min(...handleRects.map((rect) => rect.y)), h: 4 };
      handleRing = true;
    } else if (handleBounds && isWholeRowRange(handleRange) && !layout.rtl) {
      handleBounds = { ...handleBounds, x: Math.min(...handleRects.map((rect) => rect.x)), w: 4 };
      handleRing = true;
    }
    if (handleBounds && !multiRange && !underPlaceholder) {
      this.hits.fillHandle = paintFillHandle(this.ctx, handleBounds, theme, layout.rtl, handleRing);
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
