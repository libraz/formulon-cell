import { coerceInput } from '../commands/coerce-input.js';
import { validateAgainst } from '../commands/validate.js';
import { evaluateCfFromEngine } from '../engine/cf-sync.js';
import { makeRangeResolver, type RangeResolver } from '../engine/range-resolver.js';
import { findSpillRanges } from '../engine/spill.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import type { CellFormat, CellValidation, State } from '../store/store.js';
import type { ResolvedTheme } from '../theme/resolve.js';
import { evaluateConditional } from './conditional.js';
import {
  type AxisLayout,
  buildColLayout,
  buildRowLayout,
  cellRectIn,
  colLabel,
  gridOriginX,
  gridOriginY,
  isColVisible,
  isRowVisible,
  type Rect,
  rangeRects,
} from './geometry.js';
import {
  CONDITIONAL_ICON_GUTTER,
  type OutlineToggleHit,
  paintActiveCellOutline,
  paintCellBackground,
  paintCellBorders,
  paintCellFill,
  paintCellText,
  paintCommentMarker,
  paintConditionalIcon,
  paintErrorTriangle,
  paintFillHandle,
  paintFillPreview,
  paintLockMarker,
  paintOutlineGutters,
  paintRefHighlight,
  paintSpillOutline,
  paintTraceArrow,
  paintTraceDot,
  paintValidationChevron,
  paintValidationTriangle,
  TRACE_DEPENDENT_COLOR,
  TRACE_PRECEDENT_COLOR,
} from './painters.js';
import { paintCellSparkline } from './sparkline.js';

export type ErrorTriangleKind = 'error' | 'validation';

/** Hot-zone of an error/validation triangle painted in this frame. The
 *  click layer (mount.ts) hit-tests these to open the popover menu. */
export interface ErrorTriangleHit {
  rect: Rect;
  addr: { sheet: number; row: number; col: number };
  kind: ErrorTriangleKind;
}

/** Color used for formula-error triangles (Excel-style green). */
export const ERROR_TRIANGLE_COLOR = '#2ea043';
/** Color used for data-validation violation triangles (Excel-style red). */
export const VALIDATION_TRIANGLE_COLOR = '#d24545';

let cachedFillHandleRect: Rect | null = null;
let cachedValidationChevron: { rect: Rect; row: number; col: number } | null = null;
let cachedOutlineToggles: OutlineToggleHit[] = [];
let cachedErrorTriangles: ErrorTriangleHit[] = [];

/** Latest hit-rects of all error / validation triangles painted in this
 *  frame. The mount-level click handler hit-tests these to open the
 *  error-info popover. */
export function getErrorTriangleHits(): ErrorTriangleHit[] {
  return cachedErrorTriangles;
}

function setErrorTriangleHits(hits: ErrorTriangleHit[]): void {
  cachedErrorTriangles = hits;
}

/** Excel error sentinels that the renderer surfaces with a green corner
 *  triangle. Engine-typed errors take the `value.kind === 'error'` branch;
 *  the string set covers cases where a custom formatter / passthrough
 *  layer left a string sentinel in the cell. */
const ERROR_SENTINELS: ReadonlySet<string> = new Set([
  '#DIV/0!',
  '#NAME?',
  '#REF!',
  '#VALUE!',
  '#NUM!',
  '#N/A',
  '#NULL!',
  '#CIRCULAR!',
]);

/** Detect whether `cell.value` should surface an error indicator. */
export function detectErrorKind(value: CellValue): boolean {
  if (value.kind === 'error') return true;
  if (value.kind === 'text') return ERROR_SENTINELS.has(value.value);
  return false;
}

/** Detect whether `cell.value` violates `validation`. Returns false when
 *  validation is missing, when the cell is blank and `allowBlank` is set,
 *  or when the value is itself an error (we surface that as an error
 *  triangle, not a validation triangle). */
export function detectValidationViolation(
  value: CellValue,
  validation: CellValidation | undefined,
  resolveRange?: RangeResolver,
): boolean {
  if (!validation) return false;
  if (value.kind === 'error') return false;
  // Re-use coerceInput by stringifying the value first — same shape the
  // keyboard / formula-bar paths feed validateAgainst.
  let raw: string;
  switch (value.kind) {
    case 'blank':
      raw = '';
      break;
    case 'number':
      raw = String(value.value);
      break;
    case 'bool':
      raw = value.value ? 'TRUE' : 'FALSE';
      break;
    case 'text':
      raw = value.value;
      break;
  }
  const coerced = coerceInput(raw);
  const outcome = validateAgainst(validation, coerced, resolveRange);
  return !outcome.ok;
}

/** Latest hit-rects of all outline +/- toggles painted in this frame. The
 *  pointer layer hit-tests these to route clicks to collapse/expand. */
export function getOutlineToggleHits(): OutlineToggleHit[] {
  return cachedOutlineToggles;
}

function setOutlineToggles(hits: OutlineToggleHit[]): void {
  cachedOutlineToggles = hits;
}

/** Latest device-space bounds of the fill handle. Hit-tested by the pointer
 *  layer to start a fill drag. Null while the handle is offscreen. */
export function getFillHandleRect(): Rect | null {
  return cachedFillHandleRect;
}

function setFillHandleRect(r: Rect | null): void {
  cachedFillHandleRect = r;
}

/** Bounds + addr of the active cell's validation chevron, or null when the
 *  active cell has no list validation. */
export function getValidationChevron(): { rect: Rect; row: number; col: number } | null {
  return cachedValidationChevron;
}

function setValidationChevron(v: { rect: Rect; row: number; col: number } | null): void {
  cachedValidationChevron = v;
}

const inRange = (a: Addr, r: Range): boolean =>
  a.row >= r.r0 && a.row <= r.r1 && a.col >= r.c0 && a.col <= r.c1;

export interface RendererDeps {
  host: HTMLElement;
  canvas: HTMLCanvasElement;
  getState: () => State;
  getTheme: () => ResolvedTheme;
  /** Optional accessor for the active workbook. When supplied and the engine
   *  exposes `evaluateCfRange`, conditional-format rules loaded from the
   *  .xlsx are evaluated alongside the JS-side rule set and overlaid on top
   *  of the rendered cells. */
  getWb?: () => WorkbookHandle | null;
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

  private readonly getDisplay: RendererDeps['getDisplay'];

  private dpr = 1;

  private cssWidth = 0;

  private cssHeight = 0;

  private rafId = 0;

  constructor(deps: RendererDeps) {
    this.host = deps.host;
    this.canvas = deps.canvas;
    const ctx = this.canvas.getContext('2d', { alpha: false });
    if (!ctx) throw new Error('formulon-cell: 2D canvas context unavailable');
    this.ctx = ctx;
    this.getState = deps.getState;
    this.getTheme = deps.getTheme;
    this.getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
    this.getDisplay = deps.getDisplay;
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
    this.invalidate();
  }

  invalidate(): void {
    if (this.rafId) return;
    this.rafId = requestAnimationFrame(() => {
      this.rafId = 0;
      this.paint();
    });
  }

  dispose(): void {
    if (this.rafId) cancelAnimationFrame(this.rafId);
    this.rafId = 0;
  }

  private paint(): void {
    if (this.cssWidth === 0 || this.cssHeight === 0) return;

    const state = this.getState();
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

    if (state.ui.showGridLines !== false) this.paintGridLines(state, theme, cols, rows);
    this.paintCells(state, theme, cols, rows);
    this.paintBorders(state, theme, cols, rows);
    if (state.ui.showHeaders !== false) this.paintHeaders(state, theme, cols, rows);
    this.paintFreezeDividers(state, theme, cols, rows);
    this.paintSpills(state, theme, cols, rows);
    this.paintActive(state, theme, cols, rows);
    this.paintEditorRefs(state);
    this.paintTraces(state, cols, rows);

    // Hand the engine the inclusive rect of cells we just painted so the next
    //  setFormula can run a partialRecalc bounded to what the user sees.
    const wb = this.getWb();
    const firstRow = rows.visible[0];
    const lastRow = rows.visible[rows.visible.length - 1];
    const firstCol = cols.visible[0];
    const lastCol = cols.visible[cols.visible.length - 1];
    if (
      wb &&
      firstRow !== undefined &&
      lastRow !== undefined &&
      firstCol !== undefined &&
      lastCol !== undefined
    ) {
      wb.setViewportHint(state.data.sheetIndex, firstRow, firstCol, lastRow, lastCol);
    }
  }

  private paintEditorRefs(state: State): void {
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

  private paintGridLines(
    state: State,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const ctx = this.ctx;
    const { layout } = state;
    const align = 0.5 / this.dpr;

    ctx.strokeStyle = theme.rule;
    ctx.lineWidth = 1 / this.dpr;
    ctx.beginPath();

    const firstRow = rows.visible[0] ?? 0;
    const firstCol = cols.visible[0] ?? 0;

    for (const c of cols.visible) {
      const rect = cellRectIn(layout, cols, rows, firstRow, c);
      const xx = Math.round(rect.x) + align;
      ctx.moveTo(xx, 0);
      ctx.lineTo(xx, this.cssHeight);
    }
    const lastCol = cols.visible[cols.visible.length - 1];
    if (lastCol !== undefined) {
      const rect = cellRectIn(layout, cols, rows, firstRow, lastCol);
      const xx = Math.round(rect.x + rect.w) + align;
      ctx.moveTo(xx, 0);
      ctx.lineTo(xx, this.cssHeight);
    }

    for (const r of rows.visible) {
      const rect = cellRectIn(layout, cols, rows, r, firstCol);
      const yy = Math.round(rect.y) + align;
      ctx.moveTo(0, yy);
      ctx.lineTo(this.cssWidth, yy);
    }
    const lastRow = rows.visible[rows.visible.length - 1];
    if (lastRow !== undefined) {
      const rect = cellRectIn(layout, cols, rows, lastRow, firstCol);
      const yy = Math.round(rect.y + rect.h) + align;
      ctx.moveTo(0, yy);
      ctx.lineTo(this.cssWidth, yy);
    }

    ctx.stroke();
  }

  private paintCells(state: State, theme: ResolvedTheme, cols: AxisLayout, rows: AxisLayout): void {
    const ctx = this.ctx;
    const { layout, data, selection, format, merges, sparkline, errorIndicators, protection } =
      state;
    const active = selection.active;
    const conditional = evaluateConditional(state);
    const sparklines = sparkline.sparklines;
    const ignored = errorIndicators.ignoredErrors;
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

    // Compute merged cell bounds (anchor cell expanded to span the full range).
    const mergeBounds = (
      row: number,
      col: number,
    ): { x: number; y: number; w: number; h: number } | null => {
      const anchorKey = addrKey({ sheet: data.sheetIndex, row, col });
      const range = merges.byAnchor.get(anchorKey);
      if (!range) return null;
      const tl = cellRectIn(layout, cols, rows, range.r0, range.c0);
      // Sum widths/heights across the merge in a clamp-safe way.
      let x = tl.x;
      let y = tl.y;
      let w = 0;
      let h = 0;
      for (let cc = range.c0; cc <= range.c1; cc += 1) {
        if (cols.positionAt.has(cc)) {
          const r2 = cellRectIn(layout, cols, rows, row, cc);
          if (cc === range.c0) x = r2.x;
          w = r2.x + r2.w - x;
        }
      }
      for (let rr = range.r0; rr <= range.r1; rr += 1) {
        if (rows.positionAt.has(rr)) {
          const r2 = cellRectIn(layout, cols, rows, rr, col);
          if (rr === range.r0) y = r2.y;
          h = r2.y + r2.h - y;
        }
      }
      return { x, y, w, h };
    };

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
        const isMergeAnchor = merges.byAnchor.has(key);
        const spark = sparklines.get(key);
        // Render-worthy when there's data, a merge anchor, a static fill, or a
        // sparkline host — all four reasons to paint into an otherwise blank cell.
        if (!cell && !isMergeAnchor && !fmt?.fill && !spark) continue;
        const bounds = mergeBounds(r, c) ?? cellRectIn(layout, cols, rows, r, c);
        const isActive = r === active.row && c === active.col;
        const isInRange = inRange({ sheet: data.sheetIndex, row: r, col: c }, rng);

        const overlay = conditional.get(key);
        const effectiveFmt: typeof fmt =
          overlay && (overlay.fill || overlay.color || overlay.bold || overlay.italic)
            ? {
                ...fmt,
                fill: overlay.fill ?? fmt?.fill,
                color: overlay.color ?? fmt?.color,
                bold: overlay.bold || fmt?.bold,
                italic: overlay.italic || fmt?.italic,
                underline: overlay.underline || fmt?.underline,
                strike: overlay.strike || fmt?.strike,
              }
            : fmt;

        const value: CellValue = cell?.value ?? { kind: 'blank' };
        const formula = cell?.formula ?? null;
        const displayOverride =
          this.getDisplay?.({ sheet: data.sheetIndex, row: r, col: c }, value, formula, fmt) ??
          null;
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
          displayOverride,
        };

        // Static format fill OR overlay fill — both painted via paintCellFill,
        // which reads `format.fill`. The merged effectiveFmt already prefers
        // overlay over static, so a single call yields the right result.
        if (effectiveFmt?.fill) paintCellFill(paintCtx);
        if (isActive && !effectiveFmt?.fill) paintCellBackground(paintCtx);
        if (overlay?.bar !== undefined && overlay.barColor) {
          ctx.save();
          ctx.fillStyle = overlay.barColor;
          ctx.globalAlpha = 0.45;
          const w = bounds.w * overlay.bar;
          ctx.fillRect(bounds.x, bounds.y + 1, w, bounds.h - 2);
          ctx.restore();
        }
        // Icon-set: paint glyph in left gutter and shift text right by the
        // gutter width so the value reads cleanly next to the icon.
        if (overlay?.iconKind && overlay.iconSlot !== undefined) {
          paintConditionalIcon(ctx, bounds, overlay.iconKind, overlay.iconSlot);
          const insetBounds = {
            x: bounds.x + CONDITIONAL_ICON_GUTTER,
            y: bounds.y,
            w: bounds.w - CONDITIONAL_ICON_GUTTER,
            h: bounds.h,
          };
          paintCellText({ ...paintCtx, bounds: insetBounds });
        } else {
          paintCellText(paintCtx);
        }
        if (spark) paintCellSparkline(ctx, bounds, spark, state, this.getWb());
        if (fmt?.comment) paintCommentMarker(ctx, bounds);
        // Lock-icon overlay — only when the sheet is protected AND the cell
        // is explicitly unlocked, signalling which cells the user can still
        // type into despite the protection flag.
        if (sheetProtected && fmt?.locked === false) paintLockMarker(ctx, bounds, theme);

        // Error / validation triangles. Error wins over validation when both
        // would apply (an error-kind value already implies the data is bad —
        // the green triangle conveys that without piling on a red one). The
        // ignoredErrors set suppresses both kinds for the cell once the user
        // dismisses the popover via the "Ignore" action.
        const cellAddr = { sheet: data.sheetIndex, row: r, col: c };
        const cellKey = key;
        if (!ignored.has(cellKey)) {
          if (detectErrorKind(value)) {
            const hit = paintErrorTriangle(ctx, bounds, ERROR_TRIANGLE_COLOR);
            triangleHits.push({ rect: hit, addr: cellAddr, kind: 'error' });
          } else if (
            fmt?.validation &&
            detectValidationViolation(value, fmt.validation, getResolver())
          ) {
            const hit = paintValidationTriangle(ctx, bounds, VALIDATION_TRIANGLE_COLOR);
            triangleHits.push({ rect: hit, addr: cellAddr, kind: 'validation' });
          }
        }
      }
    }
    setErrorTriangleHits(triangleHits);
  }

  private paintBorders(
    state: State,
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
        const bounds = cellRectIn(layout, cols, rows, r, c);
        paintCellBorders({
          ctx,
          theme,
          bounds,
          value: { kind: 'blank' },
          formula: null,
          isActive: false,
          isInRange: false,
          format: f,
        });
      }
    }
  }

  private paintHeaders(
    state: State,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const ctx = this.ctx;
    const { layout, selection } = state;
    const active = selection.active;

    const ox = gridOriginX(layout);
    const oy = gridOriginY(layout);
    const labelTopY = layout.outlineColGutter;
    const labelLeftX = layout.outlineRowGutter;

    ctx.fillStyle = theme.bgRail;
    ctx.fillRect(0, 0, ox, oy);

    ctx.fillStyle = theme.bgRail;
    ctx.fillRect(ox, 0, this.cssWidth - ox, oy);
    ctx.fillRect(0, oy, ox, this.cssHeight - oy);

    ctx.strokeStyle = theme.ruleStrong;
    ctx.lineWidth = 1 / this.dpr;
    const align = 0.5 / this.dpr;
    ctx.beginPath();
    ctx.moveTo(0, Math.round(oy) + align);
    ctx.lineTo(this.cssWidth, Math.round(oy) + align);
    ctx.moveTo(Math.round(ox) + align, 0);
    ctx.lineTo(Math.round(ox) + align, this.cssHeight);
    ctx.stroke();

    const firstRow = rows.visible[0] ?? 0;
    const firstCol = cols.visible[0] ?? 0;

    ctx.font = `500 ${theme.textHeader}px ${theme.fontMono}`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';
    const r1c1 = state.ui.r1c1 === true;
    const fr = state.ui.filterRange;
    for (const c of cols.visible) {
      const rect = cellRectIn(layout, cols, rows, firstRow, c);
      const w = cols.sizeAt.get(c) ?? 0;
      const isActiveCol = c === active.col;
      if (isActiveCol) {
        ctx.fillStyle = theme.bgHeader;
        ctx.fillRect(rect.x, labelTopY, w, layout.headerRowHeight);
      }
      ctx.fillStyle = isActiveCol ? theme.headerFgActive : theme.headerFg;
      const label = r1c1 ? `C${c + 1}` : colLabel(c);
      ctx.fillText(label, rect.x + w / 2, labelTopY + layout.headerRowHeight / 2 + 0.5);

      // Autofilter chevron — small ▼ in the right edge of the header for any
      // column inside the active filter range.
      if (fr && c >= fr.c0 && c <= fr.c1 && w >= 28) {
        const btnRight = rect.x + w - 4;
        const btnLeft = btnRight - 14;
        const cy = labelTopY + layout.headerRowHeight / 2;
        const filterActive = state.layout.hiddenRows.size > 0;
        ctx.save();
        ctx.fillStyle = filterActive ? theme.accent : theme.bgHeader;
        ctx.strokeStyle = filterActive ? theme.accent : theme.ruleStrong;
        ctx.lineWidth = 1 / this.dpr;
        const radius = 2;
        const bx = btnLeft;
        const by = cy - 7;
        const bw = btnRight - btnLeft;
        const bh = 14;
        // Subtle background pill.
        ctx.beginPath();
        ctx.moveTo(bx + radius, by);
        ctx.lineTo(bx + bw - radius, by);
        ctx.quadraticCurveTo(bx + bw, by, bx + bw, by + radius);
        ctx.lineTo(bx + bw, by + bh - radius);
        ctx.quadraticCurveTo(bx + bw, by + bh, bx + bw - radius, by + bh);
        ctx.lineTo(bx + radius, by + bh);
        ctx.quadraticCurveTo(bx, by + bh, bx, by + bh - radius);
        ctx.lineTo(bx, by + radius);
        ctx.quadraticCurveTo(bx, by, bx + radius, by);
        ctx.closePath();
        if (filterActive) ctx.fill();
        else ctx.stroke();
        // Chevron triangle.
        ctx.fillStyle = filterActive ? theme.bg : theme.headerFg;
        ctx.beginPath();
        const tx = bx + bw / 2;
        const ty = cy + 1;
        ctx.moveTo(tx - 3.5, ty - 2);
        ctx.lineTo(tx + 3.5, ty - 2);
        ctx.lineTo(tx, ty + 2);
        ctx.closePath();
        ctx.fill();
        ctx.restore();
      }
    }

    ctx.textAlign = 'right';
    for (const r of rows.visible) {
      const rect = cellRectIn(layout, cols, rows, r, firstCol);
      const h = rows.sizeAt.get(r) ?? 0;
      const isActiveRow = r === active.row;
      if (isActiveRow) {
        ctx.fillStyle = theme.bgHeader;
        ctx.fillRect(labelLeftX, rect.y, layout.headerColWidth, h);
      }
      ctx.fillStyle = isActiveRow ? theme.headerFgActive : theme.headerFg;
      const rowLabel = r1c1 ? `R${r + 1}` : String(r + 1);
      ctx.fillText(rowLabel, ox - 8, rect.y + h / 2 + 0.5);
    }

    if (cols.positionAt.has(active.col)) {
      const aRect = cellRectIn(layout, cols, rows, firstRow, active.col);
      const w = cols.sizeAt.get(active.col) ?? 0;
      ctx.strokeStyle = theme.accent;
      ctx.lineWidth = 1.5 / this.dpr;
      ctx.beginPath();
      ctx.moveTo(aRect.x, oy - 0.5);
      ctx.lineTo(aRect.x + w, oy - 0.5);
      ctx.stroke();
    }
    if (rows.positionAt.has(active.row)) {
      const aRect = cellRectIn(layout, cols, rows, active.row, firstCol);
      const h = rows.sizeAt.get(active.row) ?? 0;
      ctx.strokeStyle = theme.accent;
      ctx.lineWidth = 1.5 / this.dpr;
      ctx.beginPath();
      ctx.moveTo(ox - 0.5, aRect.y);
      ctx.lineTo(ox - 0.5, aRect.y + h);
      ctx.stroke();
    }

    // Bracket gutters for outline groups.
    setOutlineToggles(
      paintOutlineGutters(this.ctx, state, theme, cols, rows, this.cssWidth, this.cssHeight),
    );
  }

  private paintFreezeDividers(
    state: State,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const { layout } = state;
    if (layout.freezeRows === 0 && layout.freezeCols === 0) return;
    const ctx = this.ctx;
    ctx.strokeStyle = theme.ruleStrong;
    ctx.lineWidth = 1.5 / this.dpr;
    const align = 0.5 / this.dpr;
    ctx.beginPath();
    if (layout.freezeRows > 0) {
      const yy = Math.round(layout.headerRowHeight + rows.frozenTotal) + align;
      ctx.moveTo(0, yy);
      ctx.lineTo(this.cssWidth, yy);
    }
    if (layout.freezeCols > 0) {
      const xx = Math.round(layout.headerColWidth + cols.frozenTotal) + align;
      ctx.moveTo(xx, 0);
      ctx.lineTo(xx, this.cssHeight);
    }
    ctx.stroke();
  }

  private paintSpills(
    state: State,
    theme: ResolvedTheme,
    _cols: AxisLayout,
    _rows: AxisLayout,
  ): void {
    const { layout, viewport, data } = state;
    const ranges = findSpillRanges(data.cells, data.sheetIndex);
    for (const r of ranges) {
      for (const rect of rangeRects(layout, viewport, r)) {
        paintSpillOutline(this.ctx, rect, theme);
      }
    }
  }

  private paintActive(
    state: State,
    theme: ResolvedTheme,
    cols: AxisLayout,
    rows: AxisLayout,
  ): void {
    const { layout, viewport, selection, ui, format, data } = state;
    const a = selection.active;
    setFillHandleRect(null);
    setValidationChevron(null);

    if (isRowVisible(layout, viewport, a.row) && isColVisible(layout, viewport, a.col)) {
      const bounds: Rect = cellRectIn(layout, cols, rows, a.row, a.col);
      paintActiveCellOutline(this.ctx, bounds, theme);
      const fmt = format.formats.get(addrKey({ sheet: data.sheetIndex, row: a.row, col: a.col }));
      if (fmt?.validation?.kind === 'list') {
        const rect = paintValidationChevron(this.ctx, bounds, theme);
        setValidationChevron({ rect, row: a.row, col: a.col });
      }
    }

    const preview = ui.fillPreview;
    if (preview) {
      const rects = rangeRects(layout, viewport, preview);
      for (const r of rects) paintFillPreview(this.ctx, r, theme);
    }

    const r = selection.range;
    if (isRowVisible(layout, viewport, r.r1) && isColVisible(layout, viewport, r.c1)) {
      const cornerCell = cellRectIn(layout, cols, rows, r.r1, r.c1);
      setFillHandleRect(paintFillHandle(this.ctx, cornerCell, theme));
    }
  }

  /** Top-layer trace-arrow overlay. Iterates `state.traces.items` and paints
   *  one dot + arrow per entry. Same-sheet only; entries on a different sheet
   *  than `data.sheetIndex` are silently skipped. Off-screen endpoints are
   *  also skipped — partial visibility is not handled in v1 (Excel clips
   *  arrows the same way against the freeze divider). */
  private paintTraces(state: State, cols: AxisLayout, rows: AxisLayout): void {
    const items = state.traces.items;
    if (items.length === 0) return;
    const sheet = state.data.sheetIndex;
    const ctx = this.ctx;
    const { layout, viewport } = state;
    for (const item of items) {
      if (item.from.sheet !== sheet || item.to.sheet !== sheet) continue;
      if (!isRowVisible(layout, viewport, item.from.row)) continue;
      if (!isColVisible(layout, viewport, item.from.col)) continue;
      if (!isRowVisible(layout, viewport, item.to.row)) continue;
      if (!isColVisible(layout, viewport, item.to.col)) continue;
      const fromRect = cellRectIn(layout, cols, rows, item.from.row, item.from.col);
      const toRect = cellRectIn(layout, cols, rows, item.to.row, item.to.col);
      const color = item.kind === 'precedent' ? TRACE_PRECEDENT_COLOR : TRACE_DEPENDENT_COLOR;
      paintTraceDot(ctx, fromRect, color);
      paintTraceArrow(ctx, fromRect, toRect, color);
    }
  }
}
