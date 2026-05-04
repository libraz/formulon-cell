import { formatNumber } from '../commands/format.js';
import { isColGroupCollapsed, isRowGroupCollapsed } from '../commands/outline.js';
import { REF_HIGHLIGHT_COLORS } from '../commands/refs.js';
import type { CellValue } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { CellFormat, Sparkline, State } from '../store/store.js';
import type { ResolvedTheme } from '../theme/resolve.js';
import { type AxisLayout, type Rect, gridOriginX, gridOriginY } from './geometry.js';

/** Trace-arrow accent colors. Excel paints precedents in blue and dependents
 *  in a slightly redder hue so the two relations are visually distinct even
 *  when both are active simultaneously. */
export const TRACE_PRECEDENT_COLOR = '#1f7ae0';
export const TRACE_DEPENDENT_COLOR = '#cf3a4c';

/** Paint a small filled dot at the source cell of a trace arrow. The dot
 *  sits at the center of `rect`. Mirrors Excel's blue/red round endpoint. */
export function paintTraceDot(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  color: string,
  radius = 3,
): void {
  ctx.save();
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.arc(rect.x + rect.w / 2, rect.y + rect.h / 2, radius, 0, Math.PI * 2);
  ctx.fill();
  ctx.restore();
}

/** Paint a 1px line from `fromRect` center to `toRect` center, capped with
 *  a small filled triangle at the destination. The arrow head points at
 *  `toRect`'s center so the user reads the direction of dependency. */
export function paintTraceArrow(
  ctx: CanvasRenderingContext2D,
  fromRect: Rect,
  toRect: Rect,
  color: string,
): void {
  const sx = fromRect.x + fromRect.w / 2;
  const sy = fromRect.y + fromRect.h / 2;
  const tx = toRect.x + toRect.w / 2;
  const ty = toRect.y + toRect.h / 2;
  const dx = tx - sx;
  const dy = ty - sy;
  const len = Math.hypot(dx, dy);
  if (len < 0.5) return;
  ctx.save();
  ctx.strokeStyle = color;
  ctx.fillStyle = color;
  ctx.lineWidth = 1;
  ctx.setLineDash([]);
  // Pull back slightly so the line tip butts up against the arrow head's
  // base rather than poking through it.
  const headLen = 9;
  const headHalfWidth = 4;
  const ux = dx / len;
  const uy = dy / len;
  const baseX = tx - ux * headLen;
  const baseY = ty - uy * headLen;
  ctx.beginPath();
  ctx.moveTo(sx, sy);
  ctx.lineTo(baseX, baseY);
  ctx.stroke();
  // Arrow head — filled triangle perpendicular to the line direction.
  const px = -uy;
  const py = ux;
  ctx.beginPath();
  ctx.moveTo(tx, ty);
  ctx.lineTo(baseX + px * headHalfWidth, baseY + py * headHalfWidth);
  ctx.lineTo(baseX - px * headHalfWidth, baseY - py * headHalfWidth);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

/** Paint a colored dashed border around a referenced range while a formula
 *  is being edited. Used for "trace precedents while typing" affordance. */
export function paintRefHighlight(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  colorIndex: number,
): void {
  const color = REF_HIGHLIGHT_COLORS[colorIndex % REF_HIGHLIGHT_COLORS.length] ?? '#1f7ae0';
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 2;
  ctx.setLineDash([5, 3]);
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}

/** Paint a small ▼ chevron at the right edge of the cell to indicate the
 *  cell has a list validation. Returned rect is the click hit-area. */
export function paintValidationChevron(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): Rect {
  const w = 18;
  const h = Math.min(bounds.h, 22);
  const x = bounds.x + bounds.w - w;
  const y = bounds.y + (bounds.h - h) / 2;
  ctx.save();
  ctx.fillStyle = theme.bgRail;
  ctx.fillRect(x, y, w, h);
  ctx.strokeStyle = theme.rule;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  ctx.fillStyle = theme.fg;
  ctx.beginPath();
  const cx = x + w / 2;
  const cy = y + h / 2;
  ctx.moveTo(cx - 4, cy - 2);
  ctx.lineTo(cx + 4, cy - 2);
  ctx.lineTo(cx, cy + 3);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
  return { x, y, w, h };
}

/** Side of the triangle hot-zone painted by `paintErrorTriangle` /
 *  `paintValidationTriangle`. Click hit-tests in `error-menu` use the same
 *  constant so the visual and clickable areas line up. */
export const ERROR_TRIANGLE_SIZE = 6;

/** Paint a small filled triangle in the upper-LEFT of the cell. Used for
 *  formula-error and data-validation indicators. The triangle sits flush
 *  against the top and left cell edges and points down-right. Returned rect
 *  is the click hit-area (slightly padded so the corner is comfortable to
 *  hit). */
export function paintErrorTriangle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  color: string,
): Rect {
  const size = ERROR_TRIANGLE_SIZE;
  const x = bounds.x;
  const y = bounds.y;
  ctx.save();
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.moveTo(x, y);
  ctx.lineTo(x + size, y);
  ctx.lineTo(x, y + size);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
  return { x, y, w: size, h: size };
}

/** Convenience alias — Excel paints DV violations the same shape as formula
 *  errors but red. The grid wires both through `paintErrorTriangle`; this
 *  thin wrapper keeps the call sites self-documenting. */
export function paintValidationTriangle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  color = '#d24545',
): Rect {
  return paintErrorTriangle(ctx, bounds, color);
}

/** Paint a small filled triangle in the upper-right of the cell to indicate
 *  an attached comment (Excel/Sheets convention). */
export function paintCommentMarker(ctx: CanvasRenderingContext2D, bounds: Rect): void {
  const size = 5;
  const x = bounds.x + bounds.w - size;
  const y = bounds.y;
  ctx.save();
  ctx.fillStyle = '#d24545';
  ctx.beginPath();
  ctx.moveTo(x + size, y);
  ctx.lineTo(x + size, y + size);
  ctx.lineTo(x, y);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

export interface CellPaintCtx {
  ctx: CanvasRenderingContext2D;
  theme: ResolvedTheme;
  bounds: Rect;
  value: CellValue;
  formula: string | null;
  isActive: boolean;
  isInRange: boolean;
  format?: CellFormat;
  /** When true and `formula` is non-null, paint the formula text instead of
   *  the evaluated value (Excel "Show Formulas" mode). */
  showFormulas?: boolean;
  /** Override the displayed string. Set by `paintCells` after consulting
   *  the cell registry (`inst.cells.registerFormatter`). When non-null
   *  the formatter wins over numFmt + default `formatCell`. Empty string
   *  is honored — render-blank-cell-still-padded scenarios. */
  displayOverride?: string | null;
}

/* ---- Base painters wired in MS-A. Future slots (formatting, CF, DV,
 * comments, hyperlinks) live next to these, capability-gated. ---- */

export function paintCellBackground({
  ctx,
  bounds,
  theme,
  isInRange,
  isActive,
}: CellPaintCtx): void {
  if (isActive) {
    ctx.fillStyle = theme.bgElev;
    ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
    return;
  }
  if (isInRange) {
    ctx.fillStyle = theme.accentSoft;
    ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  }
  // else: grid bg painted globally by paintGridSurface; do nothing.
}

/** User-set background fill — drawn above range tint and active highlight so
 *  cells with explicit fill stay visually filled even when selected. */
export function paintCellFill({ ctx, bounds, format }: CellPaintCtx): void {
  const fill = format?.fill;
  if (!fill) return;
  ctx.fillStyle = fill;
  ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
}

export function paintCellText({
  ctx,
  bounds,
  theme,
  value,
  formula,
  format,
  showFormulas,
  displayOverride,
}: CellPaintCtx): void {
  if (value.kind === 'blank' && !formula && displayOverride == null) return;

  const padX = 7;
  const padY = 4;
  let text: string;
  if (displayOverride != null) {
    text = displayOverride;
  } else if (showFormulas && formula) {
    text = formula;
  } else {
    text =
      value.kind === 'number' && format?.numFmt
        ? formatNumber(value.value, format.numFmt)
        : formatCell(value);
  }
  if (!text) return;

  const isNumeric = value.kind === 'number';
  const isError = value.kind === 'error';
  const isBool = value.kind === 'bool';
  const isFormula = formula != null;

  const weight = format?.bold ? 700 : isNumeric || isError || isBool ? 500 : 400;
  const styleSlant = format?.italic ? 'italic ' : '';
  const fontSize = format?.fontSize ?? theme.textCell;
  const fontFamily =
    format?.fontFamily ?? (isNumeric || isError || isFormula ? theme.fontMono : theme.fontUi);
  ctx.font = `${styleSlant}${weight} ${fontSize}px ${fontFamily}`;
  const isHyperlink = !!format?.hyperlink;
  ctx.fillStyle = format?.color
    ? format.color
    : isHyperlink
      ? theme.accent
      : isError
        ? theme.cellErrorFg
        : isBool
          ? theme.cellBoolFg
          : isNumeric
            ? theme.cellNumFg
            : theme.fg;

  let align: CanvasTextAlign;
  if (format?.align) {
    align = format.align;
  } else {
    align = isNumeric || isBool || isError ? 'right' : 'left';
  }
  const indentPx = (format?.indent ?? 0) * 8;
  const rotation = format?.rotation ?? 0;
  const wrap = !!format?.wrap;

  ctx.save();
  ctx.beginPath();
  ctx.rect(bounds.x, bounds.y, bounds.w, bounds.h);
  ctx.clip();

  // Rotated text — render around cell center, ignore wrap/indent for
  // simplicity. (Excel is more elaborate but this covers ±90° common case.)
  if (rotation !== 0) {
    ctx.translate(bounds.x + bounds.w / 2, bounds.y + bounds.h / 2);
    ctx.rotate((rotation * Math.PI) / 180);
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(text, 0, 0);
    ctx.restore();
    return;
  }

  // Word-wrap path — split into lines, paint each.
  if (wrap) {
    const lines = wrapText(ctx, text, bounds.w - padX * 2 - indentPx);
    const lineH = Math.round(fontSize * 1.25);
    const totalH = lineH * lines.length;
    const vAlign = format?.vAlign ?? 'bottom';
    let startY: number;
    if (vAlign === 'top') startY = bounds.y + padY + lineH * 0.5;
    else if (vAlign === 'middle') startY = bounds.y + (bounds.h - totalH) / 2 + lineH * 0.5;
    else startY = bounds.y + bounds.h - padY - totalH + lineH * 0.5;
    ctx.textBaseline = 'middle';
    ctx.textAlign = align;
    let tx: number;
    if (align === 'right') tx = bounds.x + bounds.w - padX;
    else if (align === 'center') tx = bounds.x + bounds.w / 2;
    else tx = bounds.x + padX + indentPx;
    for (let i = 0; i < lines.length; i += 1) {
      ctx.fillText(lines[i] ?? '', tx, startY + i * lineH);
    }
    ctx.restore();
    return;
  }

  // Single-line path with vertical alignment.
  ctx.textBaseline = 'middle';
  ctx.textAlign = align;
  let tx: number;
  if (align === 'right') tx = bounds.x + bounds.w - padX;
  else if (align === 'center') tx = bounds.x + bounds.w / 2;
  else tx = bounds.x + padX + indentPx;

  const vAlign = format?.vAlign ?? 'bottom';
  let ty: number;
  if (vAlign === 'top') ty = bounds.y + padY + fontSize * 0.6;
  else if (vAlign === 'bottom') ty = bounds.y + bounds.h - padY - fontSize * 0.45;
  else ty = bounds.y + bounds.h / 2 + 0.5;

  ctx.fillText(text, tx, ty);

  if (format?.underline || format?.strike || isHyperlink) {
    const metrics = ctx.measureText(text);
    const w = metrics.width;
    let lineX0: number;
    if (align === 'right') lineX0 = tx - w;
    else if (align === 'center') lineX0 = tx - w / 2;
    else lineX0 = tx;
    ctx.strokeStyle = ctx.fillStyle as string;
    ctx.lineWidth = 1;
    if (format?.underline || isHyperlink) {
      const uy = Math.round(ty + fontSize * 0.45) + 0.5;
      ctx.beginPath();
      ctx.moveTo(lineX0, uy);
      ctx.lineTo(lineX0 + w, uy);
      ctx.stroke();
    }
    if (format?.strike) {
      const sy = Math.round(ty) + 0.5;
      ctx.beginPath();
      ctx.moveTo(lineX0, sy);
      ctx.lineTo(lineX0 + w, sy);
      ctx.stroke();
    }
  }
  ctx.restore();
}

/** Greedy word-wrap. Splits on whitespace, packs as many words per line as
 *  fit within `maxWidth`. Long un-breakable words are emitted as a single
 *  overflowing line rather than mid-word breaking. */
function wrapText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  if (maxWidth <= 0) return [text];
  // Honor explicit \n line breaks (Alt+Enter in Excel).
  const paragraphs = text.split('\n');
  const out: string[] = [];
  for (const para of paragraphs) {
    const words = para.split(/(\s+)/);
    let line = '';
    for (const word of words) {
      const candidate = line + word;
      if (ctx.measureText(candidate).width <= maxWidth || line === '') {
        line = candidate;
      } else {
        out.push(line.trimEnd());
        line = word.trimStart();
      }
    }
    if (line) out.push(line);
    if (para === '') out.push('');
  }
  return out.length > 0 ? out : [''];
}

/** Per-cell border lines. Drawn after gridlines + cell text but before the
 *  active outline so user borders sit above the hairline grid. Each side may
 *  be a boolean (legacy thin solid) or an object with `{style, color}`. */
export function paintCellBorders({ ctx, bounds, theme, format }: CellPaintCtx): void {
  const sides = format?.borders;
  if (!sides) return;
  const drawSide = (
    side: typeof sides.top,
    x0: number,
    y0: number,
    x1: number,
    y1: number,
  ): void => {
    if (!side) return;
    const cfg = typeof side === 'object' ? side : { style: 'thin' as const };
    const color = (typeof side === 'object' && side.color) || theme.ruleStrong || theme.fg;
    const widthMap: Record<string, number> = {
      thin: 1,
      medium: 1.6,
      thick: 2.5,
      dashed: 1,
      dotted: 1,
      double: 1,
      hair: 0.5,
      mediumDashed: 1.6,
      dashDot: 1,
      mediumDashDot: 1.6,
      dashDotDot: 1,
      mediumDashDotDot: 1.6,
      slantDashDot: 1,
    };
    const dashMap: Record<string, number[]> = {
      thin: [],
      medium: [],
      thick: [],
      dashed: [4, 3],
      dotted: [1, 2],
      double: [],
      hair: [1, 1],
      mediumDashed: [6, 3],
      dashDot: [4, 2, 1, 2],
      mediumDashDot: [6, 2, 2, 2],
      dashDotDot: [4, 2, 1, 2, 1, 2],
      mediumDashDotDot: [6, 2, 2, 2, 2, 2],
      slantDashDot: [5, 2, 1, 2],
    };
    ctx.save();
    ctx.strokeStyle = color;
    ctx.lineWidth = widthMap[cfg.style] ?? 1;
    ctx.setLineDash(dashMap[cfg.style] ?? []);
    if (cfg.style === 'double') {
      // Two parallel lines, 2px apart.
      const off = 1.5;
      const dx = y0 === y1 ? 0 : 1;
      const dy = y0 === y1 ? 1 : 0;
      ctx.beginPath();
      ctx.moveTo(x0 + dx * off, y0 + dy * off);
      ctx.lineTo(x1 + dx * off, y1 + dy * off);
      ctx.moveTo(x0 - dx * off, y0 - dy * off);
      ctx.lineTo(x1 - dx * off, y1 - dy * off);
      ctx.stroke();
    } else {
      ctx.beginPath();
      ctx.moveTo(x0, y0);
      ctx.lineTo(x1, y1);
      ctx.stroke();
    }
    ctx.restore();
  };
  const yt = Math.round(bounds.y) + 0.5;
  const yb = Math.round(bounds.y + bounds.h) - 0.5;
  const xl = Math.round(bounds.x) + 0.5;
  const xr = Math.round(bounds.x + bounds.w) - 0.5;
  drawSide(sides.top, bounds.x, yt, bounds.x + bounds.w, yt);
  drawSide(sides.bottom, bounds.x, yb, bounds.x + bounds.w, yb);
  drawSide(sides.left, xl, bounds.y, xl, bounds.y + bounds.h);
  drawSide(sides.right, xr, bounds.y, xr, bounds.y + bounds.h);
  drawSide(sides.diagonalDown, bounds.x, bounds.y, bounds.x + bounds.w, bounds.y + bounds.h);
  drawSide(sides.diagonalUp, bounds.x, bounds.y + bounds.h, bounds.x + bounds.w, bounds.y);
}

/** Active cell outline. Drawn in a separate pass after all cell text so the
 *  outline never gets clipped by neighbouring cell rects.
 *
 *  Excel 365 paints the active cell with a crisp ~2px green border and no
 *  inner fill. We keep a faint outer halo (1px ring of accentSoft) so the
 *  outline reads cleanly against any cell background. */
export function paintActiveCellOutline(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): void {
  ctx.fillStyle = theme.accentSoft;
  ctx.fillRect(bounds.x - 2, bounds.y - 2, bounds.w + 4, bounds.h + 4);

  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 2;
  ctx.strokeRect(bounds.x, bounds.y, bounds.w, bounds.h);
}

/** Visible side length of the fill handle in CSS pixels. Excel 365 uses a
 *  small accent-coloured square at the bottom-right of the active selection
 *  range; the user grabs it to drag-fill into adjacent cells. */
export const FILL_HANDLE_SIZE = 6;

/** Excel-style fill handle. Drawn at the selection range's bottom-right
 *  corner (or active-cell corner when selection is a single cell). The square
 *  is filled in `theme.accent` (resolves from `--fc-accent`, falls back to
 *  `#0078d4` when unset) and surrounded by a 1px white border so it stands
 *  proud against any cell fill. The returned rect spans the white border so
 *  the pointer layer hit-tests against the visible area, not just the inner
 *  filled square. */
export function paintFillHandle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): Rect {
  const hs = FILL_HANDLE_SIZE;
  // Centre the visible square on the cell's bottom-right corner so half the
  // handle bleeds outside the selection — matches Excel.
  const x = bounds.x + bounds.w - hs / 2;
  const y = bounds.y + bounds.h - hs / 2;
  const accent = theme.accent || '#0078d4';
  // 1px white border ring. Painted as a slightly larger white square first;
  // the accent fill on top leaves a 1px ring exposed.
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(x - 1, y - 1, hs + 2, hs + 2);
  ctx.fillStyle = accent;
  ctx.fillRect(x, y, hs, hs);
  return { x: x - 1, y: y - 1, w: hs + 2, h: hs + 2 };
}

/** "Marching ants" outline for the fill drag preview. We omit the actual
 *  ant animation to keep the renderer stateless; a dashed strong rule is
 *  enough to read as a destination. */
export function paintFillPreview(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): void {
  ctx.save();
  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 1.5;
  ctx.setLineDash([4, 3]);
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}

/** Outline a dynamic-array spill range — Excel paints a 1px accent ring
 *  around the spilled rectangle so it reads as a single result. */
export function paintSpillOutline(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): void {
  ctx.save();
  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 1;
  ctx.setLineDash([]);
  ctx.globalAlpha = 0.65;
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}

/** One slot per outline level. The bracket spine sits at slot center. */
export const OUTLINE_BRACKET_SLOT = 14;
const TOGGLE_SIZE = 11;

/** Paint a small +/- toggle button. (cx, cy) is the center. */
function paintToggle(
  ctx: CanvasRenderingContext2D,
  theme: ResolvedTheme,
  cx: number,
  cy: number,
  collapsed: boolean,
): Rect {
  const half = TOGGLE_SIZE / 2;
  const x = Math.round(cx - half);
  const y = Math.round(cy - half);
  ctx.save();
  ctx.fillStyle = theme.bg;
  ctx.fillRect(x, y, TOGGLE_SIZE, TOGGLE_SIZE);
  ctx.strokeStyle = theme.ruleStrong;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, TOGGLE_SIZE - 1, TOGGLE_SIZE - 1);
  ctx.strokeStyle = theme.fg;
  ctx.lineWidth = 1;
  ctx.beginPath();
  // Horizontal stroke (always drawn)
  ctx.moveTo(x + 2.5, y + half + 0.5);
  ctx.lineTo(x + TOGGLE_SIZE - 2.5, y + half + 0.5);
  if (collapsed) {
    // Vertical stroke for `+`
    ctx.moveTo(x + half + 0.5, y + 2.5);
    ctx.lineTo(x + half + 0.5, y + TOGGLE_SIZE - 2.5);
  }
  ctx.stroke();
  ctx.restore();
  return { x, y, w: TOGGLE_SIZE, h: TOGGLE_SIZE };
}

export interface OutlineToggleHit {
  axis: 'row' | 'col';
  level: number;
  /** First/last index of the contiguous run this toggle controls. */
  i0: number;
  i1: number;
  rect: Rect;
}

/** Walk both gutters and paint brackets + toggles. Returns toggle hit-rects so
 *  the pointer layer can route clicks to collapse/expand commands. */
export function paintOutlineGutters(
  ctx: CanvasRenderingContext2D,
  state: State,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
  cssWidth: number,
  cssHeight: number,
): OutlineToggleHit[] {
  const hits: OutlineToggleHit[] = [];
  const layout = state.layout;
  const ox = gridOriginX(layout);
  const oy = gridOriginY(layout);

  // ── Row gutter (left of row-number column).
  if (layout.outlineRowGutter > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, layout.outlineRowGutter, cssHeight);
    ctx.clip();
    let maxLvl = 0;
    for (const v of layout.outlineRows.values()) if (v > maxLvl) maxLvl = v;
    for (let lvl = 1; lvl <= maxLvl; lvl += 1) {
      const slotCx = (lvl - 1) * OUTLINE_BRACKET_SLOT + OUTLINE_BRACKET_SLOT / 2;
      let runStartIdx = -1;
      for (let i = 0; i <= rows.visible.length; i += 1) {
        const r = rows.visible[i];
        const inRun = r != null && (layout.outlineRows.get(r) ?? 0) >= lvl;
        if (inRun && runStartIdx < 0) runStartIdx = i;
        if (!inRun && runStartIdx >= 0) {
          const r0 = rows.visible[runStartIdx];
          const r1 = rows.visible[i - 1];
          if (r0 === undefined || r1 === undefined) {
            runStartIdx = -1;
            continue;
          }
          const top = oy + (rows.positionAt.get(r0) ?? 0);
          const bottomRow = rows.positionAt.get(r1) ?? 0;
          const bottomH = rows.sizeAt.get(r1) ?? 0;
          const bottom = oy + bottomRow + bottomH;
          ctx.strokeStyle = theme.rule;
          ctx.lineWidth = 1;
          ctx.beginPath();
          ctx.moveTo(slotCx + 0.5, top);
          ctx.lineTo(slotCx + 0.5, bottom);
          // Foot tick on the bottom (the "summary row" side).
          ctx.moveTo(slotCx + 0.5, bottom - 0.5);
          ctx.lineTo(slotCx + 5.5, bottom - 0.5);
          ctx.stroke();
          const collapsed = isRowGroupCollapsed(layout, r0, r1);
          // Toggle sits on the row that should remain visible — typically the
          // row just below the band on Excel default ("summary below"). If
          // there is no such row visible, render at the bottom edge of the run.
          const summaryRow = r1 + 1;
          const summaryY =
            rows.positionAt.has(summaryRow) && !layout.hiddenRows.has(summaryRow)
              ? oy + (rows.positionAt.get(summaryRow) ?? 0) + (rows.sizeAt.get(summaryRow) ?? 0) / 2
              : bottom + TOGGLE_SIZE / 2 + 2;
          const rect = paintToggle(ctx, theme, slotCx, summaryY, collapsed);
          hits.push({ axis: 'row', level: lvl, i0: r0, i1: r1, rect });
          runStartIdx = -1;
        }
      }
    }
    ctx.restore();
  }

  // ── Col gutter (above col-letter row).
  if (layout.outlineColGutter > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, cssWidth, layout.outlineColGutter);
    ctx.clip();
    let maxLvl = 0;
    for (const v of layout.outlineCols.values()) if (v > maxLvl) maxLvl = v;
    for (let lvl = 1; lvl <= maxLvl; lvl += 1) {
      const slotCy = (lvl - 1) * OUTLINE_BRACKET_SLOT + OUTLINE_BRACKET_SLOT / 2;
      let runStartIdx = -1;
      for (let i = 0; i <= cols.visible.length; i += 1) {
        const c = cols.visible[i];
        const inRun = c != null && (layout.outlineCols.get(c) ?? 0) >= lvl;
        if (inRun && runStartIdx < 0) runStartIdx = i;
        if (!inRun && runStartIdx >= 0) {
          const c0 = cols.visible[runStartIdx];
          const c1 = cols.visible[i - 1];
          if (c0 === undefined || c1 === undefined) {
            runStartIdx = -1;
            continue;
          }
          const left = ox + (cols.positionAt.get(c0) ?? 0);
          const rightCol = cols.positionAt.get(c1) ?? 0;
          const rightW = cols.sizeAt.get(c1) ?? 0;
          const right = ox + rightCol + rightW;
          ctx.strokeStyle = theme.rule;
          ctx.lineWidth = 1;
          ctx.beginPath();
          ctx.moveTo(left, slotCy + 0.5);
          ctx.lineTo(right, slotCy + 0.5);
          ctx.moveTo(right - 0.5, slotCy + 0.5);
          ctx.lineTo(right - 0.5, slotCy + 5.5);
          ctx.stroke();
          const collapsed = isColGroupCollapsed(layout, c0, c1);
          const summaryCol = c1 + 1;
          const summaryX =
            cols.positionAt.has(summaryCol) && !layout.hiddenCols.has(summaryCol)
              ? ox + (cols.positionAt.get(summaryCol) ?? 0) + (cols.sizeAt.get(summaryCol) ?? 0) / 2
              : right + TOGGLE_SIZE / 2 + 2;
          const rect = paintToggle(ctx, theme, summaryX, slotCy, collapsed);
          hits.push({ axis: 'col', level: lvl, i0: c0, i1: c1, rect });
          runStartIdx = -1;
        }
      }
    }
    ctx.restore();
  }

  return hits;
}

const DEFAULT_SPARKLINE_COLOR = '#0078d4';
const DEFAULT_NEGATIVE_COLOR = '#d24545';

/** Inline mini-chart drawn inside a single cell rect. `values` is the resolved
 *  numeric series for `spec.source`; non-numeric source cells are filtered by
 *  the caller before invoking. No-op when the series is empty. */
export function paintSparkline(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  spec: Sparkline,
  values: readonly number[],
): void {
  if (values.length === 0) return;
  const inset = 2;
  const x = rect.x + inset;
  const y = rect.y + inset;
  const w = Math.max(0, rect.w - inset * 2);
  const h = Math.max(0, rect.h - inset * 2);
  if (w <= 0 || h <= 0) return;
  const color = spec.color ?? DEFAULT_SPARKLINE_COLOR;
  const negColor = spec.negativeColor ?? DEFAULT_NEGATIVE_COLOR;

  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.globalAlpha = 1;

  if (spec.kind === 'line') {
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    for (const v of values) {
      if (v < min) min = v;
      if (v > max) max = v;
    }
    const span = max - min || 1;
    const stepX = values.length > 1 ? w / (values.length - 1) : 0;
    ctx.strokeStyle = color;
    ctx.lineWidth = 1.5;
    ctx.lineJoin = 'round';
    ctx.beginPath();
    for (let i = 0; i < values.length; i += 1) {
      const v = values[i] ?? 0;
      const px = x + (values.length > 1 ? i * stepX : w / 2);
      const py = y + h - ((v - min) / span) * h;
      if (i === 0) ctx.moveTo(px, py);
      else ctx.lineTo(px, py);
    }
    ctx.stroke();
  } else if (spec.kind === 'column') {
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    for (const v of values) {
      if (v < min) min = v;
      if (v > max) max = v;
    }
    // Anchor the baseline at zero when the range straddles it; otherwise rest
    // bars on the value extreme closest to zero so all bars stay positive-side.
    const baseline = min < 0 && max > 0 ? 0 : min >= 0 ? min : max;
    const span = Math.max(Math.abs(max - baseline), Math.abs(min - baseline)) || 1;
    const slot = w / values.length;
    const gap = Math.min(1, slot * 0.2);
    const barW = Math.max(1, slot - gap);
    const baseY = y + h - ((baseline - min) / Math.max(max - min, 1e-9)) * h;
    for (let i = 0; i < values.length; i += 1) {
      const v = values[i] ?? 0;
      const isNeg = v < 0;
      const fill = spec.showNegative && isNeg ? negColor : color;
      const px = x + i * slot;
      const barH = (Math.abs(v - baseline) / span) * h;
      const py = isNeg ? baseY : baseY - barH;
      ctx.fillStyle = fill;
      ctx.fillRect(px, py, barW, Math.max(1, barH));
    }
  } else {
    // win-loss
    const half = h / 2;
    const slot = w / values.length;
    const gap = Math.min(1, slot * 0.2);
    const barW = Math.max(1, slot - gap);
    const midY = y + half;
    for (let i = 0; i < values.length; i += 1) {
      const v = values[i] ?? 0;
      if (v === 0) continue;
      const isNeg = v < 0;
      const fill = spec.showNegative && isNeg ? negColor : color;
      ctx.fillStyle = fill;
      const px = x + i * slot;
      if (isNeg) ctx.fillRect(px, midY, barW, Math.max(1, half - 1));
      else ctx.fillRect(px, midY - Math.max(1, half - 1), barW, Math.max(1, half - 1));
    }
  }

  ctx.restore();
}
