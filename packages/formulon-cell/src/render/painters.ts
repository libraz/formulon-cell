import { formatNumber } from '../commands/format.js';
import { REF_HIGHLIGHT_COLORS } from '../commands/refs.js';
import type { CellValue } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { CellFormat, ConditionalIconSet } from '../store/store.js';
import type { ResolvedTheme } from '../theme/resolve.js';
import type { Rect } from './geometry.js';

/** Trace-arrow accent colors. Spreadsheets paint precedents in blue and dependents
 *  in a slightly redder hue so the two relations are visually distinct even
 *  when both are active simultaneously. */
export const TRACE_PRECEDENT_COLOR = '#1f7ae0';
export const TRACE_DEPENDENT_COLOR = '#cf3a4c';

/** Paint a small filled dot at the source cell of a trace arrow. The dot
 *  sits at the center of `rect`. Mirrors the blue/red round endpoint convention. */
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

/** Paint the spreadsheet Table header filter/dropdown affordance. This is visual
 *  only today; table filtering still routes through the normal filter model. */
export function paintTableHeaderChevron(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): Rect {
  const size = 14;
  const x = bounds.x + bounds.w - size - 3;
  const y = bounds.y + Math.max(2, (bounds.h - size) / 2);
  ctx.save();
  ctx.fillStyle = 'rgba(255,255,255,0.72)';
  ctx.fillRect(x, y, size, size);
  ctx.strokeStyle = theme.rule;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, size - 1, size - 1);
  ctx.fillStyle = theme.fgMute || theme.fg;
  ctx.beginPath();
  const cx = x + size / 2;
  const cy = y + size / 2 + 1;
  ctx.moveTo(cx - 3.5, cy - 2);
  ctx.lineTo(cx + 3.5, cy - 2);
  ctx.lineTo(cx, cy + 2.5);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
  return { x, y, w: size, h: size };
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

/** Convenience alias — spreadsheets paint DV violations the same shape as formula
 *  errors but red. The grid wires both through `paintErrorTriangle`; this
 *  thin wrapper keeps the call sites self-documenting. */
export function paintValidationTriangle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  color = '#d24545',
): Rect {
  return paintErrorTriangle(ctx, bounds, color);
}

/** Paint a small lock-icon overlay in the upper-right corner of a cell.
 *  Used to flag cells that are still writable when the sheet is otherwise
 *  protected (i.e. `format.locked === false` while the sheet is protected).
 *  The shape is a 7×8 rounded body + arched shackle drawn in
 *  `theme.accent` so it reads as an affordance, not an error. */
export function paintLockMarker(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): void {
  const w = 8;
  const h = 9;
  const x = Math.round(bounds.x + bounds.w - w - 2);
  const y = Math.round(bounds.y + 2);
  const color = theme.accent || '#0078d4';
  ctx.save();
  // Shackle (top arch, drawn as a stroked half-circle).
  ctx.strokeStyle = color;
  ctx.lineWidth = 1.25;
  ctx.beginPath();
  ctx.arc(x + w / 2, y + 3, 2.25, Math.PI, 0);
  ctx.stroke();
  // Body.
  ctx.fillStyle = color;
  ctx.fillRect(x + 1, y + 3, w - 2, h - 4);
  ctx.restore();
}

/** Paint a small filled triangle in the upper-right of the cell to indicate
 *  an attached comment (spreadsheet convention). */
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

/** Excel-style "General" number rendering with overflow handling. Tries the
 *  initial Intl.NumberFormat output first; if it doesn't fit, walks fractional
 *  digits down to integer, then falls through to scientific notation, and
 *  finally returns a #### filler sized to the cell. */
function fitGeneralNumberToWidth(
  ctx: CanvasRenderingContext2D,
  value: number,
  availableWidth: number,
  locale: string,
  fontSize: number,
): string {
  if (!Number.isFinite(value)) return String(value);
  const toSci = (prec: number): string =>
    value.toExponential(prec).replace(/e([+-]?)(\d+)$/i, (_m, sign: string, exp: string) => {
      const s = sign === '-' ? '-' : '+';
      return `E${s}${exp.padStart(2, '0')}`;
    });
  const abs = Math.abs(value);
  // Already in scientific range (matches formatGeneralNumber's threshold) —
  // only the exponent precision can be reduced.
  if (abs > 0 && (abs >= 1e11 || abs < 1e-9)) {
    for (let prec = 5; prec >= 0; prec -= 1) {
      const t = toSci(prec);
      if (ctx.measureText(t).width <= availableWidth) return t;
    }
  } else {
    for (let digits = 12; digits >= 0; digits -= 1) {
      const t = new Intl.NumberFormat(locale, { maximumFractionDigits: digits }).format(value);
      if (ctx.measureText(t).width <= availableWidth) return t;
    }
    for (let prec = 5; prec >= 0; prec -= 1) {
      const t = toSci(prec);
      if (ctx.measureText(t).width <= availableWidth) return t;
    }
  }
  const hashWidth = Math.max(1, ctx.measureText('#').width || fontSize * 0.62);
  return '#'.repeat(Math.max(1, Math.floor(availableWidth / hashWidth)));
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
   *  the evaluated value (the desktop-spreadsheet "Show Formulas" mode). */
  showFormulas?: boolean;
  /** Override the displayed string. Set by `paintCells` after consulting
   *  the cell registry (`inst.cells.registerFormatter`). When non-null
   *  the formatter wins over numFmt + default `formatCell`. Empty string
   *  is honored — render-blank-cell-still-padded scenarios. */
  displayOverride?: string | null;
  /** BCP 47 locale used for number/date formatting. */
  locale?: string;
}

export type TextVAlign = 'top' | 'middle' | 'bottom';

export interface TextMetricsBox {
  ascent: number;
  descent: number;
}

export const textMetricsBox = (metrics: TextMetrics, fontSize: number): TextMetricsBox => ({
  ascent:
    Number.isFinite(metrics.actualBoundingBoxAscent) && metrics.actualBoundingBoxAscent > 0
      ? metrics.actualBoundingBoxAscent
      : fontSize * 0.72,
  descent:
    Number.isFinite(metrics.actualBoundingBoxDescent) && metrics.actualBoundingBoxDescent > 0
      ? metrics.actualBoundingBoxDescent
      : fontSize * 0.22,
});

export const stableTextMetricsBox = (fontSize: number): TextMetricsBox => ({
  ascent: fontSize * 0.72,
  descent: fontSize * 0.22,
});

export const textBaselineY = (
  bounds: Rect,
  box: TextMetricsBox,
  vAlign: TextVAlign,
  padY: number,
): number => {
  if (vAlign === 'top') return bounds.y + padY + box.ascent;
  if (vAlign === 'middle') return bounds.y + (bounds.h - box.ascent - box.descent) / 2 + box.ascent;
  return bounds.y + bounds.h - padY - box.descent;
};

const fontCss = (family: string): string =>
  family
    .split(',')
    .map((part) => {
      const name = part.trim();
      if (!name) return name;
      if (name.startsWith('"') || name.startsWith("'")) return name;
      if (/^[a-zA-Z-]+$/.test(name)) return name;
      return `"${name.replaceAll('"', '\\"')}"`;
    })
    .filter(Boolean)
    .join(', ');

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

/** Width of the left-gutter that hosts conditional-format icon-set glyphs.
 *  Cell text is right-shifted by this amount when an icon overlay is
 *  present so the artwork and value don't overlap. */
export const CONDITIONAL_ICON_GUTTER = 16;

/** Paint a small icon-set glyph at the left of a cell rect. The artwork is
 *  drawn with primitive shapes — no SVG / image asset — so the renderer
 *  stays self-contained. `slot` is 0-based; the family decides the count
 *  (3 or 5) and how the slot maps to color/direction. */
export function paintConditionalIcon(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  kind: ConditionalIconSet,
  slot: number,
): void {
  const cx = rect.x + CONDITIONAL_ICON_GUTTER / 2;
  const cy = rect.y + rect.h / 2;
  const r = Math.min(5, Math.floor(rect.h * 0.32));
  ctx.save();
  if (kind === 'arrows3' || kind === 'arrows5') {
    // Map slot to a unit-vector direction. arrows3: [down, right, up];
    // arrows5: [down, down-right, right, up-right, up]. Color tracks slot:
    // low → red, mid → amber, high → green. arrows5 reuses the same ramp
    // with finer steps.
    const slots = kind === 'arrows5' ? 5 : 3;
    const idx = Math.max(0, Math.min(slots - 1, slot));
    const colors =
      kind === 'arrows5'
        ? ['#d24545', '#e07a4d', '#cfa64a', '#7fb352', '#3aa055']
        : ['#d24545', '#cfa64a', '#3aa055'];
    const color = colors[idx] ?? '#777';
    // Direction angle in radians, 0 = right, -π/2 = up, π/2 = down.
    const angles =
      kind === 'arrows5'
        ? [Math.PI / 2, Math.PI / 4, 0, -Math.PI / 4, -Math.PI / 2]
        : [Math.PI / 2, 0, -Math.PI / 2];
    const a = angles[idx] ?? 0;
    const len = r + 2;
    const dx = Math.cos(a) * len;
    const dy = Math.sin(a) * len;
    ctx.strokeStyle = color;
    ctx.fillStyle = color;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(cx - dx, cy - dy);
    ctx.lineTo(cx + dx, cy + dy);
    ctx.stroke();
    // Arrow head — small triangle perpendicular to the direction.
    const px = -Math.sin(a);
    const py = Math.cos(a);
    const headLen = 3;
    const headHalf = 2.5;
    const baseX = cx + dx - Math.cos(a) * headLen;
    const baseY = cy + dy - Math.sin(a) * headLen;
    ctx.beginPath();
    ctx.moveTo(cx + dx, cy + dy);
    ctx.lineTo(baseX + px * headHalf, baseY + py * headHalf);
    ctx.lineTo(baseX - px * headHalf, baseY - py * headHalf);
    ctx.closePath();
    ctx.fill();
  } else if (kind === 'traffic3') {
    // Three colored circles, only the slot's circle is filled solid; the
    // other two are dim outlines so the icon reads as a single state.
    const idx = Math.max(0, Math.min(2, slot));
    const colors = ['#d24545', '#cfa64a', '#3aa055'];
    const color = colors[idx] ?? '#777';
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.fill();
  } else {
    // stars3 — outlined or filled 5-pointed star. slot 0 = empty,
    // slot 1 = half (filled lower-half), slot 2 = full.
    const idx = Math.max(0, Math.min(2, slot));
    const star = (filled: boolean): void => {
      ctx.beginPath();
      const spikes = 5;
      const outerR = r;
      const innerR = r * 0.45;
      for (let i = 0; i < spikes * 2; i += 1) {
        const angle = (i * Math.PI) / spikes - Math.PI / 2;
        const radius = i % 2 === 0 ? outerR : innerR;
        const px = cx + Math.cos(angle) * radius;
        const py = cy + Math.sin(angle) * radius;
        if (i === 0) ctx.moveTo(px, py);
        else ctx.lineTo(px, py);
      }
      ctx.closePath();
      if (filled) ctx.fill();
      else ctx.stroke();
    };
    ctx.strokeStyle = '#cfa64a';
    ctx.fillStyle = '#cfa64a';
    ctx.lineWidth = 1;
    if (idx === 2) {
      star(true);
    } else if (idx === 1) {
      // Half-fill — clip the right half then fill, then outline whole.
      ctx.save();
      ctx.beginPath();
      ctx.rect(rect.x, cy - r - 1, cx - rect.x, (r + 1) * 2);
      ctx.clip();
      star(true);
      ctx.restore();
      star(false);
    } else {
      star(false);
    }
  }
  ctx.restore();
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
  locale,
}: CellPaintCtx): void {
  if (value.kind === 'blank' && !formula && displayOverride == null) return;

  const padX = 7;
  const padY = 3;
  let text: string;
  if (displayOverride != null) {
    text = displayOverride;
  } else if (showFormulas && formula) {
    text = formula;
  } else {
    text =
      value.kind === 'number' && format?.numFmt
        ? formatNumber(value.value, format.numFmt, locale)
        : formatCell(value, locale);
  }
  if (!text) return;

  const isNumeric = value.kind === 'number';
  const isError = value.kind === 'error';
  const isBool = value.kind === 'bool';
  const isFormulaDisplay = showFormulas && formula != null;
  const weight = format?.bold ? 700 : 400;
  const styleSlant = format?.italic ? 'italic ' : '';
  const fontSize = format?.fontSize ?? theme.textCell;
  const fontFamily = format?.fontFamily ?? (isFormulaDisplay ? theme.fontMono : theme.fontUi);
  ctx.font = `${styleSlant}${weight} ${fontSize}px ${fontCss(fontFamily)}`;
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
    align = isFormulaDisplay ? 'left' : isNumeric ? 'right' : isBool || isError ? 'center' : 'left';
  }
  const indentPx = (format?.indent ?? 0) * 8;
  const rotation = format?.rotation ?? 0;
  const wrap = !!format?.wrap;

  ctx.save();
  ctx.beginPath();
  ctx.rect(bounds.x, bounds.y, bounds.w, bounds.h);
  ctx.clip();

  // Rotated text — render around cell center, ignore wrap/indent for
  // simplicity. (Real desktop spreadsheets are more elaborate but this covers ±90° common case.)
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
    const lineH = Math.round(fontSize * 1.28);
    const totalH = lineH * lines.length;
    const vAlign = format?.vAlign ?? 'bottom';
    const measured = stableTextMetricsBox(fontSize);
    let startY = textBaselineY(bounds, measured, vAlign, padY);
    if (vAlign === 'middle') startY -= (totalH - lineH) / 2;
    else if (vAlign === 'bottom') startY -= (lines.length - 1) * lineH;
    ctx.textBaseline = 'alphabetic';
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
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = align;
  let tx: number;
  if (align === 'right') tx = bounds.x + bounds.w - padX;
  else if (align === 'center') tx = bounds.x + bounds.w / 2;
  else tx = bounds.x + padX + indentPx;

  const vAlign = format?.vAlign ?? 'bottom';
  const availableTextWidth = Math.max(0, bounds.w - padX * 2 - indentPx);
  if (isNumeric && ctx.measureText(text).width > availableTextWidth) {
    if (format?.numFmt) {
      // Explicit number format: Excel shows #### when the formatted value
      // can't shrink without losing the user's chosen presentation.
      const hashWidth = Math.max(1, ctx.measureText('#').width || fontSize * 0.62);
      text = '#'.repeat(Math.max(1, Math.floor(availableTextWidth / hashWidth)));
    } else if (!isFormulaDisplay && value.kind === 'number') {
      // General number: Excel progressively trims fractional digits before
      // falling back to scientific notation and finally ####. Without this,
      // right-aligned text just clips and the user sees only the trailing
      // digits — making the result look corrupted.
      text = fitGeneralNumberToWidth(
        ctx,
        value.value,
        availableTextWidth,
        locale ?? 'en-US',
        fontSize,
      );
    }
  }
  const metrics = ctx.measureText(text);
  const box = stableTextMetricsBox(fontSize);
  const ty = textBaselineY(bounds, box, vAlign, padY);

  ctx.fillText(text, tx, ty);

  if (format?.underline || format?.strike || isHyperlink) {
    const w = metrics.width;
    let lineX0: number;
    if (align === 'right') lineX0 = tx - w;
    else if (align === 'center') lineX0 = tx - w / 2;
    else lineX0 = tx;
    ctx.strokeStyle = ctx.fillStyle as string;
    ctx.lineWidth = 1;
    if (format?.underline || isHyperlink) {
      const uy = Math.round(ty + Math.max(2, box.descent * 0.55)) + 0.5;
      ctx.beginPath();
      ctx.moveTo(lineX0, uy);
      ctx.lineTo(lineX0 + w, uy);
      ctx.stroke();
    }
    if (format?.strike) {
      const sy = Math.round(ty - box.ascent * 0.34) + 0.5;
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
  // Honor explicit \n line breaks (Alt+Enter in a desktop spreadsheet).
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
    doubleInside: 1 | -1,
  ): void => {
    if (!side) return;
    const cfg = typeof side === 'object' ? side : { style: 'thin' as const };
    const color =
      (typeof side === 'object' && side.color) || theme.fgStrong || theme.fg || '#000000';
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
      // Spreadsheets keep double borders inside the owning cell. Drawing both
      // strokes inward avoids clipping on viewport and sheet edges.
      const gap = 3;
      const horizontal = y0 === y1;
      const dx = horizontal ? 0 : doubleInside * gap;
      const dy = horizontal ? doubleInside * gap : 0;
      ctx.beginPath();
      ctx.moveTo(x0, y0);
      ctx.lineTo(x1, y1);
      ctx.moveTo(x0 + dx, y0 + dy);
      ctx.lineTo(x1 + dx, y1 + dy);
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
  // Keep user borders inside the cell rect. This is especially important for
  // row/column 1 and viewport edges, where centered strokes can be clipped.
  drawSide(sides.top, xl, yt, xr, yt, 1);
  drawSide(sides.bottom, xl, yb, xr, yb, -1);
  drawSide(sides.left, xl, yt, xl, yb, 1);
  drawSide(sides.right, xr, yt, xr, yb, -1);
  drawSide(sides.diagonalDown, xl, yt, xr, yb, 1);
  drawSide(sides.diagonalUp, xl, yb, xr, yt, 1);
}

/** Active cell outline. Drawn in a separate pass after all cell text so the
 *  outline never gets clipped by neighbouring cell rects.
 *
 *  Desktop spreadsheets paint the active cell with a crisp ~2px green border and no
 *  inner fill or outer halo. */
export function paintActiveCellOutline(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): void {
  ctx.save();
  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 2;
  ctx.setLineDash([]);
  ctx.strokeRect(
    Math.round(bounds.x) + 1,
    Math.round(bounds.y) + 1,
    Math.max(0, Math.round(bounds.w) - 2),
    Math.max(0, Math.round(bounds.h) - 2),
  );
  ctx.restore();
}

/** Visible side length of the fill handle in CSS pixels. Desktop spreadsheets use a
 *  small accent-coloured square at the bottom-right of the active selection
 *  range; the user grabs it to drag-fill into adjacent cells. */
export const FILL_HANDLE_SIZE = 6;

/** Spreadsheet-style fill handle. Drawn at the selection range's bottom-right
 *  corner (or active-cell corner when selection is a single cell). The square
 *  is filled in `theme.accent` (resolves from `--fc-accent`, falls back to
 *  `#217346` when unset). */
export function paintFillHandle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): Rect {
  const hs = FILL_HANDLE_SIZE;
  // Centre the visible square on the cell's bottom-right corner so half the
  // handle bleeds outside the selection — matches the spreadsheet convention.
  const x = bounds.x + bounds.w - hs / 2;
  const y = bounds.y + bounds.h - hs / 2;
  const accent = theme.accent || '#0078d4';
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

/** Animated dashed copy-source marquee ("marching ants"). */
export function paintCopyMarquee(ctx: CanvasRenderingContext2D, bounds: Rect, phase = 0): void {
  ctx.save();
  ctx.lineWidth = 1;
  ctx.setLineDash([4, 3]);
  ctx.lineDashOffset = -phase;
  ctx.strokeStyle = '#ffffff';
  ctx.strokeRect(
    bounds.x + 1.5,
    bounds.y + 1.5,
    Math.max(0, bounds.w - 3),
    Math.max(0, bounds.h - 3),
  );
  ctx.lineDashOffset = 3.5 - phase;
  ctx.strokeStyle = '#111111';
  ctx.strokeRect(
    bounds.x + 0.5,
    bounds.y + 0.5,
    Math.max(0, bounds.w - 1),
    Math.max(0, bounds.h - 1),
  );
  ctx.restore();
}

/** Outline a dynamic-array spill range — spreadsheets paint a 1px accent ring
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

/** Highlight a cell that is blocking a #SPILL! result with a red dashed
 *  outline. Mirrors the "spill obstruction" indicator. */
export function paintSpillBlocker(ctx: CanvasRenderingContext2D, bounds: Rect): void {
  ctx.save();
  ctx.strokeStyle = '#d83b3b';
  ctx.lineWidth = 1.5;
  ctx.setLineDash([3, 3]);
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}

/** Hit rect for a painted checkbox glyph. Pointer code asks for this so a
 *  click on the box can flip the underlying value without dispatching the
 *  default "enter cell" action. */

export * from './painters/controls.js';
