import { formatNumber } from '../../commands/format.js';
import { formatCell } from '../../engine/value.js';
import type { CellFormat } from '../../store/store.js';
import type { Rect } from '../geometry.js';
import type { CellPaintCtx, TextMetricsBox, TextVAlign } from './types.js';

function canvasTextAlign(align: CellFormat['align'] | undefined): CanvasTextAlign | null {
  switch (align) {
    case 'left':
    case 'center':
    case 'right':
      return align;
    case 'centerContinuous':
      return 'center';
    case 'fill':
    case 'justify':
    case 'distributed':
      return 'left';
    default:
      return null;
  }
}

function canvasTextVAlign(align: CellFormat['vAlign'] | undefined): TextVAlign {
  switch (align) {
    case 'top':
    case 'middle':
    case 'bottom':
      return align;
    default:
      return 'middle';
  }
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

  const padX = 3;
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
  const redNegative =
    value.kind === 'number' &&
    value.value < 0 &&
    (format?.numFmt?.kind === 'fixed' || format?.numFmt?.kind === 'currency') &&
    (format.numFmt.negativeStyle === 'red' || format.numFmt.negativeStyle === 'red-parens');
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
          : redNegative
            ? '#c00000'
            : isNumeric
              ? theme.cellNumFg
              : theme.fg;

  let align: CanvasTextAlign;
  if (format?.align) {
    align = canvasTextAlign(format.align) ?? 'left';
  } else {
    align = isFormulaDisplay ? 'left' : isNumeric ? 'right' : isBool || isError ? 'center' : 'left';
  }
  const indentPx = (format?.indent ?? 0) * 8;
  const rotation = format?.rotation ?? 0;
  const wrap = !!format?.wrap;
  ctx.direction =
    format?.textDirection === 'rtl' ? 'rtl' : format?.textDirection === 'ltr' ? 'ltr' : 'inherit';

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
    const vAlign = canvasTextVAlign(format?.vAlign ?? 'bottom');
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

  const vAlign = canvasTextVAlign(format?.vAlign ?? 'bottom');
  const availableTextWidth = Math.max(0, bounds.w - padX * 2 - indentPx);
  let drawFontSize = fontSize;
  if (format?.shrinkToFit) {
    const measuredWidth = ctx.measureText(text).width;
    if (measuredWidth > availableTextWidth && availableTextWidth > 0) {
      drawFontSize = Math.max(8, Math.floor(fontSize * (availableTextWidth / measuredWidth)));
      ctx.font = `${styleSlant}${weight} ${drawFontSize}px ${fontCss(fontFamily)}`;
    }
  }
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
        drawFontSize,
      );
    }
  }
  const metrics = ctx.measureText(text);
  const box = stableTextMetricsBox(drawFontSize);
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
