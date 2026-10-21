import { formatNumber } from '../../commands/format.js';
import {
  phoneticAlignment,
  phoneticGuideSize,
  phoneticKanaForm,
} from '../../engine/phonetic-sync.js';
import { formatCell, formatScientificNumber } from '../../engine/value.js';
import { convertKana } from '../../format/kana.js';
import { normalizeFormatLocale } from '../../format/locale.js';
import type { CellFormat } from '../../store/store.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import type { Rect } from '../geometry.js';
import { cellTextHorizontalMetrics, wrapText } from '../text-wrap.js';
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
  const toSci = (prec: number): string => formatScientificNumber(value, prec, locale);
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
      const t = new Intl.NumberFormat(normalizeFormatLocale(locale), {
        maximumFractionDigits: digits,
      }).format(value);
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

/** Normalize the view multiplier at the painter/DOM boundary. A malformed or
 * missing view value must not make a cell disappear by producing a zero-sized
 * font. */
export const visualZoom = (zoom: number | undefined): number =>
  zoom !== undefined && Number.isFinite(zoom) && zoom > 0 ? zoom : 1;

/** The size reduction used by Excel for whole-cell superscript/subscript. */
export const fontSizeWithVerticalAlign = (
  size: number,
  verticalAlign: CellFormat['fontVertAlign'] | undefined,
): number => (verticalAlign ? size * 0.7 : size);

/** Baseline displacement matching the stable metrics used by the canvas
 * painter. Positive values move the run down (subscript). */
export const cellTextVerticalOffset = (
  fontSize: number,
  verticalAlign: CellFormat['fontVertAlign'] | undefined,
): number => {
  if (verticalAlign === 'superscript') return -fontSize * 0.72 * 0.35;
  if (verticalAlign === 'subscript') return fontSize * 0.22 * 0.65;
  return 0;
};

/** Return a CSS font-size suitable for a DOM cell surface. The variable
 * fallback keeps an inherited workbook/default font zoomable when the native
 * effective style has not yet been projected into the store. */
export const cssCellFontSize = (
  size: number | undefined,
  verticalAlign: CellFormat['fontVertAlign'] | undefined,
  zoom = 1,
  fallback = 'var(--fc-cell-text-size)',
): string => {
  const factor = visualZoom(zoom) * (verticalAlign ? 0.7 : 1);
  if (size !== undefined && Number.isFinite(size)) return `${size * factor}px`;
  return factor === 1 ? '' : `calc(${fallback} * ${factor})`;
};

/** Shared shrink-to-fit calculation for DOM previews and canvas painting. */
export const shrinkFontSizeToFit = (
  fontSize: number,
  measuredWidth: number,
  availableWidth: number,
  minimum = 8,
): number => {
  if (
    !Number.isFinite(fontSize) ||
    !Number.isFinite(measuredWidth) ||
    !Number.isFinite(availableWidth) ||
    measuredWidth <= availableWidth ||
    availableWidth <= 0
  ) {
    return fontSize;
  }
  return Math.max(
    Math.min(fontSize, minimum),
    Math.floor(fontSize * (availableWidth / measuredWidth)),
  );
};

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

const splitFontFamily = (family: string): string[] => {
  const parts: string[] = [];
  let start = 0;
  let depth = 0;
  let quote = '';
  for (let i = 0; i < family.length; i += 1) {
    const char = family[i];
    if (quote) {
      if (char === quote && family[i - 1] !== '\\') quote = '';
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }
    if (char === '(') depth += 1;
    else if (char === ')' && depth > 0) depth -= 1;
    else if (char === ',' && depth === 0) {
      parts.push(family.slice(start, i).trim());
      start = i + 1;
    }
  }
  parts.push(family.slice(start).trim());
  return parts.filter(Boolean);
};

const normalizedFontFamily = (family: string): string => {
  const trimmed = family.trim();
  if (
    trimmed.length >= 2 &&
    ((trimmed.startsWith('"') && trimmed.endsWith('"')) ||
      (trimmed.startsWith("'") && trimmed.endsWith("'")))
  ) {
    return trimmed.slice(1, -1).trim().toLocaleLowerCase();
  }
  return trimmed.toLocaleLowerCase();
};

/**
 * Return the raw CSS font-family stack used for cell text. An authored cell
 * family stays first so an installed workbook font wins; the UI stack then
 * supplies a deterministic fallback when that named font is unavailable.
 * Missing cell families deliberately return the supplied theme stack unchanged
 * so formula-display mode can continue using its monospace theme.
 */
export const cellFontFamily = (explicitFamily: string | undefined, themeUi: string): string => {
  const explicit = explicitFamily?.trim() ?? '';
  if (!explicit) return themeUi.trim();
  const fallback = themeUi.trim();
  if (!fallback) return explicit;
  const seen = new Set(splitFontFamily(explicit).map(normalizedFontFamily));
  const additions = splitFontFamily(fallback).filter((family) => {
    const normalized = normalizedFontFamily(family);
    if (seen.has(normalized)) return false;
    seen.add(normalized);
    return true;
  });
  return additions.length > 0 ? `${explicit}, ${additions.join(', ')}` : explicit;
};

const fontCss = (family: string): string =>
  splitFontFamily(family)
    .map((part) => {
      const name = part.trim();
      if (!name) return name;
      if (name.startsWith('"') || name.startsWith("'")) return name;
      if (/^[a-zA-Z-]+$/.test(name)) return name;
      return `"${name.replaceAll('"', '\\"')}"`;
    })
    .filter(Boolean)
    .join(', ');

/** Theme fields that decide a cell's default font. */
export type CellFontTheme = Pick<ResolvedTheme, 'textCell' | 'fontUi' | 'fontMono'>;

export interface CellFont {
  slant: string;
  weight: number;
  size: number;
  family: string;
}

/** Font a cell's text paints in; autofit measures with the same font. */
export const cellFont = (
  format:
    | Pick<CellFormat, 'bold' | 'italic' | 'fontSize' | 'fontFamily' | 'fontVertAlign'>
    | undefined,
  theme: CellFontTheme,
  formulaDisplay: boolean,
): CellFont => ({
  slant: format?.italic ? 'italic ' : '',
  weight: format?.bold ? 700 : 400,
  size: fontSizeWithVerticalAlign(format?.fontSize ?? theme.textCell, format?.fontVertAlign),
  family: format?.fontFamily
    ? cellFontFamily(format.fontFamily, theme.fontUi)
    : formulaDisplay
      ? theme.fontMono
      : theme.fontUi,
});

export const cellFontCss = (font: CellFont, size = font.size): string =>
  `${font.slant}${font.weight} ${size}px ${fontCss(font.family)}`;

const isDoubleUnderline = (underline: CellFormat['underline']): boolean =>
  underline === 'double' || underline === 'doubleAccounting';

const paintTextDecorations = (
  ctx: CanvasRenderingContext2D,
  format: CellFormat | undefined,
  isHyperlink: boolean,
  lineX0: number,
  width: number,
  baselineY: number,
  box: TextMetricsBox,
): void => {
  if (!format?.underline && !format?.strike && !isHyperlink) return;
  ctx.strokeStyle = ctx.fillStyle as string;
  ctx.lineWidth = 1;
  if (format?.underline || isHyperlink) {
    const underlineY = Math.round(baselineY + Math.max(2, box.descent * 0.55)) + 0.5;
    const ys = isDoubleUnderline(format?.underline) ? [underlineY, underlineY + 2] : [underlineY];
    for (const y of ys) {
      ctx.beginPath();
      ctx.moveTo(lineX0, y);
      ctx.lineTo(lineX0 + width, y);
      ctx.stroke();
    }
  }
  if (format?.strike) {
    const strikeY = Math.round(baselineY - box.ascent * 0.34) + 0.5;
    ctx.beginPath();
    ctx.moveTo(lineX0, strikeY);
    ctx.lineTo(lineX0 + width, strikeY);
    ctx.stroke();
  }
};

export function paintCellText({
  ctx,
  bounds,
  theme,
  value,
  formula,
  format,
  showFormulas,
  showZeros,
  displayOverride,
  locale,
  rtl = false,
  zoom,
}: CellPaintCtx): void {
  const hyperlinkDisplay = format?.hyperlinkDisplay ?? '';
  if (value.kind === 'blank' && !formula && displayOverride == null && !hyperlinkDisplay) return;
  if (showZeros === false && value.kind === 'number' && value.value === 0 && !showFormulas) return;

  const padY = 3;
  let text: string;
  if (displayOverride != null) {
    text = displayOverride;
  } else if (showFormulas && formula) {
    text = formula;
  } else if (value.kind === 'blank' && hyperlinkDisplay) {
    text = hyperlinkDisplay;
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
  const fontVertAlign = format?.fontVertAlign;
  const baseFont = cellFont(format, theme, isFormulaDisplay === true);
  const zoomScale = visualZoom(zoom);
  const font = {
    ...baseFont,
    size: baseFont.size * zoomScale,
  };
  const horizontal = cellTextHorizontalMetrics(bounds.w, format?.indent ?? 0, zoomScale);
  const padX = horizontal.padX;
  const indentPx = horizontal.indentPx;
  const fontSize = font.size;
  const paintFontCss = cellFontCss(font);
  ctx.font = paintFontCss;
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
    // An explicit alignment stays physical — the ribbon's left/right buttons
    // mean the same thing whichever way the sheet runs.
    align = canvasTextAlign(format.align) ?? 'left';
  } else {
    // General alignment follows reading order, so it flips with the sheet.
    const general: CanvasTextAlign = isFormulaDisplay
      ? 'left'
      : isNumeric
        ? 'right'
        : isBool || isError
          ? 'center'
          : 'left';
    align = rtl && general === 'left' ? 'right' : rtl && general === 'right' ? 'left' : general;
  }
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
    const logicalHorizontal = cellTextHorizontalMetrics(
      bounds.w / zoomScale,
      format?.indent ?? 0,
      1,
    );
    ctx.font = cellFontCss(baseFont);
    const lines = wrapText(text, {
      maxWidth: logicalHorizontal.availableWidth,
      measureText: (value) => ctx.measureText(value).width,
      locale,
    });
    ctx.font = paintFontCss;
    const lineH = Math.round(baseFont.size * 1.28) * zoomScale;
    const totalH = lineH * lines.length;
    const vAlign = canvasTextVAlign(format?.vAlign ?? 'bottom');
    const measured = stableTextMetricsBox(fontSize);
    let startY = textBaselineY(bounds, measured, vAlign, padY);
    startY += cellTextVerticalOffset(fontSize, fontVertAlign);
    if (vAlign === 'middle') startY -= (totalH - lineH) / 2;
    else if (vAlign === 'bottom') startY -= (lines.length - 1) * lineH;
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = align;
    let tx: number;
    if (align === 'right') tx = bounds.x + bounds.w - padX - (rtl ? indentPx : 0);
    else if (align === 'center') tx = bounds.x + bounds.w / 2;
    else tx = bounds.x + padX + indentPx;
    for (let i = 0; i < lines.length; i += 1) {
      const line = lines[i] ?? '';
      const baselineY = startY + i * lineH;
      ctx.fillText(line, tx, baselineY);
      const width = ctx.measureText(line).width;
      const lineX0 = align === 'right' ? tx - width : align === 'center' ? tx - width / 2 : tx;
      paintTextDecorations(ctx, format, isHyperlink, lineX0, width, baselineY, measured);
    }
    ctx.restore();
    return;
  }

  // Single-line path with vertical alignment.
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = align;
  let tx: number;
  if (align === 'right') tx = bounds.x + bounds.w - padX - (rtl ? indentPx : 0);
  else if (align === 'center') tx = bounds.x + bounds.w / 2;
  else tx = bounds.x + padX + indentPx;

  const vAlign = canvasTextVAlign(format?.vAlign ?? 'bottom');
  const availableTextWidthForCell = horizontal.availableWidth;
  let drawFontSize = fontSize;
  if (format?.shrinkToFit) {
    const measuredWidth = ctx.measureText(text).width;
    if (measuredWidth > availableTextWidthForCell && availableTextWidthForCell > 0) {
      drawFontSize = shrinkFontSizeToFit(
        fontSize,
        measuredWidth,
        availableTextWidthForCell,
        8 * zoomScale,
      );
      ctx.font = cellFontCss(font, drawFontSize);
    }
  }
  if (isNumeric && ctx.measureText(text).width > availableTextWidthForCell) {
    if (format?.numFmt && format.numFmt.kind !== 'general') {
      // Explicit number format: Excel shows #### when the formatted value
      // can't shrink without losing the user's chosen presentation.
      const hashWidth = Math.max(1, ctx.measureText('#').width || fontSize * 0.62);
      text = '#'.repeat(Math.max(1, Math.floor(availableTextWidthForCell / hashWidth)));
    } else if (!isFormulaDisplay && value.kind === 'number') {
      // General number: Excel progressively trims fractional digits before
      // falling back to scientific notation and finally ####. Without this,
      // right-aligned text just clips and the user sees only the trailing
      // digits — making the result look corrupted.
      text = fitGeneralNumberToWidth(
        ctx,
        value.value,
        availableTextWidthForCell,
        locale ?? 'en-US',
        drawFontSize,
      );
    }
  }
  const metrics = ctx.measureText(text);
  const box = stableTextMetricsBox(drawFontSize);
  let ty = textBaselineY(bounds, box, vAlign, padY);
  ty += cellTextVerticalOffset(drawFontSize, fontVertAlign);

  // A phonetic guide (Japanese furigana/ruby) is stored separately from the
  // cell value by OOXML, one run per annotated span. Each reading sits over
  // the span it annotates, measured in the base font before the ruby font is
  // selected; a run whose span does not land inside the drawn text — the
  // whole-cell reading an engine without a per-run surface reports, or a guide
  // left over from a longer value — sits over the whole text instead.
  // `phoneticProperties` picks the kana form and the alignment over the span.
  // The guide is drawn only while Show Phonetic is on for the cell.
  if (format?.showPhonetic && format.phonetic?.length && !isFormulaDisplay) {
    const textLeft =
      align === 'right' ? tx - metrics.width : align === 'center' ? tx - metrics.width / 2 : tx;
    const spans = format.phonetic.map((run) => {
      const start = Math.max(0, Math.min(run.start, text.length));
      const end = Math.max(start, Math.min(run.end, text.length));
      if (end === start) return { left: textLeft, width: metrics.width };
      const before = start === 0 ? 0 : ctx.measureText(text.slice(0, start)).width;
      return { left: textLeft + before, width: ctx.measureText(text.slice(start, end)).width };
    });
    const kana = phoneticKanaForm(format.phoneticProperties);
    const alignment = phoneticAlignment(format.phoneticProperties);
    const phoneticSize = phoneticGuideSize(drawFontSize);
    ctx.save();
    ctx.font = cellFontCss(font, phoneticSize);
    ctx.textBaseline = 'alphabetic';
    const rubyY = ty - box.ascent - 1;
    for (const [i, run] of format.phonetic.entries()) {
      const span = spans[i] ?? { left: textLeft, width: metrics.width };
      const reading = convertKana(run.text, kana);
      const chars = [...reading];
      if (alignment === 'left') {
        ctx.textAlign = 'left';
        ctx.fillText(reading, span.left, rubyY);
      } else if (
        alignment === 'distributed' &&
        chars.length > 1 &&
        ctx.measureText(reading).width < span.width
      ) {
        // Each character takes an equal slot of the span.
        ctx.textAlign = 'center';
        const slot = span.width / chars.length;
        for (const [k, ch] of chars.entries()) {
          ctx.fillText(ch, span.left + slot * (k + 0.5), rubyY);
        }
      } else {
        ctx.textAlign = 'center';
        ctx.fillText(reading, span.left + span.width / 2, rubyY);
      }
    }
    ctx.restore();
  }
  ctx.fillText(text, tx, ty);

  const lineX0 =
    align === 'right' ? tx - metrics.width : align === 'center' ? tx - metrics.width / 2 : tx;
  paintTextDecorations(ctx, format, isHyperlink, lineX0, metrics.width, ty, box);
  ctx.restore();
}
