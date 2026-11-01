import { formatNumber } from '../../commands/format.js';
import {
  phoneticAlignment,
  phoneticGuideSize,
  phoneticKanaForm,
} from '../../engine/phonetic-sync.js';
import { formatCell, formatGeneralScientific } from '../../engine/value.js';
import { intlFormatterCache } from '../../format/intl-cache.js';
import { convertKana } from '../../format/kana.js';
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

/** Width the spreadsheet charges a General number: every character, sign
 *  and separator included, costs one digit advance. */
const generalTextWidth = (ctx: CanvasRenderingContext2D, text: string): number =>
  [...text].length * ctx.measureText('0').width;

/** Excel-style "General" number rendering when the value does not fit. Drops
 *  fractional digits while a significant digit survives; a value that would
 *  round away to zero takes scientific notation if that fits, else the
 *  rounded zero; anything else falls back to scientific, then a `#` filler.
 *  General never groups thousands. */
function fitGeneralNumberToWidth(
  ctx: CanvasRenderingContext2D,
  value: number,
  availableWidth: number,
  locale: string,
): string {
  if (!Number.isFinite(value)) return String(value);
  const fits = (t: string): boolean => generalTextWidth(ctx, t) <= availableWidth;
  const scientific = (): string | null => {
    for (let prec = 5; prec >= 0; prec -= 1) {
      const t = formatGeneralScientific(value, prec, locale);
      if (fits(t)) return t;
    }
    return null;
  };
  const abs = Math.abs(value);
  // Already in scientific range (matches formatGeneralNumber's threshold) —
  // only the exponent precision can be reduced.
  if (abs > 0 && (abs >= 1e11 || abs < 1e-9)) {
    const sci = scientific();
    if (sci) return sci;
  } else {
    let roundedToZero: string | null = null;
    for (let digits = 12; digits >= 0; digits -= 1) {
      const t = intlFormatterCache
        .numberFormat(locale, { maximumFractionDigits: digits, useGrouping: false })
        .format(value);
      if (!fits(t)) continue;
      if (value === 0 || /[1-9]/.test(t)) return t;
      roundedToZero = t;
      break;
    }
    const sci = scientific();
    if (sci) return sci;
    if (roundedToZero !== null) return roundedToZero;
  }
  const digit = Math.max(1, ctx.measureText('0').width);
  return '#'.repeat(Math.max(1, Math.floor(availableWidth / digit)));
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

/** The face's own ascent and descent at the current `ctx.font`. The
 *  spreadsheet places a run by these, not by the ink of the glyphs in it;
 *  the stable ratios stand in where the canvas does not report them. */
export const fontMetricsBox = (ctx: CanvasRenderingContext2D, fontSize: number): TextMetricsBox => {
  const m = ctx.measureText('Hg');
  const ascent = m.fontBoundingBoxAscent;
  const descent = m.fontBoundingBoxDescent;
  return Number.isFinite(ascent) && ascent > 0 && Number.isFinite(descent) && descent >= 0
    ? { ascent, descent }
    : stableTextMetricsBox(fontSize);
};

/** Anchor x for text aligned in a cell; the indent applies on the side the
 *  text is aligned to. The right inset also clears half of the one-pixel
 *  gridline the cell owns on that edge, as measured against the spreadsheet. */
export const alignedTextAnchor = (
  bounds: Rect,
  align: CanvasTextAlign,
  padX: number,
  indentPx: number,
): number => {
  if (align === 'right') return bounds.x + bounds.w - 0.5 - padX - indentPx;
  if (align === 'center') return bounds.x + bounds.w / 2;
  return bounds.x + padX + indentPx;
};

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
  if (vAlign === 'top') return Math.round(bounds.y + padY + box.ascent);
  if (vAlign === 'middle') {
    return Math.round(bounds.y + (bounds.h - box.ascent - box.descent) / 2 + box.ascent);
  }
  return Math.round(bounds.y + bounds.h - padY - box.descent);
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

/** Set up the Outline / Shadow font effects on `ctx` (inside the caller's
 *  save) and return the glyph painter: outlined text is a hollow stroke in the
 *  text colour, shadowed text gets a hard offset shadow. */
const applyFontEffects = (
  ctx: CanvasRenderingContext2D,
  format: CellFormat | undefined,
  fontSize: number,
): ((text: string, x: number, y: number) => void) => {
  if (format?.fontShadow) {
    const offset = Math.max(1, fontSize / 14);
    ctx.shadowColor = 'rgba(0, 0, 0, 0.35)';
    ctx.shadowOffsetX = offset;
    ctx.shadowOffsetY = offset;
    ctx.shadowBlur = 0;
  }
  if (!format?.fontOutline) return (text, x, y) => ctx.fillText(text, x, y);
  const color = ctx.fillStyle;
  const width = Math.max(0.75, fontSize / 18);
  return (text, x, y) => {
    ctx.strokeStyle = color;
    ctx.lineWidth = width;
    ctx.strokeText(text, x, y);
  };
};

const paintTextDecorations = (
  ctx: CanvasRenderingContext2D,
  format: CellFormat | undefined,
  isHyperlink: boolean,
  lineX0: number,
  width: number,
  baselineY: number,
  fontSize: number,
): void => {
  const underline = format?.underline !== false && (!!format?.underline || isHyperlink);
  if (!underline && !format?.strike) return;
  ctx.strokeStyle = ctx.fillStyle as string;
  ctx.lineWidth = 1;
  // Placed from the run's own size, not the row-shared descent: the rule
  // starts 0.1 em under the baseline and strikeout sits 0.26 em above it.
  if (underline) {
    const underlineY = Math.round(baselineY + Math.max(1, fontSize * 0.1)) + 0.5;
    const ys = isDoubleUnderline(format?.underline) ? [underlineY, underlineY + 2] : [underlineY];
    for (const y of ys) {
      ctx.beginPath();
      ctx.moveTo(lineX0, y);
      ctx.lineTo(lineX0 + width, y);
      ctx.stroke();
    }
  }
  if (format?.strike) {
    const strikeY = Math.round(baselineY - fontSize * 0.26) + 0.5;
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
  date1904 = false,
  rtl = false,
  zoom,
  formatPaint: formatPaintFor,
  fontEffects = false,
  rowBaselineFont,
}: CellPaintCtx): void {
  const hyperlinkDisplay = format?.hyperlinkDisplay ?? '';
  if (value.kind === 'blank' && !formula && displayOverride == null && !hyperlinkDisplay) return;
  if (showZeros === false && value.kind === 'number' && value.value === 0 && !showFormulas) return;

  // Text keeps clear of the gridline the cell owns on its bottom edge and of
  // the one before it; the face's own ascent and descent do the rest.
  const padY = 1;
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
        ? formatNumber(value.value, format.numFmt, locale, date1904)
        : formatCell(value, locale);
  }
  if (!text) return;

  const isNumeric = value.kind === 'number';
  const isError = value.kind === 'error';
  const isBool = value.kind === 'bool';
  const isFormulaDisplay = showFormulas && formula != null;
  const formatPaint = isFormulaDisplay ? null : formatPaintFor;
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
  // Bottom-aligned runs in a row share the baseline the deepest descent sets.
  const withRowDescent = (box: TextMetricsBox, vAlign: TextVAlign): TextMetricsBox => {
    if (vAlign !== 'bottom' || !rowBaselineFont) return box;
    const rowFont = cellFont(rowBaselineFont, theme, false);
    ctx.font = cellFontCss({ ...rowFont, size: rowFont.size * zoomScale });
    const rowBox = fontMetricsBox(ctx, rowFont.size * zoomScale);
    ctx.font = paintFontCss;
    return { ascent: box.ascent, descent: Math.max(box.descent, rowBox.descent) };
  };
  ctx.font = paintFontCss;
  const isHyperlink = !!format?.hyperlink;
  // A number-format colour tag outranks the cell's font colour.
  ctx.fillStyle = formatPaint?.color
    ? formatPaint.color
    : format?.color
      ? format.color
      : isHyperlink
        ? theme.accent
        : isError
          ? theme.cellErrorFg
          : isBool
            ? theme.cellBoolFg
            : isNumeric
              ? theme.cellNumFg
              : theme.cellAutoFg;

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
  const drawText = applyFontEffects(ctx, fontEffects ? format : undefined, fontSize);

  // Rotated text — render around cell center, ignore wrap/indent for
  // simplicity. (Real desktop spreadsheets are more elaborate but this covers ±90° common case.)
  if (rotation !== 0) {
    ctx.translate(bounds.x + bounds.w / 2, bounds.y + bounds.h / 2);
    ctx.rotate((rotation * Math.PI) / 180);
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    drawText(text, 0, 0);
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
    const measured = withRowDescent(fontMetricsBox(ctx, fontSize), vAlign);
    let startY = textBaselineY(bounds, measured, vAlign, padY);
    startY += cellTextVerticalOffset(fontSize, fontVertAlign);
    if (vAlign === 'middle') startY -= (totalH - lineH) / 2;
    else if (vAlign === 'bottom') startY -= (lines.length - 1) * lineH;
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = align;
    const anchor = alignedTextAnchor(bounds, align, padX, indentPx);
    for (let i = 0; i < lines.length; i += 1) {
      const line = lines[i] ?? '';
      const baselineY = startY + i * lineH;
      const width = ctx.measureText(line).width;
      const tx = anchor;
      drawText(line, tx, baselineY);
      const lineX0 = align === 'right' ? tx - width : align === 'center' ? tx - width / 2 : tx;
      paintTextDecorations(ctx, format, isHyperlink, lineX0, width, baselineY, fontSize);
    }
    ctx.restore();
    return;
  }

  // Single-line path with vertical alignment.
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = align;
  const tx = alignedTextAnchor(bounds, align, padX, indentPx);

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
  const generalNumber =
    !isFormulaDisplay &&
    value.kind === 'number' &&
    (!format?.numFmt || format.numFmt.kind === 'general');
  const overflows = generalNumber
    ? generalTextWidth(ctx, text) > availableTextWidthForCell
    : ctx.measureText(text).width > availableTextWidthForCell;
  if (isNumeric && overflows) {
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
      );
    }
  }
  const metrics = ctx.measureText(text);
  const box = withRowDescent(fontMetricsBox(ctx, drawFontSize), vAlign);
  let ty = textBaselineY(bounds, box, vAlign, padY);
  ty += cellTextVerticalOffset(drawFontSize, fontVertAlign);

  // A `*x` fill spans the cell whatever the alignment: one part sits at
  // each edge with the fill character repeated between them.
  const fill = formatPaint?.fill;
  const fillParts =
    !fill || rtl
      ? null
      : fill.lead !== undefined && text.startsWith(fill.lead)
        ? [fill.lead, text.slice(fill.lead.length)]
        : fill.trail !== undefined && text.endsWith(fill.trail)
          ? [text.slice(0, text.length - fill.trail.length), fill.trail]
          : null;
  if (fill && fillParts) {
    const [leading = '', trailing = ''] = fillParts;
    const left = bounds.x + padX + indentPx;
    const right = bounds.x + bounds.w - padX;
    const leadingWidth = ctx.measureText(leading).width;
    const trailingWidth = ctx.measureText(trailing).width;
    ctx.textAlign = 'left';
    drawText(leading, left, ty);
    ctx.textAlign = 'right';
    drawText(trailing, right, ty);
    const charWidth = ctx.measureText(fill.char).width;
    const gap = right - trailingWidth - (left + leadingWidth);
    if (fill.char.trim() && charWidth > 0 && gap >= charWidth) {
      ctx.textAlign = 'left';
      drawText(fill.char.repeat(Math.floor(gap / charWidth)), left + leadingWidth, ty);
    }
    paintTextDecorations(ctx, format, isHyperlink, left, right - left, ty, drawFontSize);
    ctx.restore();
    return;
  }

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
  drawText(text, tx, ty);

  const lineX0 =
    align === 'right' ? tx - metrics.width : align === 'center' ? tx - metrics.width / 2 : tx;
  paintTextDecorations(ctx, format, isHyperlink, lineX0, metrics.width, ty, drawFontSize);
  ctx.restore();
}
