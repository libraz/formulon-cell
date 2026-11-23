import { formatNumber } from '../../commands/format.js';
import { fontPointsToCssPx } from '../../engine/font-units.js';
import {
  phoneticAlignment,
  phoneticGuideSize,
  phoneticKanaForm,
} from '../../engine/phonetic-sync.js';
import { formatCell, formatGeneralScientific } from '../../engine/value.js';
import { splitFontFamily } from '../../format/font-family.js';
import { intlFormatterCache } from '../../format/intl-cache.js';
import { convertKana } from '../../format/kana.js';
import type { CellFormat } from '../../store/store.js';
import { type ResolvedTheme, WORKBOOK_TEXT_COLOR } from '../../theme/resolve.js';
import type { Rect } from '../geometry.js';
import {
  cellLinePitch,
  cellTextWidth,
  drawCellTextRun,
  paintedFontSize,
  wholePointPadX,
} from '../text-metrics.js';
import { cellTextHorizontalMetrics, graphemes, wrapText } from '../text-wrap.js';
import { hasCellFill } from './cell.js';
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
const generalTextWidth = (measure: (text: string) => number, text: string): number =>
  [...text].length * measure('0');

/** Excel-style "General" number rendering when the value does not fit. Drops
 *  fractional digits while a significant digit survives; a value that would
 *  round away to zero takes scientific notation if that fits, else the
 *  rounded zero; anything else falls back to scientific, then a `#` filler.
 *  General never groups thousands. */
function fitGeneralNumberToWidth(
  measure: (text: string) => number,
  value: number,
  availableWidth: number,
  locale: string,
): string {
  if (!Number.isFinite(value)) return String(value);
  const fits = (t: string): boolean => generalTextWidth(measure, t) <= availableWidth;
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
  const digit = Math.max(1, measure('0'));
  return '#'.repeat(Math.max(1, Math.floor(availableWidth / digit)));
}

export const stableTextMetricsBox = (fontSize: number): TextMetricsBox => ({
  ascent: fontSize * 0.72,
  descent: fontSize * 0.22,
});

/** The face's own ascent and descent at the current `ctx.font`. The
 *  spreadsheet places a run by these, not by the ink of the glyphs in it;
 *  the stable ratios stand in where the canvas does not report them. The
 *  canvas reports them from the current `textBaseline`, so measure from the
 *  alphabetic one whatever an earlier painter left set. */
export const fontMetricsBox = (ctx: CanvasRenderingContext2D, fontSize: number): TextMetricsBox => {
  const baseline = ctx.textBaseline;
  ctx.textBaseline = 'alphabetic';
  const m = ctx.measureText('Hg');
  ctx.textBaseline = baseline;
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

/** Whole-cell superscript/subscript size, as a fraction of the base size. */
const SCRIPT_SIZE = 0.7;
/** Script baseline shifts in em of the base size: Arial's OS/2 offsets, which
 *  match Excel's 10 px rise and 3-4 px drop at 11 pt, DPR 2. */
const SUPERSCRIPT_RISE_EM = 0.477;
const SUBSCRIPT_DROP_EM = 0.138;

/** The size reduction used by Excel for whole-cell superscript/subscript. */
export const fontSizeWithVerticalAlign = (
  size: number,
  verticalAlign: CellFormat['fontVertAlign'] | undefined,
): number => (verticalAlign ? size * SCRIPT_SIZE : size);

/** Baseline displacement for a run drawn at the reduced script `fontSize`.
 *  Positive values move the run down (subscript). */
export const cellTextVerticalOffset = (
  fontSize: number,
  verticalAlign: CellFormat['fontVertAlign'] | undefined,
): number => {
  const base = fontSize / SCRIPT_SIZE;
  if (verticalAlign === 'superscript') return -base * SUPERSCRIPT_RISE_EM;
  if (verticalAlign === 'subscript') return base * SUBSCRIPT_DROP_EM;
  return 0;
};

/** An accounting underline in a row lifts the row's shared bottom baseline by
 *  this many em of the row font: 4 device px at Arial 11, DPR 2. */
const ACCOUNTING_BASELINE_RISE_EM = 2 / 11;

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

/** Shrink-to-fit font size. The spreadsheet has no minimum size: a long run
 *  shrinks as far as the cell needs, down to a one-pixel floor. */
export const shrinkFontSizeToFit = (
  fontSize: number,
  measuredWidth: number,
  availableWidth: number,
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
  return Math.max(1, Math.floor(fontSize * (availableWidth / measuredWidth)));
};

/** Bottom gap, in points, from the baseline to the cell's bottom edge for
 *  fonts whose placement was measured: constant across sizes and row heights.
 *  Calibri 11 pt in a 24 pt row lands 4 pt above the paint bounds at 50-200%. */
const BOTTOM_GAP_PT: Record<string, number> = {
  calibri: 4,
  'aptos narrow': 4,
};

/** Bottom-aligned baseline rise above the cell's bottom edge, `slope * size +
 *  intercept` CSS px, for fonts measured against the size ladder at 11 pt.
 *  The face's own descent cannot stand in: these fonts are often absent from
 *  the host, and no single em ratio fits their rise across zoom levels. */
const BOTTOM_BASELINE_RISE: Record<string, { slope: number; intercept: number }> = {
  'ms pgothic': { slope: 0.45, intercept: -0.6 },
  'ms pゴシック': { slope: 0.45, intercept: -0.6 },
};

/** Painted font plus the render scale (CSS px per point) it was sized with. */
type BaselineFace = { family: string; size: number; scale?: number };

const bottomBaselineRise = (face: BaselineFace): number | null => {
  const first = splitFontFamily(face.family)[0]
    ?.replace(/^["']|["']$/g, '')
    .normalize('NFKC')
    .toLowerCase();
  const gapPt = first ? BOTTOM_GAP_PT[first] : undefined;
  if (gapPt !== undefined) return fontPointsToCssPx(gapPt) * (face.scale ?? 1);
  const fit = first ? BOTTOM_BASELINE_RISE[first] : undefined;
  return fit ? fit.slope * face.size + fit.intercept : null;
};

/** `merged`: the cell is a merged range, whose centred line box rounds a
 *  half pixel down the sheet where a single cell floors it. `face` names the
 *  painted font so a bottom-aligned run can use its measured baseline. */
export const textBaselineY = (
  bounds: Rect,
  box: TextMetricsBox,
  vAlign: TextVAlign,
  padY: number,
  merged = false,
  face?: BaselineFace,
): number => {
  if (vAlign === 'top') return Math.round(bounds.y + padY + box.ascent);
  if (vAlign === 'middle') {
    const centredTop = bounds.y + (bounds.h - box.ascent - box.descent) / 2;
    const top = merged ? centredTop : Math.floor(centredTop);
    return Math.round(top + box.ascent);
  }
  const rise = face ? bottomBaselineRise(face) : null;
  if (rise !== null) return Math.round(bounds.y + bounds.h - rise);
  return Math.round(bounds.y + bounds.h - padY - box.descent);
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

/** Rule rows relative to the single underline, as measured at 11 px: double
 *  brackets it, the accounting styles sit one pixel lower. */
const underlineRows = (underline: CellFormat['underline']): number[] => {
  switch (underline) {
    case 'double':
      return [-2, 0];
    case 'singleAccounting':
      return [1];
    case 'doubleAccounting':
      return [1, 3];
    default:
      return [0];
  }
};

/** Horizontal extent of a text cell's padded area; an accounting underline
 *  on text spans it. */
interface DecorationSpan {
  left: number;
  right: number;
}

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
  cellSpan?: DecorationSpan,
): void => {
  const underline = format?.underline !== false && (!!format?.underline || isHyperlink);
  if (!underline && !format?.strike) return;
  ctx.strokeStyle = ctx.fillStyle as string;
  ctx.lineWidth = 1;
  // Placed from the run's own size, not the row-shared descent: the rule
  // starts 0.1 em under the baseline and strikeout sits 0.26 em above it.
  if (underline) {
    const underlineY = Math.round(baselineY + Math.max(1, fontSize * 0.1)) + 0.5;
    const accounting =
      format?.underline === 'singleAccounting' || format?.underline === 'doubleAccounting';
    // Accounting rules cover the padded cell for text and the run for
    // numbers, on whole pixels.
    const x0 = accounting ? Math.round(cellSpan?.left ?? lineX0) : lineX0;
    const x1 = accounting ? Math.round(cellSpan?.right ?? lineX0 + width) : lineX0 + width;
    for (const row of underlineRows(format?.underline)) {
      ctx.beginPath();
      ctx.moveTo(x0, underlineY + row);
      ctx.lineTo(x1, underlineY + row);
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

/** Wrap at 100%: the canvas run widths of the unzoomed font against the
 *  cell's stored width, so a zoom never moves a line break. */
const wrapLogical = (
  ctx: CanvasRenderingContext2D,
  text: string,
  baseFont: CellFont,
  logicalWidth: number,
  indent: number,
  locale: string | undefined,
): string[] => {
  ctx.font = cellFontCss(baseFont);
  return wrapText(text, {
    maxWidth: cellTextHorizontalMetrics(logicalWidth, indent, 1).availableWidth,
    measureText: (value) => ctx.measureText(value).width,
    locale,
  });
};

/** Stacked text: one character per line, centred on its column's axis. Wrap
 *  breaks words into columns no taller than the cell, laid right to left;
 *  each column takes the vertical alignment on its own. The indent moves the
 *  run away from the vertical edge it is aligned to. */
function paintStackedText(
  ctx: CanvasRenderingContext2D,
  drawText: (text: string, x: number, y: number) => void,
  text: string,
  bounds: Rect,
  opts: {
    format: CellFormat | undefined;
    wrap: boolean;
    locale: string | undefined;
    fontSize: number;
    padX: number;
    padY: number;
    indentPx: number;
  },
): void {
  const { format, fontSize, padX, padY, indentPx } = opts;
  const lineH = cellLinePitch(fontSize);
  const maxLines = Math.max(1, Math.floor((bounds.h - padY * 2) / lineH));
  const columns = (
    opts.wrap
      ? wrapText(text, {
          maxWidth: maxLines,
          measureText: (run) => graphemes(run, opts.locale).length,
          locale: opts.locale,
        })
      : [text.replace(/\r?\n/g, '')]
  ).map((run) => graphemes(run, opts.locale));
  // Columns are an em wide on axes one line pitch apart. Measured against the
  // desktop, the left side keeps half a padding more clear and the right side
  // one more padding, as rotated runs do.
  const blockW = fontSize + (columns.length - 1) * lineH;
  const left = bounds.x + padX * 1.5;
  const right = alignedTextAnchor(bounds, 'right', padX, 0) - padX;
  const blockRight =
    format?.align === 'left'
      ? left + blockW
      : format?.align === 'right'
        ? right
        : (left + right) / 2 + blockW / 2;
  const vAlign = canvasTextVAlign(format?.vAlign ?? 'bottom');
  const { ascent } = fontMetricsBox(ctx, fontSize);
  const vertOffset = cellTextVerticalOffset(fontSize, format?.fontVertAlign);
  ctx.textAlign = 'center';
  ctx.textBaseline = 'alphabetic';
  columns.forEach((chars, column) => {
    const axis = blockRight - fontSize / 2 - column * lineH;
    // Each character owns a line box one pitch tall with its baseline an
    // ascent below the box top; the column's boxes take the alignment.
    const blockH = chars.length * lineH;
    const top =
      vAlign === 'top'
        ? bounds.y + padY + indentPx
        : vAlign === 'middle'
          ? bounds.y + (bounds.h - blockH) / 2
          : bounds.y + bounds.h - padY - blockH - indentPx;
    const firstY = Math.round(top + ascent) + vertOffset;
    chars.forEach((char, line) => {
      drawText(char, axis, firstY + line * lineH);
    });
  });
}

/** Text 1 of the default workbook theme. */
const DEFAULT_WORKBOOK_TEXT = '#000000';

/** Colour of automatic cell text. On a painted fill it is the theme's
 *  fill-text colour, which may defer to the workbook's Text 1: automatic
 *  means Text 1 in the file format, and an authored fill is drawn for it. */
export function automaticTextColor(
  theme: Pick<ResolvedTheme, 'cellAutoFg' | 'cellNumFg' | 'cellFillAutoFg'>,
  format: CellFormat | undefined,
  isNumeric: boolean,
  workbookText = DEFAULT_WORKBOOK_TEXT,
): string {
  if (hasCellFill(format)) {
    return theme.cellFillAutoFg === WORKBOOK_TEXT_COLOR ? workbookText : theme.cellFillAutoFg;
  }
  return isNumeric ? theme.cellNumFg : theme.cellAutoFg;
}

export function paintCellText({
  ctx,
  bounds,
  clip,
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
  merged = false,
  wholePointAdvances = false,
  workbookText,
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
    size: paintedFontSize(baseFont.size, zoomScale, wholePointAdvances),
  };
  const horizontal = cellTextHorizontalMetrics(bounds.w, format?.indent ?? 0, zoomScale);
  const padX = horizontal.padX;
  const indentPx = horizontal.indentPx;
  const fontSize = font.size;
  const paintFontCss = cellFontCss(font);
  const cellSpan = isNumeric
    ? undefined
    : { left: bounds.x + padX, right: alignedTextAnchor(bounds, 'right', padX, 0) };
  // Bottom-aligned runs in a row share the baseline the deepest descent sets.
  const withRowDescent = (box: TextMetricsBox, vAlign: TextVAlign): TextMetricsBox => {
    if (vAlign !== 'bottom' || !rowBaselineFont) return box;
    const rowFont = cellFont(rowBaselineFont, theme, false);
    const current = ctx.font;
    const rowSize = paintedFontSize(rowFont.size, zoomScale, wholePointAdvances);
    ctx.font = cellFontCss({ ...rowFont, size: rowSize });
    const rowBox = fontMetricsBox(ctx, rowSize);
    ctx.font = current;
    const accountingRise = rowBaselineFont.accountingUnderline
      ? rowSize * ACCOUNTING_BASELINE_RISE_EM
      : 0;
    return {
      ascent: box.ascent,
      descent: Math.max(box.descent, rowBox.descent) + accountingRise,
    };
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
            : automaticTextColor(theme, format, isNumeric, workbookText);

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
  const area = clip ?? bounds;
  ctx.rect(area.x, area.y, area.w, area.h);
  ctx.clip();
  const drawGlyphs = applyFontEffects(ctx, fontEffects ? format : undefined, fontSize);
  const measure = (run: string): number => cellTextWidth(ctx, run, wholePointAdvances);
  // Paints a run by `ctx.textAlign` and returns its left edge.
  const drawText = (run: string, x: number, y: number): number =>
    drawCellTextRun(ctx, drawGlyphs, run, x, y, wholePointAdvances);

  if (rotation === 'vertical') {
    paintStackedText(ctx, drawText, text, bounds, {
      format,
      wrap,
      locale,
      fontSize,
      padX,
      padY,
      indentPx,
    });
    ctx.restore();
    return;
  }

  // Rotated text (positive = counterclockwise) is a single line placed by
  // its rotated line box; wrap does not apply.
  if (rotation !== 0) {
    const theta = (rotation * Math.PI) / 180;
    const cos = Math.cos(theta);
    const sin = Math.sin(theta);
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
    const width = measure(text);
    const { ascent, descent } = fontMetricsBox(ctx, fontSize);
    // Line-box corners relative to the run's start on the baseline.
    const xs = [
      -ascent * sin,
      descent * sin,
      width * cos - ascent * sin,
      width * cos + descent * sin,
    ];
    const ys = [
      -ascent * cos,
      descent * cos,
      -width * sin - ascent * cos,
      -width * sin + descent * cos,
    ];
    const minX = Math.min(...xs);
    const maxX = Math.max(...xs);
    const minY = Math.min(...ys);
    const maxY = Math.max(...ys);
    // General: a slanted run sits on the side of its lower end, a vertical
    // one on the side its descent faces.
    const side: CanvasTextAlign = format?.align
      ? align
      : Math.abs(rotation) === 90
        ? rotation > 0
          ? 'right'
          : 'left'
        : rotation > 0
          ? 'left'
          : 'right';
    // Rotated runs keep one more padding width clear of the right edge.
    const rotatedPadX = wholePointAdvances ? wholePointPadX(fontSize) : padX;
    const right = alignedTextAnchor(bounds, 'right', rotatedPadX, indentPx) - rotatedPadX;
    const originX =
      side === 'right'
        ? right - maxX
        : side === 'center'
          ? bounds.x + bounds.w / 2 - (minX + maxX) / 2
          : bounds.x + rotatedPadX + indentPx - minX;
    const vAlign = canvasTextVAlign(format?.vAlign ?? 'bottom');
    // Bottom: the lower end of the baseline sits on the horizontal bottom
    // baseline. Top and middle place the line box.
    const originY =
      vAlign === 'bottom'
        ? textBaselineY(bounds, { ascent, descent }, 'bottom', padY) - Math.max(0, -width * sin)
        : vAlign === 'top'
          ? bounds.y + padY - minY
          : bounds.y + (bounds.h - (maxY - minY)) / 2 - minY;
    ctx.translate(originX, originY);
    ctx.rotate(-theta);
    drawText(text, 0, 0);
    ctx.restore();
    return;
  }

  // Word-wrap path — split into lines, paint each.
  if (wrap) {
    // Whole-point advances round at the painted size, so they wrap there;
    // otherwise lines break at 100% whatever the zoom.
    const lines = wholePointAdvances
      ? wrapText(text, { maxWidth: horizontal.availableWidth, measureText: measure, locale })
      : wrapLogical(ctx, text, baseFont, bounds.w / zoomScale, format?.indent ?? 0, locale);
    ctx.font = paintFontCss;
    // Line pitch rounds in rendered pixels: 14 px for an 11 px run.
    const lineH = cellLinePitch(fontSize);
    const totalH = lineH * lines.length;
    const vAlign = canvasTextVAlign(format?.vAlign ?? 'bottom');
    const measured = withRowDescent(fontMetricsBox(ctx, fontSize), vAlign);
    let startY = textBaselineY(bounds, measured, vAlign, padY, merged, {
      family: font.family,
      size: fontSize,
      scale: zoomScale,
    });
    startY += cellTextVerticalOffset(fontSize, fontVertAlign);
    if (vAlign === 'middle') startY -= (totalH - lineH) / 2;
    else if (vAlign === 'bottom') startY -= (lines.length - 1) * lineH;
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = align;
    const anchor = alignedTextAnchor(
      bounds,
      align,
      wholePointAdvances ? wholePointPadX(fontSize) : padX,
      indentPx,
    );
    for (let i = 0; i < lines.length; i += 1) {
      const line = lines[i] ?? '';
      const baselineY = startY + i * lineH;
      const width = measure(line);
      const lineX0 = drawText(line, anchor, baselineY);
      paintTextDecorations(ctx, format, isHyperlink, lineX0, width, baselineY, fontSize, cellSpan);
    }
    ctx.restore();
    return;
  }

  // Single-line path with vertical alignment. A distributed single word has
  // no gaps to spread and sits centred.
  if (format?.align === 'distributed' && !/\s/.test(text)) align = 'center';
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = align;

  const vAlign = canvasTextVAlign(format?.vAlign ?? 'bottom');
  const availableTextWidthForCell = horizontal.availableWidth;
  let drawFontSize = fontSize;
  let fullSizeDescent = 0;
  if (format?.shrinkToFit) {
    const measuredWidth = measure(text);
    if (measuredWidth > availableTextWidthForCell && availableTextWidthForCell > 0) {
      fullSizeDescent = fontMetricsBox(ctx, fontSize).descent;
      drawFontSize = shrinkFontSizeToFit(fontSize, measuredWidth, availableTextWidthForCell);
      ctx.font = cellFontCss(font, drawFontSize);
      // Rounded advances can leave the scaled run too wide: step down a
      // whole size until it fits (Arial 11 'A very very long shrink to fit
      // sentence' in a 10-char column scales to 3 pt, 61 px, and paints at 2).
      while (wholePointAdvances && drawFontSize > 1 && measure(text) > availableTextWidthForCell) {
        drawFontSize -= 1;
        ctx.font = cellFontCss(font, drawFontSize);
      }
    }
  }
  // Whole-point runs keep a size-dependent inset, from the size they paint at.
  const tx = alignedTextAnchor(
    bounds,
    align,
    wholePointAdvances ? wholePointPadX(drawFontSize) : padX,
    indentPx,
  );
  const generalNumber =
    !isFormulaDisplay &&
    value.kind === 'number' &&
    (!format?.numFmt || format.numFmt.kind === 'general');
  const overflows = generalNumber
    ? generalTextWidth(measure, text) > availableTextWidthForCell
    : measure(text) > availableTextWidthForCell;
  if (isNumeric && overflows) {
    if (format?.numFmt && format.numFmt.kind !== 'general') {
      // Explicit number format: Excel shows #### when the formatted value
      // can't shrink without losing the user's chosen presentation.
      const hashWidth = Math.max(1, measure('#') || fontSize * 0.62);
      text = '#'.repeat(Math.max(1, Math.floor(availableTextWidthForCell / hashWidth)));
    } else if (!isFormulaDisplay && value.kind === 'number') {
      // General number: Excel progressively trims fractional digits before
      // falling back to scientific notation and finally ####. Without this,
      // right-aligned text just clips and the user sees only the trailing
      // digits — making the result look corrupted.
      text = fitGeneralNumberToWidth(
        measure,
        value.value,
        availableTextWidthForCell,
        locale ?? 'en-US',
      );
    }
  }
  // Fill repeats the run as many whole times as the cell holds.
  if (format?.align === 'fill' && !isFormulaDisplay) {
    const runWidth = measure(text);
    if (runWidth > 0) {
      text = text.repeat(Math.max(1, Math.floor(availableTextWidthForCell / runWidth)));
    }
  }
  const metrics = { width: measure(text) };
  const ownBox = fontMetricsBox(ctx, drawFontSize);
  // A shrunk bottom-aligned run keeps the baseline of its full-size font.
  const box = withRowDescent(
    vAlign === 'bottom'
      ? { ...ownBox, descent: Math.max(ownBox.descent, fullSizeDescent) }
      : ownBox,
    vAlign,
  );
  let ty = textBaselineY(bounds, box, vAlign, padY, merged, {
    family: font.family,
    size: fontSize,
    scale: zoomScale,
  });
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
    const leadingWidth = measure(leading);
    const trailingWidth = measure(trailing);
    ctx.textAlign = 'left';
    drawText(leading, left, ty);
    ctx.textAlign = 'right';
    drawText(trailing, right, ty);
    const charWidth = measure(fill.char);
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
      const before = start === 0 ? 0 : measure(text.slice(0, start));
      return { left: textLeft + before, width: measure(text.slice(start, end)) };
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
  const lineX0 = drawText(text, tx, ty);
  paintTextDecorations(ctx, format, isHyperlink, lineX0, metrics.width, ty, drawFontSize, cellSpan);
  ctx.restore();
}
