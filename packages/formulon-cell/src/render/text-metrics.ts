/**
 * @file Cell-text advance model.
 *
 * The Mac spreadsheet lays a cell's text out on whole points: every glyph
 * advance is rounded to a whole CSS px at the effective size (font points ×
 * zoom), the run's width is the sum, and no pair kerning or sub-pixel
 * accumulation applies. Painting, overflow, wrapping and autofit all measure
 * through {@link cellTextWidth} so a run is laid out the same way everywhere;
 * {@link drawCellTextRun} paints the glyphs at those positions.
 *
 * Only scripts whose glyphs stand alone take the per-glyph model; text that
 * shapes across glyphs (combining marks, joined or right-to-left scripts,
 * emoji sequences) keeps the canvas's own run layout so it is never broken
 * apart. Rounded advances are cached per canvas font and character.
 */

import type { LayoutSlice } from '../store/types.js';

/** Characters drawn one glyph at a time without changing their shape. */
const STANDALONE_GLYPHS = new RegExp(
  '^[' +
    [
      '\\u0020-\\u007e', // ASCII
      '\\u00a0-\\u02ff', // Latin-1, Latin Extended, IPA, spacing modifiers
      '\\u0370-\\u052f', // Greek, Cyrillic
      '\\u1e00-\\u1fff', // Latin and Greek extended
      '\\u2010-\\u2027\\u2030-\\u205e', // punctuation, without bidi and joiner controls
      '\\u20a0-\\u20cf\\u2100-\\u23ff\\u2460-\\u25ff', // currency, letterlike, math, shapes
      '\\u3000-\\u3029\\u3030-\\u303f', // CJK punctuation, without tone marks
      '\\u3041-\\u3096\\u309b-\\u30ff', // kana, without combining sound marks
      '\\u3105-\\u312f\\u3131-\\u318e\\u31f0-\\u31ff\\u3200-\\u33ff', // bopomofo, jamo, CJK symbols
      '\\u3400-\\u4dbf\\u4e00-\\u9fff\\uac00-\\ud7a3\\uf900-\\ufaff', // ideographs, Hangul syllables
      '\\uff01-\\uffef', // half- and full-width forms
    ].join('') +
    ']*$',
);

/** Fonts whose advances are kept before the oldest is dropped. */
const FONT_CACHE_LIMIT = 64;

const advanceCache = new Map<string, Map<string, number>>();

/** Whether cell text lays out on whole-point advances: the Mac point model,
 *  whose screen scale maps CSS px to points. */
export const wholePointAdvances = (layout: Pick<LayoutSlice, 'screenScale'>): boolean =>
  (layout.screenScale ?? 1) !== 1;

const fontAdvances = (font: string): Map<string, number> => {
  let advances = advanceCache.get(font);
  if (advances) return advances;
  advances = new Map();
  advanceCache.set(font, advances);
  if (advanceCache.size > FONT_CACHE_LIMIT) {
    const oldest = advanceCache.keys().next().value;
    if (oldest !== undefined) advanceCache.delete(oldest);
  }
  return advances;
};

/** Whole-px advance of every character of `text` at the current `ctx.font`,
 *  or null when the run must keep the canvas's own layout. */
export const glyphAdvances = (ctx: CanvasRenderingContext2D, text: string): number[] | null => {
  if (ctx.direction === 'rtl' || !STANDALONE_GLYPHS.test(text)) return null;
  const cache = fontAdvances(ctx.font);
  const out: number[] = [];
  for (const char of text) {
    let advance = cache.get(char);
    if (advance === undefined) {
      advance = Math.round(ctx.measureText(char).width);
      cache.set(char, advance);
    }
    out.push(advance);
  }
  return out;
};

/** Width of a single-line run at the current `ctx.font`. */
export const cellTextWidth = (
  ctx: CanvasRenderingContext2D,
  text: string,
  wholePoint: boolean,
): number => {
  const advances = wholePoint ? glyphAdvances(ctx, text) : null;
  if (!advances) return ctx.measureText(text).width;
  let width = 0;
  for (const advance of advances) width += advance;
  return width;
};

/** Painted size of a run: on whole-point advances the zoomed size snaps to a
 *  whole px too (Arial 11 pt paints at 8, 14, 17 and 22 px at 75, 125, 150
 *  and 200%); otherwise the exact product. */
/** The line pitch of a cell's text at CSS font size `size`, shared by painting and autofit. */
export const cellLinePitch = (size: number): number => Math.round(size * 1.28);

export const paintedFontSize = (size: number, scale: number, wholePoint: boolean): number =>
  wholePoint ? Math.max(1, Math.round(size * scale)) : size * scale;

/** Effective sizes (CSS px) up to which a single-line run keeps 1 and 2 px
 *  clear of the cell edge; larger runs keep 3. Measured on Arial 2-22 px
 *  ('Hxg' ladder 8-20 pt, shrink-to-fit runs, 11 pt at 75-200% zoom). */
const INSET_STEPS = [8.5, 15] as const;

/** Horizontal inset of a single-line whole-point run at `fontSize` (CSS px,
 *  zoom applied), in the `padX` form `alignedTextAnchor` takes: the whole px
 *  kept clear plus the quarter px the right anchor's gridline half and floor
 *  leave, so both edges keep the same whole px clear. */
export const wholePointPadX = (fontSize: number): number =>
  1 + INSET_STEPS.filter((step) => fontSize >= step).length + 0.25;

/** The whole px a run starts on: the one at or before its laid-out start, as
 *  a centred run with a half-px slack sits on the spreadsheet. */
export const wholePxStart = (x: number): number => Math.floor(x + 1e-6);

/**
 * Paint a single-line run anchored at `x` by the current `ctx.textAlign`
 * (left, center or right) through `draw`, and return the run's left edge.
 * On whole-point advances the run starts on a whole px and each glyph sits at
 * the sum of the rounded advances before it.
 */
export const drawCellTextRun = (
  ctx: CanvasRenderingContext2D,
  draw: (text: string, x: number, y: number) => void,
  text: string,
  x: number,
  y: number,
  wholePoint: boolean,
): number => {
  const align = ctx.textAlign;
  const advances = wholePoint ? glyphAdvances(ctx, text) : null;
  if (!advances) {
    draw(text, x, y);
    const width = ctx.measureText(text).width;
    return align === 'right' ? x - width : align === 'center' ? x - width / 2 : x;
  }
  let width = 0;
  for (const advance of advances) width += advance;
  const left = wholePxStart(align === 'right' ? x - width : align === 'center' ? x - width / 2 : x);
  ctx.textAlign = 'left';
  let pen = left;
  let i = 0;
  for (const char of text) {
    draw(char, pen, y);
    pen += advances[i] ?? 0;
    i += 1;
  }
  ctx.textAlign = align;
  return left;
};
