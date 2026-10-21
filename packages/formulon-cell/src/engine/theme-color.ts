/**
 * Theme colour arithmetic as the engine and the spreadsheet do it: ARGB and CSS
 * hex conversions, and a tint applied to luminance in the integer HLS space, so
 * a tinted theme colour comes out as the spreadsheet paints it.
 */

const CSS_HEX_COLOR = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i;

/** `#rrggbb` from an engine ARGB colour; the alpha byte is carried by transparency. */
export const cssColorFromArgb = (argb: number): string =>
  `#${(argb & 0xffffff).toString(16).padStart(6, '0')}`;

/** Opaque ARGB from `#rgb` / `#rrggbb`; null for any other CSS colour. */
export const argbFromCssColor = (color: string): number | null => {
  const digits = color.trim().match(CSS_HEX_COLOR)?.[1];
  if (digits === undefined) return null;
  const hex =
    digits.length === 3 ? [...digits].map((digit) => `${digit}${digit}`).join('') : digits;
  return (0xff000000 | Number.parseInt(hex, 16)) >>> 0;
};

const HLS_MAX = 240;
const RGB_MAX = 255;

const hueToRgb = (n1: number, n2: number, hue: number): number => {
  const h = ((hue % HLS_MAX) + HLS_MAX) % HLS_MAX;
  if (h < HLS_MAX / 6) return n1 + Math.trunc(((n2 - n1) * h + HLS_MAX / 12) / (HLS_MAX / 6));
  if (h < HLS_MAX / 2) return n2;
  if (h < (HLS_MAX * 2) / 3) {
    return n1 + Math.trunc(((n2 - n1) * ((HLS_MAX * 2) / 3 - h) + HLS_MAX / 12) / (HLS_MAX / 6));
  }
  return n1;
};

/** `argb` with an Excel tint (-1..1) applied to its luminance, in the engine's integer HLS. */
const tintedArgb = (argb: number, tint: number): number => {
  if (tint === 0) return (argb | 0xff000000) >>> 0;
  const r = (argb >>> 16) & 0xff;
  const g = (argb >>> 8) & 0xff;
  const b = argb & 0xff;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lum = Math.trunc(((max + min) * HLS_MAX + RGB_MAX) / (2 * RGB_MAX));
  let sat = 0;
  let hue = 0;
  if (max !== min) {
    sat =
      lum <= HLS_MAX / 2
        ? Math.trunc(((max - min) * HLS_MAX + (max + min) / 2) / (max + min))
        : Math.trunc(
            ((max - min) * HLS_MAX + (2 * RGB_MAX - max - min) / 2) / (2 * RGB_MAX - max - min),
          );
    const span = max - min;
    const delta = (c: number): number => Math.trunc(((max - c) * (HLS_MAX / 6) + span / 2) / span);
    if (r === max) hue = delta(b) - delta(g);
    else if (g === max) hue = HLS_MAX / 3 + delta(r) - delta(b);
    else hue = (2 * HLS_MAX) / 3 + delta(g) - delta(r);
    if (hue < 0) hue += HLS_MAX;
    if (hue > HLS_MAX) hue -= HLS_MAX;
  }
  const keep = tint < 0 ? 1 + tint : 1 - tint;
  const scaled = tint < 0 ? lum * keep : lum * keep + (HLS_MAX - HLS_MAX * keep);
  const newLum = Math.min(HLS_MAX, Math.max(0, Math.floor(scaled)));
  const pack = (red: number, green: number, blue: number): number =>
    (0xff000000 | (red << 16) | (green << 8) | blue) >>> 0;
  if (sat === 0) {
    const v = Math.trunc((newLum * RGB_MAX + HLS_MAX / 2) / HLS_MAX);
    return pack(v, v, v);
  }
  const magic2 =
    newLum <= HLS_MAX / 2
      ? Math.trunc((newLum * (HLS_MAX + sat) + HLS_MAX / 2) / HLS_MAX)
      : newLum + sat - Math.trunc((newLum * sat + HLS_MAX / 2) / HLS_MAX);
  const magic1 = 2 * newLum - magic2;
  const channel = (h: number): number =>
    Math.trunc((hueToRgb(magic1, magic2, h) * RGB_MAX + HLS_MAX / 2) / HLS_MAX);
  return pack(channel(hue + HLS_MAX / 3), channel(hue), channel(hue - HLS_MAX / 3));
};

/** `color` (`#rrggbb`) with an Excel tint (-1..1) applied; null for any other CSS colour. */
export const tintCssColor = (color: string, tint: number): string | null => {
  const argb = argbFromCssColor(color);
  return argb === null ? null : cssColorFromArgb(tintedArgb(argb, tint));
};
