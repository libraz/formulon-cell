/**
 * Theme colour arithmetic as the engine and the spreadsheet do it: ARGB and CSS
 * hex conversions, and a tint applied to luminance in the integer HLS space, so
 * a tinted theme colour comes out as the spreadsheet paints it.
 */

import type { Color } from '@libraz/formulon';
import { ColorContext, ColorKind } from '@libraz/formulon';
import type { WorkbookHandle } from './workbook-handle.js';

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
  // A file stores a tint as a quantised fraction (0.8 as 0.79998...); the
  // spreadsheet works from the percentage it stands for.
  const nominal = Math.round(tint * 10_000) / 10_000;
  const keep = nominal < 0 ? 1 + nominal : 1 - nominal;
  const scaled = nominal < 0 ? lum * keep : lum * keep + (HLS_MAX - HLS_MAX * keep);
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

/** CSS colour of an engine colour spec through the workbook theme; undefined when the spec is
 *  absent or cannot be resolved. Without a workbook only an RGB spec resolves. */
export const cssOfColorSpec = (
  wb: Pick<WorkbookHandle, 'resolveColorArgb'> | null,
  spec: Color | null,
): string | undefined => {
  if (spec === null || spec.kind === ColorKind.None) return undefined;
  const argb = wb?.resolveColorArgb(spec) ?? (spec.kind === ColorKind.Rgb ? spec.argb : null);
  return argb === null ? undefined : cssColorFromArgb(argb);
};

const hueToChannel = (m1: number, m2: number, hue: number): number => {
  const h = ((hue % 1) + 1) % 1;
  if (h < 1 / 6) return m1 + (m2 - m1) * h * 6;
  if (h < 1 / 2) return m2;
  if (h < 2 / 3) return m1 + (m2 - m1) * (2 / 3 - h) * 6;
  return m1;
};

/** `argb` with a DrawingML `lumMod` / `lumOff` applied in floating-point HSL, each
 *  channel rounded, as charts and shapes are drawn. The engine reads the pair as a tint
 *  (`lumOff` = tint, `lumMod` = 1 - tint; or `lumMod` = 1 + tint for a negative one). */
const luminanceArgb = (argb: number, tint: number): number => {
  // A tint is stored in 1/100000 units, as the file's lumMod / lumOff are.
  const nominal = Math.round(tint * 100_000) / 100_000;
  const [r, g, b] = [16, 8, 0].map((shift) => ((argb >>> shift) & 0xff) / RGB_MAX) as [
    number,
    number,
    number,
  ];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lum = (max + min) / 2;
  let sat = 0;
  let hue = 0;
  if (max !== min) {
    const span = max - min;
    sat = lum <= 0.5 ? span / (max + min) : span / (2 - max - min);
    if (r === max) hue = (g - b) / span;
    else if (g === max) hue = 2 + (b - r) / span;
    else hue = 4 + (r - g) / span;
    hue /= 6;
  }
  const mod = nominal > 0 ? 1 - nominal : 1 + nominal;
  const off = Math.max(nominal, 0);
  const l = Math.min(1, Math.max(0, lum * mod + off));
  const m2 = l <= 0.5 ? l * (1 + sat) : l + sat - l * sat;
  const m1 = 2 * l - m2;
  const channel = (value: number): number => Math.round(value * RGB_MAX);
  const [nr, ng, nb] =
    sat === 0
      ? [l, l, l]
      : [
          hueToChannel(m1, m2, hue + 1 / 3),
          hueToChannel(m1, m2, hue),
          hueToChannel(m1, m2, hue - 1 / 3),
        ];
  return (0xff000000 | (channel(nr) << 16) | (channel(ng) << 8) | channel(nb)) >>> 0;
};

/** `color` as a chart or shape paints it: a theme colour's base resolved through the
 *  workbook theme, its tint applied as DrawingML luminance; any other colour, or one the
 *  engine cannot resolve, unchanged. */
export const renderedThemeColor = (
  wb: Pick<WorkbookHandle, 'resolveColorArgb'>,
  color: Color,
): Color => {
  if (color.kind !== ColorKind.Theme) return color;
  const base = wb.resolveColorArgb({ ...color, tint: 0 }, ColorContext.Font);
  if (base === null) return color;
  const argb = color.tint === 0 ? base : luminanceArgb(base, color.tint);
  return { ...color, kind: ColorKind.Rgb, argb };
};
