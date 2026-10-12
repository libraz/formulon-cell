/**
 * Font sizes have two public representations in formulon-cell:
 *
 * - `CellFormat.fontSize` is CSS pixels, used by canvas and DOM painting.
 * - Formulon's `FontRecord.size` is typographic points, as stored by XLSX.
 *
 * Keep the conversion at the engine boundary.  Neither side should round the
 * value: Excel accepts fractional point sizes and a repeated save/load cycle
 * must not slowly inflate or shrink a font.
 */
const CSS_PIXELS_PER_POINT = 96 / 72;
const POINTS_PER_CSS_PIXEL = 72 / 96;

/** Convert a finite native point size to CSS pixels. */
export function fontPointsToCssPx(points: number): number {
  return Number.isFinite(points) ? points * CSS_PIXELS_PER_POINT : 0;
}

/** Convert a finite public CSS-pixel size to native typographic points. */
export function fontCssPxToPoints(px: number): number {
  return Number.isFinite(px) ? px * POINTS_PER_CSS_PIXEL : 0;
}

/** Stabilize the UI projection only. Engine conversions retain all precision,
 * while named size options must recognize 14pt after a px round trip. */
export function fontCssPxToDisplayPoints(px: number): number {
  return Number(fontCssPxToPoints(px).toPrecision(14));
}
