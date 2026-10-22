/**
 * @file Displayed column width: the figure the Column Width dialog and
 * `Range.ColumnWidth` show. A column `px` screen pixels wide shows
 * `(px - padding) / mdw` characters to two decimals; one narrower than a
 * character plus its padding scales by `mdw + padding` instead. Both desktop
 * platforms follow it, Windows in 96-DPI pixels and Mac in points.
 */

/** Displayed width of a column `px` screen pixels wide. */
export function displayCharsForScreenPx(px: number, mdw: number, padding: number): number {
  if (px <= 0) return 0;
  const chars = px < mdw + padding ? px / (mdw + padding) : (px - padding) / mdw;
  return Math.round(chars * 100) / 100;
}

/** Whole screen pixels of a column set to show `chars` characters. */
export function screenPxForDisplayChars(chars: number, mdw: number, padding: number): number {
  if (chars <= 0) return 0;
  return Math.round(chars < 1 ? chars * (mdw + padding) : chars * mdw + padding);
}
