// Minimal per-paint context shared by the free chrome painters (gridlines,
// headers, freeze dividers, traces). Bundles the canvas state that the host
// `GridRenderer` would otherwise read off `this.ctx` / `this.dpr` etc.

export interface ChromePaintContext {
  ctx: CanvasRenderingContext2D;
  dpr: number;
  cssWidth: number;
  cssHeight: number;
}

/** Device pixels in one rule: one CSS pixel snapped to whole device pixels. */
export const rulePx = (dpr: number): number => Math.max(1, Math.round(dpr));

/** A one-CSS-pixel rule, snapped to whole device pixels, as the spreadsheet
 *  draws gridlines at every display scale. */
export const ruleWidth = (dpr: number): number => rulePx(dpr) / dpr;

/** CSS centre of the rule a boundary owns. The spreadsheet draws each
 *  gridline on the last device pixels before the boundary — after it when
 *  `mirrored` (RTL columns), where the trailing edge is on the left. */
export const ruleCenter = (boundary: number, dpr: number, mirrored = false): number => {
  const g = rulePx(dpr);
  const edge = Math.round(boundary * dpr);
  return (mirrored ? edge + g / 2 : edge - g / 2) / dpr;
};

interface PaintRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** A cell fill's rectangle: the cell snapped to device pixels, grown by one
 *  rule over the gridline before it, so a filled cell hides both of its
 *  gridlines as in the spreadsheet. */
export const fillRect = (bounds: PaintRect, dpr: number, mirrored = false): PaintRect => {
  const g = rulePx(dpr);
  const l = Math.round(bounds.x * dpr);
  const r = Math.round((bounds.x + bounds.w) * dpr);
  const t = Math.round(bounds.y * dpr);
  const b = Math.round((bounds.y + bounds.h) * dpr);
  return {
    x: (mirrored ? l : l - g) / dpr,
    y: (t - g) / dpr,
    w: (r - l + g) / dpr,
    h: (b - t + g) / dpr,
  };
};

/** A cell's surface between its gridlines: the device-snapped cell minus the
 *  rule it owns on its trailing edges. */
export const surfaceRect = (bounds: PaintRect, dpr: number, mirrored = false): PaintRect => {
  const g = rulePx(dpr);
  const l = Math.round(bounds.x * dpr);
  const r = Math.round((bounds.x + bounds.w) * dpr);
  const t = Math.round(bounds.y * dpr);
  const b = Math.round((bounds.y + bounds.h) * dpr);
  return {
    x: (mirrored ? l + g : l) / dpr,
    y: t / dpr,
    w: Math.max(0, r - l - g) / dpr,
    h: Math.max(0, b - t - g) / dpr,
  };
};
